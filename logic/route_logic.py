import pandas as pd
import re
from collections import defaultdict
import os
import osmnx as ox
import networkx as nx
import folium

# ======================
# CONFIGURATION
# ======================
BASE_DIR = os.getcwd()
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
MAPS_DIR = os.path.join(BASE_DIR, "maps")

os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(MAPS_DIR, exist_ok=True)

PER_DAY_CAPACITY = 45
DAYS = ["MON", "TUE", "WED", "THU", "FRI", "SAT"]
WEEKS = [1, 2, 3, 4]


# ======================
# HELPERS
# ======================
def normalize_input(df):
    """Standardize column names & ensure correct ID mapping."""
    return df.rename(columns={
        'User ': 'SO_NAME',
        'User Erp Id': 'SO_ERP_ID',
        'Visit Count': 'Frequency',
        'Beats Name': 'Beat',
        'Outlet Erp Id': 'Outlet_Erp_Id',
        'Outlets Name': 'Outlet_Name',
        'Latitude': 'Latitude',
        'Longitude': 'Longitude',
        'DAY': 'Preferred_DAY'
    })


def parse_so(name):
    if pd.isna(name):
        return ("UNK", "")
    parts = str(name).strip().split()
    first = parts[0] if parts else "UNK"
    last = parts[-1] if len(parts) > 1 else ""
    return (first.upper(), last.upper())


def parse_freq(val):
    """Parse frequency values (e.g., '2W' -> 2)."""
    if pd.isna(val):
        return 1
    s = str(val)
    m = re.search(r'(\d+)', s)
    if m:
        return int(m.group(1))
    try:
        return int(s)
    except Exception:
        return 1


def pick_weeks(freq):
    """Decide which weeks a beat is visited based on frequency."""
    if freq == 1:
        return None
    if freq == 2:
        return [1, 3]
    if freq == 3:
        return [1, 2, 4]
    return [1, 2, 3, 4]


# ======================
# ROUTE GENERATION
# ======================
def generate_route_plan(input_path, so_name, so_erp):
    """Generate route plan based on input file and SO details."""
    raw = pd.read_excel(input_path)
    df = normalize_input(raw)

    df = df[
        (df['SO_NAME'].str.strip().str.lower() == so_name.lower()) &
        (df['SO_ERP_ID'].astype(str).str.strip() == so_erp)
    ]

    if df.empty:
        print(f"❌ No records found for {so_name} ({so_erp}).")
        return None

    so_day_count = defaultdict(int)
    records = []
    first, last = parse_so(so_name)
    week_load = {w: 0 for w in WEEKS}

    for _, row in df.iterrows():
        freq = parse_freq(row.get('Frequency', 1))
        outlet_id = row.get('Outlet_Erp_Id')
        beat = row.get('Beat')
        preferred = row.get('Preferred_DAY', None)
        lat, lon = row.get('Latitude'), row.get('Longitude')

        weeks_assigned = pick_weeks(freq)
        if weeks_assigned is None:
            target_week = min(week_load, key=week_load.get)
            weeks_assigned = [target_week]

        for wk in weeks_assigned:
            if isinstance(preferred, str) and preferred.strip():
                day = preferred.strip()[:3].upper()
                if day not in DAYS:
                    day = DAYS[(wk - 1) % len(DAYS)]
            else:
                day_counts = {d: so_day_count[(so_name, wk, d)] for d in DAYS}
                day = min(day_counts, key=day_counts.get)

            if so_day_count[(so_name, wk, day)] >= PER_DAY_CAPACITY:
                for d in DAYS:
                    if so_day_count[(so_name, wk, d)] < PER_DAY_CAPACITY:
                        day = d
                        break

            so_day_count[(so_name, wk, day)] += 1
            week_load[wk] += 1

            rid = f"{wk}{first[:3]}{last[:1]}_{str(beat).replace(' ', '').upper()}_W{wk}_{day}"
            records.append({
                'SO NAME': so_name,
                'SO_ERP_ID': so_erp,
                'BEAT NAME': beat,
                'ROUTE NAME': rid,
                'ROUTE ERP ID': rid,
                'Outlet_Erp_Id': outlet_id,
                'Outlet_Name': row.get('Outlet_Name'),
                'Latitude': lat,
                'Longitude': lon,
                'WEEK': wk,
                'DAY': day,
                'VISIT_ORDER': so_day_count[(so_name, wk, day)]
            })

    route_df = pd.DataFrame(records)
    print(f"✅ Route plan generated for {so_name}. Rows: {len(route_df)}")
    return route_df


# ======================
# ROUTE OPTIMIZATION
# ======================
def optimize_daily_route(df_day):
    """Optimize route order for one day's stores using OSM road data."""
    if df_day[['Latitude', 'Longitude']].isna().any().any() or len(df_day) < 2:
        df_day = df_day.reset_index(drop=True)
        df_day['VISIT_ORDER'] = df_day.index + 1
        return df_day, None, None, 0.0

    G = ox.graph_from_point(
        (df_day['Latitude'].mean(), df_day['Longitude'].mean()),
        dist=7000, network_type='drive'
    )

    points = list(zip(df_day['Latitude'], df_day['Longitude']))
    nodes = [ox.distance.nearest_nodes(G, lon, lat) for lat, lon in points]

    start = nodes[0]
    route = [start]
    unvisited = set(nodes[1:])
    total_length = 0.0

    while unvisited:
        current = route[-1]
        nearest = min(unvisited, key=lambda n: nx.shortest_path_length(G, current, n, weight='length'))
        try:
            path_len = nx.shortest_path_length(G, current, nearest, weight='length')
            total_length += path_len
        except Exception:
            pass
        unvisited.remove(nearest)
        route.append(nearest)

    order = [nodes.index(n) for n in route if n in nodes]
    df_day = df_day.iloc[order].reset_index(drop=True)
    df_day['VISIT_ORDER'] = df_day.index + 1

    return df_day, route, G, (total_length / 1000.0)


# ======================
# MAPPING
# ======================
def visualize_route(df_day, route, G, so_name, week, day):
    """Save daily route map as an interactive HTML file and return path + computed km."""
    if route is None or G is None:
        try:
            m = folium.Map(location=[df_day['Latitude'].mean(), df_day['Longitude'].mean()], zoom_start=12)
            for _, r in df_day.iterrows():
                folium.Marker([r['Latitude'], r['Longitude']],
                              popup=f"{r['Outlet_Name']} (#{r['VISIT_ORDER']})").add_to(m)
            map_file = os.path.join(MAPS_DIR, f"{so_name.replace(' ', '_')}_W{week}_{day}.html")
            m.save(map_file)
            return map_file, 0.0
        except Exception as e:
            print(f"⚠️ Could not write fallback map: {e}")
            return None, 0.0

    m = folium.Map(location=[df_day['Latitude'].mean(), df_day['Longitude'].mean()], zoom_start=12)
    total_length = 0.0

    def get_edge_length(G, path):
        edges = ox.routing.route_to_gdf(G, path, weight='length')
        return edges['length'].sum()

    for i in range(len(route) - 1):
        try:
            path = nx.shortest_path(G, route[i], route[i + 1], weight="length")
            coords = [(G.nodes[n]['y'], G.nodes[n]['x']) for n in path]
            folium.PolyLine(coords, color='blue', weight=4, opacity=0.8).add_to(m)
            total_length += get_edge_length(G, path)
        except Exception as e:
            print(f"⚠️ Segment {i} skipped: {e}")

    for _, r in df_day.iterrows():
        folium.Marker([r['Latitude'], r['Longitude']],
                      popup=f"{r['Outlet_Name']} (#{r['VISIT_ORDER']})").add_to(m)

    total_km = total_length / 1000.0
    map_file = os.path.join(MAPS_DIR, f"{so_name.replace(' ', '_')}_W{week}_{day}.html")
    try:
        m.save(map_file)
    except Exception as e:
        print(f"❌ Failed saving map: {e}")
        return None, total_km

    print(f"🗺️ Saved map: {map_file}")
    print(f"📏 Total distance: {total_km:.2f} km")
    return map_file, total_km


# ======================
# MAIN ENTRY
# ======================
def process_route(input_path, format_path, so_name, so_erp, week, day):
    try:
        route_df = generate_route_plan(input_path, so_name, so_erp)
        if route_df is None:
            return {"status": "error", "message": f"No records for {so_name} ({so_erp})"}

        df_day = route_df[(route_df['WEEK'] == week) & (route_df['DAY'] == day)]
        if df_day.empty:
            return {"status": "error", "message": f"No data for Week {week}, {day}"}

        df_day_opt, route, G, opt_km = optimize_daily_route(df_day)
        map_file, map_km = visualize_route(df_day_opt, route, G, so_name, week, day)

        distance_km = opt_km or map_km or 0.0

        required_cols = [
            "SO NAME", "SO_ERP_ID", "BEAT NAME", "ROUTE NAME", "ROUTE ERP ID",
            "Outlet_Erp_Id", "Outlet_Name", "Latitude", "Longitude",
            "WEEK", "DAY", "VISIT_ORDER"
        ]
        existing = [c for c in required_cols if c in route_df.columns]
        clean_df = route_df[existing].copy()
        clean_df = clean_df.sort_values(by=["WEEK", "DAY", "VISIT_ORDER"], ascending=True)

        output_file = os.path.join(OUTPUT_DIR, f"generated_route_{so_name.replace(' ', '_')}.xlsx")
        clean_df.to_excel(output_file, index=False)

        excel_url = f"http://127.0.0.1:8000/output/{os.path.basename(output_file)}" if os.path.exists(output_file) else None
        map_url = f"http://127.0.0.1:8000/maps/{os.path.basename(map_file)}" if map_file and os.path.exists(map_file) else None

        return {
            "status": "success",
            "message": f"✅ Route generated successfully for {so_name}",
            "excel_file": excel_url,
            "map_file": map_url,
            "distance_km": round(distance_km, 2)
        }

    except Exception as e:
        print(f"❌ Error in process_route: {e}")
        return {"status": "error", "message": str(e)}


# ======================
# NEW: DAILY SUMMARY FUNCTION
# ======================
def get_day_summary(so_name, week, day):
    """
    Reads generated Excel and returns total km + visit list + regenerates map.
    """
    try:
        excel_path = os.path.join(OUTPUT_DIR, f"generated_route_{so_name.replace(' ', '_')}.xlsx")
        if not os.path.exists(excel_path):
            return {"status": "error", "message": "Generated Excel not found. Please generate route first."}

        df = pd.read_excel(excel_path)
        df_day = df[(df["WEEK"] == int(week)) & (df["DAY"] == day)]
        if df_day.empty:
            return {"status": "error", "message": f"No data for Week {week}, {day}"}

        df_day_opt, route, G, total_km = optimize_daily_route(df_day)
        map_file, map_km = visualize_route(df_day_opt, route, G, so_name, week, day)
        distance_km = total_km or map_km or 0.0

        visit_list = df_day_opt[[
            "VISIT_ORDER", "Outlet_Erp_Id", "Outlet_Name", "Latitude", "Longitude"
        ]].sort_values(by="VISIT_ORDER").to_dict(orient="records")

        map_url = f"http://127.0.0.1:8000/maps/{os.path.basename(map_file)}" if map_file and os.path.exists(map_file) else None

        return {
            "status": "success",
            "message": f"Summary for {so_name} - Week {week} {day}",
            "total_distance_km": round(distance_km, 2),
            "total_outlets": len(visit_list),
            "visit_list": visit_list,
            "map_file": map_url
        }

    except Exception as e:
        print(f"❌ Error in get_day_summary: {e}")
        return {"status": "error", "message": str(e)}
