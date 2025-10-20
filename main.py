from fastapi import FastAPI, UploadFile, Form, File
from fastapi.responses import JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
import shutil
import os

# Import both main route function and new summary function
from logic.route_logic import process_route, get_day_summary

# =====================================================
# Initialize FastAPI app
# =====================================================
app = FastAPI(title="Route Planner API")

# =====================================================
# Enable CORS (for React frontend)
# =====================================================
app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "http://localhost:5173",
        "http://127.0.0.1:5173",
        "http://localhost:5174",
        "http://127.0.0.1:5174",
    ],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# =====================================================
# Define and create directories
# =====================================================
UPLOAD_DIR = "uploads"
OUTPUT_DIR = "output"
MAPS_DIR = "maps"

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(MAPS_DIR, exist_ok=True)

# =====================================================
# Serve static files (for frontend access)
# =====================================================
app.mount("/output", StaticFiles(directory=OUTPUT_DIR), name="output")
app.mount("/maps", StaticFiles(directory=MAPS_DIR), name="maps")

# =====================================================
# Root endpoint
# =====================================================
@app.get("/")
def home():
    return {"message": "✅ Route Planner API is running and ready!"}


# =====================================================
# Route generation endpoint
# =====================================================
@app.post("/generate-route/")
async def generate_route(
    data_file: UploadFile = File(...),       # input_file_2.xlsx (data file)
    format_file: UploadFile = File(...),     # Route Plan Format.xlsx (template)
    so_name: str = Form(...),
    so_erp: str = Form(...),
    week: int = Form(...),
    day: str = Form(...)
):
    try:
        # ------------------------------------------------
        # Save uploaded Excel files to the uploads folder
        # ------------------------------------------------
        input_data_path = os.path.join(UPLOAD_DIR, data_file.filename)
        format_data_path = os.path.join(UPLOAD_DIR, format_file.filename)

        with open(input_data_path, "wb") as f:
            shutil.copyfileobj(data_file.file, f)
        with open(format_data_path, "wb") as f:
            shutil.copyfileobj(format_file.file, f)

        print("\n📂 Files received:")
        print(f"   - Data file: {data_file.filename}")
        print(f"   - Format file: {format_file.filename}")
        print(f"👤 SO: {so_name} | ERP: {so_erp} | Week: {week} | Day: {day}")

        # ------------------------------------------------
        # Process route logic
        # ------------------------------------------------
        result = process_route(
            input_path=input_data_path,
            format_path=format_data_path,
            so_name=so_name,
            so_erp=so_erp,
            week=week,
            day=day
        )

        # ------------------------------------------------
        # Error handling
        # ------------------------------------------------
        if result.get("status") == "error":
            print(f"❌ Error during processing: {result.get('message')}")
            return JSONResponse(content=result, status_code=400)

        # ------------------------------------------------
        # Success response
        # ------------------------------------------------
        print("✅ Route generated successfully!")
        return JSONResponse(content=result, status_code=200)

    except Exception as e:
        print(f"❌ Internal Server Error: {str(e)}")
        return JSONResponse(
            content={"status": "error", "message": str(e)},
            status_code=500
        )


# =====================================================
# NEW: Daily summary endpoint
# =====================================================
@app.get("/day-summary/")
async def day_summary(so_name: str, week: int, day: str):
    """
    Returns optimized route summary (map + total distance + visit order list)
    for a given SO, week, and day.
    """
    try:
        print(f"📅 Fetching summary for {so_name} - Week {week}, {day}")
        result = get_day_summary(so_name, week, day)

        if result.get("status") == "error":
            print(f"❌ Error: {result.get('message')}")
            return JSONResponse(content=result, status_code=400)

        print(f"✅ Day summary generated successfully for {so_name} ({day})")
        return JSONResponse(content=result, status_code=200)

    except Exception as e:
        print(f"❌ Internal Server Error in /day-summary/: {str(e)}")
        return JSONResponse(
            content={"status": "error", "message": str(e)},
            status_code=500
        )
