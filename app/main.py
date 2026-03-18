"""FastAPI web application for PDF-to-PPTX conversion."""

import os
import uuid
import asyncio
import threading
from pathlib import Path

from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from starlette.responses import StreamingResponse

from app.converter import convert_pdf_to_pptx
from app.pptx_translator import translate_pptx

BASE_DIR = Path(__file__).resolve().parent.parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"
STATIC_DIR = BASE_DIR / "static"

UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

app = FastAPI(title="PDF to PPTX Converter")

# Store job progress
jobs = {}


@app.get("/", response_class=HTMLResponse)
async def index():
    index_file = STATIC_DIR / "index.html"
    return HTMLResponse(content=index_file.read_text())


@app.post("/api/convert")
async def convert(file: UploadFile = File(...)):
    """Upload a PDF or PPTX and start conversion."""
    job_id = str(uuid.uuid4())
    original_name = Path(file.filename).stem if file.filename else "output"
    ext = Path(file.filename).suffix.lower() if file.filename else ""

    if ext not in (".pdf", ".pptx"):
        return {"error": "Unsupported file type. Please upload a PDF or PPTX file."}

    # Save uploaded file
    input_path = UPLOAD_DIR / f"{job_id}{ext}"
    content = await file.read()
    with open(input_path, "wb") as f:
        f.write(content)

    output_path = OUTPUT_DIR / f"{job_id}.pptx"

    jobs[job_id] = {
        "status": "processing",
        "progress": [],
        "original_name": original_name,
        "output_path": str(output_path),
        "done": False,
        "error": None,
        "mode": "pptx_translate" if ext == ".pptx" else "pdf_convert",
    }

    # Run conversion in background thread
    thread = threading.Thread(
        target=_run_conversion,
        args=(job_id, str(input_path), str(output_path), ext),
    )
    thread.start()

    return {"job_id": job_id}


def _run_conversion(job_id, input_path, output_path, ext):
    """Run the conversion pipeline in a background thread."""
    def progress_callback(message):
        if job_id in jobs:
            jobs[job_id]["progress"].append(message)

    try:
        if ext == ".pptx":
            translate_pptx(input_path, output_path, progress_callback=progress_callback)
        else:
            convert_pdf_to_pptx(input_path, output_path, progress_callback=progress_callback)
        jobs[job_id]["status"] = "complete"
        jobs[job_id]["done"] = True
    except Exception as e:
        jobs[job_id]["status"] = "error"
        jobs[job_id]["error"] = str(e)
        jobs[job_id]["done"] = True
    finally:
        # Clean up uploaded file
        try:
            os.remove(input_path)
        except OSError:
            pass


@app.get("/api/status/{job_id}")
async def status_stream(job_id: str):
    """SSE endpoint for conversion progress."""
    if job_id not in jobs:
        return {"error": "Job not found"}

    async def event_generator():
        last_index = 0
        while True:
            job = jobs.get(job_id)
            if not job:
                break

            # Send new progress messages
            while last_index < len(job["progress"]):
                msg = job["progress"][last_index]
                yield f"data: {msg}\n\n"
                last_index += 1

            if job["done"]:
                if job["status"] == "complete":
                    yield f"data: DONE\n\n"
                else:
                    yield f"data: ERROR: {job.get('error', 'Unknown error')}\n\n"
                break

            await asyncio.sleep(0.5)

    return StreamingResponse(
        event_generator(),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "Connection": "keep-alive",
        },
    )


@app.get("/api/download/{job_id}")
async def download(job_id: str):
    """Download the converted PPTX file."""
    job = jobs.get(job_id)
    if not job:
        return {"error": "Job not found"}

    if job["status"] != "complete":
        return {"error": "Conversion not complete"}

    output_path = job["output_path"]
    if not os.path.exists(output_path):
        return {"error": "Output file not found"}

    filename = f"{job['original_name']}.pptx"
    return FileResponse(
        output_path,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        filename=filename,
    )
