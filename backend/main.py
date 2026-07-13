"""
main.py - FastAPI backend for a Word<->PDF converter website (iLovePDF-style).

Endpoints:
  POST /api/to-pdf       - upload a .docx, get back a .pdf
  POST /api/to-word      - upload a .pdf,  get back a .docx
  POST /api/to-excel     - upload a .pdf,  get back a .xlsx
  POST /api/from-excel   - upload a .xlsx, get back a .pdf
  POST /api/merge-pdf    - upload 2+ .pdf files (field "files"), get back one merged .pdf
  POST /api/split-pdf    - upload a .pdf + start_page/end_page form fields, get back that page range as a .pdf
  POST /api/compress-pdf - upload a .pdf, get back a smaller .pdf (images recompressed)
  GET  /api/health       - health check

Files are saved to a temp job folder, converted, streamed back to the
user, and deleted afterwards (via background task) so nothing piles up
on disk.
"""

import os
import uuid
import shutil
from pathlib import Path
from typing import List

from fastapi import FastAPI, UploadFile, File, Form, HTTPException, BackgroundTasks
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse

from converter import (
    word_to_pdf, pdf_to_word, excel_to_pdf, pdf_to_excel,
    merge_pdfs, split_pdf, compress_pdf,
)

APP_TMP_DIR = Path(os.environ.get("CONVERTER_TMP_DIR", "/tmp/converter_jobs"))
APP_TMP_DIR.mkdir(parents=True, exist_ok=True)

MAX_FILE_SIZE_MB = 25  # keep uploads reasonable; raise if you need bigger files

app = FastAPI(title="Word/PDF Converter API")

# Allow your frontend domain(s) to call this API.
# In production, replace "*" with your actual website domain(s).
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


def cleanup_job_dir(job_dir: Path):
    shutil.rmtree(job_dir, ignore_errors=True)


def make_job_dir() -> Path:
    job_dir = APP_TMP_DIR / uuid.uuid4().hex
    job_dir.mkdir(parents=True, exist_ok=True)
    return job_dir


async def save_upload(upload: UploadFile, dest: Path):
    size = 0
    with open(dest, "wb") as f:
        while chunk := await upload.read(1024 * 1024):
            size += len(chunk)
            if size > MAX_FILE_SIZE_MB * 1024 * 1024:
                raise HTTPException(413, f"File exceeds {MAX_FILE_SIZE_MB}MB limit.")
            f.write(chunk)


@app.get("/api/health")
def health():
    return {"status": "ok"}


@app.post("/api/to-pdf")
async def convert_to_pdf(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    if not file.filename.lower().endswith(".docx"):
        raise HTTPException(400, "Please upload a .docx file.")

    job_dir = make_job_dir()
    input_path = job_dir / "input.docx"
    output_path = job_dir / "output.pdf"

    await save_upload(file, input_path)

    try:
        word_to_pdf(str(input_path), str(output_path))
    except Exception as e:
        cleanup_job_dir(job_dir)
        raise HTTPException(500, f"Conversion failed: {e}")

    background_tasks.add_task(cleanup_job_dir, job_dir)
    download_name = os.path.splitext(file.filename)[0] + ".pdf"
    return FileResponse(
        output_path, media_type="application/pdf", filename=download_name,
        background=background_tasks,
    )


@app.post("/api/to-word")
async def convert_to_word(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(400, "Please upload a .pdf file.")

    job_dir = make_job_dir()
    input_path = job_dir / "input.pdf"
    output_path = job_dir / "output.docx"

    await save_upload(file, input_path)

    try:
        pdf_to_word(str(input_path), str(output_path))
    except Exception as e:
        cleanup_job_dir(job_dir)
        raise HTTPException(500, f"Conversion failed: {e}")

    background_tasks.add_task(cleanup_job_dir, job_dir)
    download_name = os.path.splitext(file.filename)[0] + ".docx"
    return FileResponse(
        output_path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=download_name,
        background=background_tasks,
    )


@app.post("/api/to-excel")
async def convert_to_excel(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(400, "Please upload a .pdf file.")

    job_dir = make_job_dir()
    input_path = job_dir / "input.pdf"
    output_path = job_dir / "output.xlsx"

    await save_upload(file, input_path)

    try:
        pdf_to_excel(str(input_path), str(output_path))
    except Exception as e:
        cleanup_job_dir(job_dir)
        raise HTTPException(500, f"Conversion failed: {e}")

    background_tasks.add_task(cleanup_job_dir, job_dir)
    download_name = os.path.splitext(file.filename)[0] + ".xlsx"
    return FileResponse(
        output_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=download_name,
        background=background_tasks,
    )


@app.post("/api/from-excel")
async def convert_from_excel(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    if not file.filename.lower().endswith(".xlsx"):
        raise HTTPException(400, "Please upload a .xlsx file.")

    job_dir = make_job_dir()
    input_path = job_dir / "input.xlsx"
    output_path = job_dir / "output.pdf"

    await save_upload(file, input_path)

    try:
        excel_to_pdf(str(input_path), str(output_path))
    except Exception as e:
        cleanup_job_dir(job_dir)
        raise HTTPException(500, f"Conversion failed: {e}")

    background_tasks.add_task(cleanup_job_dir, job_dir)
    download_name = os.path.splitext(file.filename)[0] + ".pdf"
    return FileResponse(
        output_path, media_type="application/pdf", filename=download_name,
        background=background_tasks,
    )


@app.post("/api/merge-pdf")
async def merge_pdf_endpoint(background_tasks: BackgroundTasks, files: List[UploadFile] = File(...)):
    if len(files) < 2:
        raise HTTPException(400, "Please upload at least 2 PDF files to merge.")
    for f in files:
        if not f.filename.lower().endswith(".pdf"):
            raise HTTPException(400, f"'{f.filename}' is not a .pdf file.")

    job_dir = make_job_dir()
    input_paths = []
    try:
        for i, f in enumerate(files):
            p = job_dir / f"input_{i}.pdf"
            await save_upload(f, p)
            input_paths.append(str(p))

        output_path = job_dir / "merged.pdf"
        merge_pdfs(input_paths, str(output_path))
    except HTTPException:
        cleanup_job_dir(job_dir)
        raise
    except Exception as e:
        cleanup_job_dir(job_dir)
        raise HTTPException(500, f"Merge failed: {e}")

    background_tasks.add_task(cleanup_job_dir, job_dir)
    return FileResponse(
        output_path, media_type="application/pdf", filename="merged.pdf",
        background=background_tasks,
    )


@app.post("/api/split-pdf")
async def split_pdf_endpoint(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...),
    start_page: int = Form(...),
    end_page: int = Form(...),
):
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(400, "Please upload a .pdf file.")

    job_dir = make_job_dir()
    input_path = job_dir / "input.pdf"
    output_path = job_dir / "split.pdf"

    await save_upload(file, input_path)

    try:
        split_pdf(str(input_path), str(output_path), start_page, end_page)
    except ValueError as e:
        cleanup_job_dir(job_dir)
        raise HTTPException(400, str(e))
    except Exception as e:
        cleanup_job_dir(job_dir)
        raise HTTPException(500, f"Split failed: {e}")

    background_tasks.add_task(cleanup_job_dir, job_dir)
    download_name = os.path.splitext(file.filename)[0] + f"_p{start_page}-{end_page}.pdf"
    return FileResponse(
        output_path, media_type="application/pdf", filename=download_name,
        background=background_tasks,
    )


@app.post("/api/compress-pdf")
async def compress_pdf_endpoint(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(400, "Please upload a .pdf file.")

    job_dir = make_job_dir()
    input_path = job_dir / "input.pdf"
    output_path = job_dir / "compressed.pdf"

    await save_upload(file, input_path)

    try:
        compress_pdf(str(input_path), str(output_path))
    except Exception as e:
        cleanup_job_dir(job_dir)
        raise HTTPException(500, f"Compression failed: {e}")

    background_tasks.add_task(cleanup_job_dir, job_dir)
    download_name = os.path.splitext(file.filename)[0] + "_compressed.pdf"
    return FileResponse(
        output_path, media_type="application/pdf", filename=download_name,
        background=background_tasks,
    )
