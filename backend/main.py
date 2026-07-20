"""
main.py - FastAPI backend for a Word<->PDF converter website (iLovePDF-style).

Endpoints:
  POST /api/to-pdf        - upload a .docx, get back a .pdf
  POST /api/to-word       - upload a .pdf,  get back a .docx
  POST /api/to-excel      - upload a .pdf,  get back a .xlsx
  POST /api/from-excel    - upload a .xlsx, get back a .pdf
  POST /api/merge-pdf     - upload 2+ .pdf files (field "files"), get back one merged .pdf
  POST /api/split-pdf     - upload a .pdf + start_page/end_page form fields, get back that page range as a .pdf
  POST /api/compress-pdf  - upload a .pdf, get back a smaller .pdf (images recompressed)
  POST /api/unlock-pdf    - upload a .pdf + password form field, get back the unlocked .pdf
  POST /api/protect-pdf   - upload a .pdf + password form field, get back a password-protected .pdf
  POST /api/watermark-pdf - upload a .pdf + text form field, get back a watermarked .pdf
  POST /api/generate-qr   - data (+ optional fill_color/back_color) form fields, get back a .png QR code
  POST /api/chat-pdf      - upload a .pdf + question form field, get back {"answer": "..."} from Gemini
  POST /api/img-to-pdf    - upload 1+ image files (field "files"), get back one .pdf (one image per page)
  POST /api/pdf-to-img    - upload a .pdf + optional dpi form field, get back a .zip of one PNG per page
  POST /api/compress-img  - upload an image + optional quality/max_dimension form fields, get back a smaller .jpg
  POST /api/upscale-img   - upload an image + optional scale_factor form field (default 2.0, max 4.0), get back a larger image
  GET  /api/health        - health check

Chat with PDF requires a GEMINI_API_KEY environment variable set on the
server (get a free key at https://aistudio.google.com/apikey). Without
it, /api/chat-pdf returns a clear 500 error rather than crashing.

Files are saved to a temp job folder, converted, streamed back to the
user, and deleted afterwards (via background task) so nothing piles up
on disk.
"""

import os
import uuid
import shutil
import zipfile
from pathlib import Path
from typing import List

from fastapi import FastAPI, UploadFile, File, Form, HTTPException, BackgroundTasks
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse

from converter import (
    word_to_pdf, pdf_to_word, excel_to_pdf, pdf_to_excel,
    merge_pdfs, split_pdf, compress_pdf,
    unlock_pdf, protect_pdf, watermark_pdf,
    generate_qr_code, extract_pdf_text, chat_with_pdf,
    images_to_pdf, pdf_to_images, compress_image, upscale_image,
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


@app.post("/api/unlock-pdf")
async def unlock_pdf_endpoint(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...),
    password: str = Form(...),
):
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(400, "Please upload a .pdf file.")

    job_dir = make_job_dir()
    input_path = job_dir / "input.pdf"
    output_path = job_dir / "unlocked.pdf"

    await save_upload(file, input_path)

    try:
        unlock_pdf(str(input_path), str(output_path), password)
    except ValueError as e:
        cleanup_job_dir(job_dir)
        raise HTTPException(400, str(e))
    except Exception as e:
        cleanup_job_dir(job_dir)
        raise HTTPException(500, f"Unlock failed: {e}")

    background_tasks.add_task(cleanup_job_dir, job_dir)
    download_name = os.path.splitext(file.filename)[0] + "_unlocked.pdf"
    return FileResponse(
        output_path, media_type="application/pdf", filename=download_name,
        background=background_tasks,
    )


@app.post("/api/protect-pdf")
async def protect_pdf_endpoint(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...),
    password: str = Form(...),
):
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(400, "Please upload a .pdf file.")

    job_dir = make_job_dir()
    input_path = job_dir / "input.pdf"
    output_path = job_dir / "protected.pdf"

    await save_upload(file, input_path)

    try:
        protect_pdf(str(input_path), str(output_path), password)
    except ValueError as e:
        cleanup_job_dir(job_dir)
        raise HTTPException(400, str(e))
    except Exception as e:
        cleanup_job_dir(job_dir)
        raise HTTPException(500, f"Protection failed: {e}")

    background_tasks.add_task(cleanup_job_dir, job_dir)
    download_name = os.path.splitext(file.filename)[0] + "_protected.pdf"
    return FileResponse(
        output_path, media_type="application/pdf", filename=download_name,
        background=background_tasks,
    )


@app.post("/api/watermark-pdf")
async def watermark_pdf_endpoint(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...),
    text: str = Form(...),
):
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(400, "Please upload a .pdf file.")
    if not text.strip():
        raise HTTPException(400, "Watermark text cannot be empty.")

    job_dir = make_job_dir()
    input_path = job_dir / "input.pdf"
    output_path = job_dir / "watermarked.pdf"

    await save_upload(file, input_path)

    try:
        watermark_pdf(str(input_path), str(output_path), text)
    except Exception as e:
        cleanup_job_dir(job_dir)
        raise HTTPException(500, f"Watermarking failed: {e}")

    background_tasks.add_task(cleanup_job_dir, job_dir)
    download_name = os.path.splitext(file.filename)[0] + "_watermarked.pdf"
    return FileResponse(
        output_path, media_type="application/pdf", filename=download_name,
        background=background_tasks,
    )


@app.post("/api/generate-qr")
async def generate_qr_endpoint(
    background_tasks: BackgroundTasks,
    data: str = Form(...),
    fill_color: str = Form("black"),
    back_color: str = Form("white"),
):
    job_dir = make_job_dir()
    output_path = job_dir / "qrcode.png"

    try:
        generate_qr_code(data, str(output_path), fill_color=fill_color, back_color=back_color)
    except ValueError as e:
        cleanup_job_dir(job_dir)
        raise HTTPException(400, str(e))
    except Exception as e:
        cleanup_job_dir(job_dir)
        raise HTTPException(500, f"QR code generation failed: {e}")

    background_tasks.add_task(cleanup_job_dir, job_dir)
    return FileResponse(
        output_path, media_type="image/png", filename="qrcode.png",
        background=background_tasks,
    )


GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")


@app.post("/api/chat-pdf")
async def chat_pdf_endpoint(
    file: UploadFile = File(...),
    question: str = Form(...),
):
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(400, "Please upload a .pdf file.")

    job_dir = make_job_dir()
    input_path = job_dir / "input.pdf"

    try:
        await save_upload(file, input_path)
        document_text = extract_pdf_text(str(input_path))
        answer = chat_with_pdf(document_text, question, GEMINI_API_KEY)
    except ValueError as e:
        raise HTTPException(400, str(e))
    except RuntimeError as e:
        raise HTTPException(500, str(e))
    except Exception as e:
        raise HTTPException(500, f"Chat with PDF failed: {e}")
    finally:
        cleanup_job_dir(job_dir)

    return JSONResponse({"answer": answer})


ALLOWED_IMAGE_EXTS = (".jpg", ".jpeg", ".png", ".bmp", ".webp", ".tiff")


@app.post("/api/img-to-pdf")
async def img_to_pdf_endpoint(background_tasks: BackgroundTasks, files: List[UploadFile] = File(...)):
    if len(files) < 1:
        raise HTTPException(400, "Please upload at least 1 image.")
    for f in files:
        if not f.filename.lower().endswith(ALLOWED_IMAGE_EXTS):
            raise HTTPException(400, f"'{f.filename}' is not a supported image type.")

    job_dir = make_job_dir()
    input_paths = []
    try:
        for i, f in enumerate(files):
            ext = os.path.splitext(f.filename)[1]
            p = job_dir / f"input_{i}{ext}"
            await save_upload(f, p)
            input_paths.append(str(p))

        output_path = job_dir / "images.pdf"
        images_to_pdf(input_paths, str(output_path))
    except HTTPException:
        cleanup_job_dir(job_dir)
        raise
    except Exception as e:
        cleanup_job_dir(job_dir)
        raise HTTPException(500, f"Image to PDF failed: {e}")

    background_tasks.add_task(cleanup_job_dir, job_dir)
    return FileResponse(
        output_path, media_type="application/pdf", filename="images.pdf",
        background=background_tasks,
    )


@app.post("/api/pdf-to-img")
async def pdf_to_img_endpoint(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...),
    dpi: int = Form(150),
):
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(400, "Please upload a .pdf file.")

    job_dir = make_job_dir()
    input_path = job_dir / "input.pdf"
    images_dir = job_dir / "images"
    images_dir.mkdir()

    await save_upload(file, input_path)

    try:
        image_paths = pdf_to_images(str(input_path), str(images_dir), dpi=dpi)
    except ValueError as e:
        cleanup_job_dir(job_dir)
        raise HTTPException(400, str(e))
    except Exception as e:
        cleanup_job_dir(job_dir)
        raise HTTPException(500, f"PDF to image failed: {e}")

    zip_path = job_dir / "pages.zip"
    with zipfile.ZipFile(zip_path, "w") as zf:
        for p in image_paths:
            zf.write(p, arcname=os.path.basename(p))

    background_tasks.add_task(cleanup_job_dir, job_dir)
    download_name = os.path.splitext(file.filename)[0] + "_pages.zip"
    return FileResponse(
        zip_path, media_type="application/zip", filename=download_name,
        background=background_tasks,
    )


@app.post("/api/compress-img")
async def compress_img_endpoint(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...),
    quality: int = Form(60),
    max_dimension: int = Form(None),
):
    if not file.filename.lower().endswith(ALLOWED_IMAGE_EXTS):
        raise HTTPException(400, "Please upload a supported image file.")

    job_dir = make_job_dir()
    ext = os.path.splitext(file.filename)[1]
    input_path = job_dir / f"input{ext}"
    output_path = job_dir / "compressed.jpg"

    await save_upload(file, input_path)

    try:
        compress_image(str(input_path), str(output_path), quality=quality, max_dimension=max_dimension)
    except Exception as e:
        cleanup_job_dir(job_dir)
        raise HTTPException(500, f"Image compression failed: {e}")

    background_tasks.add_task(cleanup_job_dir, job_dir)
    download_name = os.path.splitext(file.filename)[0] + "_compressed.jpg"
    return FileResponse(
        output_path, media_type="image/jpeg", filename=download_name,
        background=background_tasks,
    )


@app.post("/api/upscale-img")
async def upscale_img_endpoint(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...),
    scale_factor: float = Form(2.0),
):
    if not file.filename.lower().endswith(ALLOWED_IMAGE_EXTS):
        raise HTTPException(400, "Please upload a supported image file.")

    job_dir = make_job_dir()
    ext = os.path.splitext(file.filename)[1]
    input_path = job_dir / f"input{ext}"
    output_path = job_dir / f"upscaled{ext}"

    await save_upload(file, input_path)

    try:
        upscale_image(str(input_path), str(output_path), scale_factor=scale_factor)
    except ValueError as e:
        cleanup_job_dir(job_dir)
        raise HTTPException(400, str(e))
    except Exception as e:
        cleanup_job_dir(job_dir)
        raise HTTPException(500, f"Image upscaling failed: {e}")

    background_tasks.add_task(cleanup_job_dir, job_dir)
    download_name = os.path.splitext(file.filename)[0] + "_upscaled" + ext
    media_type = "image/jpeg" if ext.lower() in (".jpg", ".jpeg") else "image/png"
    return FileResponse(
        output_path, media_type=media_type, filename=download_name,
        background=background_tasks,
    )
