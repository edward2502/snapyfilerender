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
  POST /api/ppt-to-pdf    - upload a .pptx, get back a .pdf
  POST /api/pdf-to-ppt    - upload a .pdf + optional dpi form field, get back a .pptx (one image-slide per page)
  POST /api/remove-pages  - upload a .pdf + pages form field (e.g. "1,3,5"), get back the PDF with those pages removed
  POST /api/organize-pdf  - upload a .pdf + order form field (e.g. "3,1,2"), get back the PDF rebuilt in that page order
  POST /api/extract-pages - upload a .pdf + pages form field (e.g. "2,4"), get back a new PDF with just those pages
  POST /api/repair-pdf    - upload a possibly-corrupted .pdf, get back a repaired copy
  POST /api/html-to-pdf   - upload a .html file, get back a .pdf
  POST /api/page-numbers  - upload a .pdf + optional position/start_at/font_size, get back the PDF with page numbers stamped on
  POST /api/crop-pdf      - upload a .pdf + left/top/right/bottom margins (points), get back the cropped PDF
  POST /api/translate-pdf - upload a .pdf + target_language form field, get back a translated .pdf
  POST /api/pdf-to-markdown - upload a .pdf, get back a .md file
  POST /api/redact-pdf    - upload a .pdf + terms form field (comma-separated), get back the PDF with those terms permanently blacked out
  POST /api/scan-pdf      - upload a scanned/image .pdf + optional language, get back the same PDF with an invisible OCR text layer (searchable/selectable)
  GET  /api/health        - health check

Chat with PDF requires a GEMINI_API_KEY environment variable set on the
server (get a free key at https://aistudio.google.com/apikey). Without
it, /api/chat-pdf and /api/translate-pdf return a clear 500 error rather
than crashing.

Scan PDF (OCR) requires the Tesseract OCR engine to be installed on the
server (the provided Dockerfile installs it via apt). Without it,
/api/scan-pdf returns a clear 500 error rather than crashing.

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
    pptx_to_pdf, pdf_to_pptx,
    remove_pages, organize_pdf, extract_pages, repair_pdf, html_to_pdf,
    add_page_numbers, crop_pdf, translate_pdf, pdf_to_markdown,
    redact_pdf, scan_pdf,
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


def parse_page_list(raw: str) -> list:
    """Parse a comma-separated string of page numbers, e.g. '1,3,5' or
    '3, 1, 2', into a list of ints. Raises HTTPException on bad input."""
    if not raw or not raw.strip():
        raise HTTPException(400, "Please provide at least one page number.")
    pages = []
    for part in raw.split(","):
        part = part.strip()
        if not part:
            continue
        if not part.isdigit():
            raise HTTPException(400, f"'{part}' is not a valid page number.")
        pages.append(int(part))
    if not pages:
        raise HTTPException(400, "Please provide at least one page number.")
    return pages


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


@app.post("/api/ppt-to-pdf")
async def ppt_to_pdf_endpoint(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    if not file.filename.lower().endswith(".pptx"):
        raise HTTPException(400, "Please upload a .pptx file.")

    job_dir = make_job_dir()
    input_path = job_dir / "input.pptx"
    output_path = job_dir / "output.pdf"

    await save_upload(file, input_path)

    try:
        pptx_to_pdf(str(input_path), str(output_path))
    except Exception as e:
        cleanup_job_dir(job_dir)
        raise HTTPException(500, f"Conversion failed: {e}")

    background_tasks.add_task(cleanup_job_dir, job_dir)
    download_name = os.path.splitext(file.filename)[0] + ".pdf"
    return FileResponse(
        output_path, media_type="application/pdf", filename=download_name,
        background=background_tasks,
    )


@app.post("/api/pdf-to-ppt")
async def pdf_to_ppt_endpoint(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...),
    dpi: int = Form(150),
):
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(400, "Please upload a .pdf file.")

    job_dir = make_job_dir()
    input_path = job_dir / "input.pdf"
    output_path = job_dir / "output.pptx"

    await save_upload(file, input_path)

    try:
        pdf_to_pptx(str(input_path), str(output_path), dpi=dpi)
    except Exception as e:
        cleanup_job_dir(job_dir)
        raise HTTPException(500, f"Conversion failed: {e}")

    background_tasks.add_task(cleanup_job_dir, job_dir)
    download_name = os.path.splitext(file.filename)[0] + ".pptx"
    return FileResponse(
        output_path,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        filename=download_name,
        background=background_tasks,
    )


# ---------------------------------------------------------------------------
# Newer PDF tools
# ---------------------------------------------------------------------------

@app.post("/api/remove-pages")
async def remove_pages_endpoint(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...),
    pages: str = Form(...),
):
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(400, "Please upload a .pdf file.")
    pages_to_remove = parse_page_list(pages)

    job_dir = make_job_dir()
    input_path = job_dir / "input.pdf"
    output_path = job_dir / "removed.pdf"
    await save_upload(file, input_path)

    try:
        remove_pages(str(input_path), str(output_path), pages_to_remove)
    except ValueError as e:
        cleanup_job_dir(job_dir)
        raise HTTPException(400, str(e))
    except Exception as e:
        cleanup_job_dir(job_dir)
        raise HTTPException(500, f"Removing pages failed: {e}")

    background_tasks.add_task(cleanup_job_dir, job_dir)
    download_name = os.path.splitext(file.filename)[0] + "_edited.pdf"
    return FileResponse(
        output_path, media_type="application/pdf", filename=download_name,
        background=background_tasks,
    )


@app.post("/api/organize-pdf")
async def organize_pdf_endpoint(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...),
    order: str = Form(...),
):
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(400, "Please upload a .pdf file.")
    page_order = parse_page_list(order)

    job_dir = make_job_dir()
    input_path = job_dir / "input.pdf"
    output_path = job_dir / "organized.pdf"
    await save_upload(file, input_path)

    try:
        organize_pdf(str(input_path), str(output_path), page_order)
    except ValueError as e:
        cleanup_job_dir(job_dir)
        raise HTTPException(400, str(e))
    except Exception as e:
        cleanup_job_dir(job_dir)
        raise HTTPException(500, f"Organizing PDF failed: {e}")

    background_tasks.add_task(cleanup_job_dir, job_dir)
    download_name = os.path.splitext(file.filename)[0] + "_organized.pdf"
    return FileResponse(
        output_path, media_type="application/pdf", filename=download_name,
        background=background_tasks,
    )


@app.post("/api/extract-pages")
async def extract_pages_endpoint(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...),
    pages: str = Form(...),
):
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(400, "Please upload a .pdf file.")
    pages_to_extract = parse_page_list(pages)

    job_dir = make_job_dir()
    input_path = job_dir / "input.pdf"
    output_path = job_dir / "extracted.pdf"
    await save_upload(file, input_path)

    try:
        extract_pages(str(input_path), str(output_path), pages_to_extract)
    except ValueError as e:
        cleanup_job_dir(job_dir)
        raise HTTPException(400, str(e))
    except Exception as e:
        cleanup_job_dir(job_dir)
        raise HTTPException(500, f"Extracting pages failed: {e}")

    background_tasks.add_task(cleanup_job_dir, job_dir)
    download_name = os.path.splitext(file.filename)[0] + "_extracted.pdf"
    return FileResponse(
        output_path, media_type="application/pdf", filename=download_name,
        background=background_tasks,
    )


@app.post("/api/repair-pdf")
async def repair_pdf_endpoint(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(400, "Please upload a .pdf file.")

    job_dir = make_job_dir()
    input_path = job_dir / "input.pdf"
    output_path = job_dir / "repaired.pdf"
    await save_upload(file, input_path)

    try:
        repair_pdf(str(input_path), str(output_path))
    except ValueError as e:
        cleanup_job_dir(job_dir)
        raise HTTPException(400, str(e))
    except Exception as e:
        cleanup_job_dir(job_dir)
        raise HTTPException(500, f"Repair failed: {e}")

    background_tasks.add_task(cleanup_job_dir, job_dir)
    download_name = os.path.splitext(file.filename)[0] + "_repaired.pdf"
    return FileResponse(
        output_path, media_type="application/pdf", filename=download_name,
        background=background_tasks,
    )


@app.post("/api/html-to-pdf")
async def html_to_pdf_endpoint(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    if not file.filename.lower().endswith((".html", ".htm")):
        raise HTTPException(400, "Please upload an .html file.")

    job_dir = make_job_dir()
    input_path = job_dir / "input.html"
    output_path = job_dir / "output.pdf"
    await save_upload(file, input_path)

    try:
        html_to_pdf(str(input_path), str(output_path))
    except Exception as e:
        cleanup_job_dir(job_dir)
        raise HTTPException(500, f"Conversion failed: {e}")

    background_tasks.add_task(cleanup_job_dir, job_dir)
    download_name = os.path.splitext(file.filename)[0] + ".pdf"
    return FileResponse(
        output_path, media_type="application/pdf", filename=download_name,
        background=background_tasks,
    )


@app.post("/api/page-numbers")
async def page_numbers_endpoint(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...),
    position: str = Form("bottom-center"),
    start_at: int = Form(1),
    font_size: int = Form(11),
):
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(400, "Please upload a .pdf file.")

    job_dir = make_job_dir()
    input_path = job_dir / "input.pdf"
    output_path = job_dir / "numbered.pdf"
    await save_upload(file, input_path)

    try:
        add_page_numbers(str(input_path), str(output_path), position, start_at, font_size)
    except ValueError as e:
        cleanup_job_dir(job_dir)
        raise HTTPException(400, str(e))
    except Exception as e:
        cleanup_job_dir(job_dir)
        raise HTTPException(500, f"Adding page numbers failed: {e}")

    background_tasks.add_task(cleanup_job_dir, job_dir)
    download_name = os.path.splitext(file.filename)[0] + "_numbered.pdf"
    return FileResponse(
        output_path, media_type="application/pdf", filename=download_name,
        background=background_tasks,
    )


@app.post("/api/crop-pdf")
async def crop_pdf_endpoint(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...),
    left: float = Form(0),
    top: float = Form(0),
    right: float = Form(0),
    bottom: float = Form(0),
):
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(400, "Please upload a .pdf file.")

    job_dir = make_job_dir()
    input_path = job_dir / "input.pdf"
    output_path = job_dir / "cropped.pdf"
    await save_upload(file, input_path)

    try:
        crop_pdf(str(input_path), str(output_path), left, top, right, bottom)
    except ValueError as e:
        cleanup_job_dir(job_dir)
        raise HTTPException(400, str(e))
    except Exception as e:
        cleanup_job_dir(job_dir)
        raise HTTPException(500, f"Cropping failed: {e}")

    background_tasks.add_task(cleanup_job_dir, job_dir)
    download_name = os.path.splitext(file.filename)[0] + "_cropped.pdf"
    return FileResponse(
        output_path, media_type="application/pdf", filename=download_name,
        background=background_tasks,
    )


@app.post("/api/translate-pdf")
async def translate_pdf_endpoint(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...),
    target_language: str = Form(...),
):
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(400, "Please upload a .pdf file.")

    job_dir = make_job_dir()
    input_path = job_dir / "input.pdf"
    output_path = job_dir / "translated.pdf"
    await save_upload(file, input_path)

    try:
        translate_pdf(str(input_path), str(output_path), target_language, GEMINI_API_KEY)
    except ValueError as e:
        cleanup_job_dir(job_dir)
        raise HTTPException(400, str(e))
    except RuntimeError as e:
        cleanup_job_dir(job_dir)
        raise HTTPException(500, str(e))
    except Exception as e:
        cleanup_job_dir(job_dir)
        raise HTTPException(500, f"Translation failed: {e}")

    background_tasks.add_task(cleanup_job_dir, job_dir)
    download_name = os.path.splitext(file.filename)[0] + "_translated.pdf"
    return FileResponse(
        output_path, media_type="application/pdf", filename=download_name,
        background=background_tasks,
    )


@app.post("/api/pdf-to-markdown")
async def pdf_to_markdown_endpoint(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(400, "Please upload a .pdf file.")

    job_dir = make_job_dir()
    input_path = job_dir / "input.pdf"
    output_path = job_dir / "output.md"
    await save_upload(file, input_path)

    try:
        pdf_to_markdown(str(input_path), str(output_path))
    except ValueError as e:
        cleanup_job_dir(job_dir)
        raise HTTPException(400, str(e))
    except Exception as e:
        cleanup_job_dir(job_dir)
        raise HTTPException(500, f"Conversion failed: {e}")

    background_tasks.add_task(cleanup_job_dir, job_dir)
    download_name = os.path.splitext(file.filename)[0] + ".md"
    return FileResponse(
        output_path, media_type="text/markdown", filename=download_name,
        background=background_tasks,
    )


@app.post("/api/redact-pdf")
async def redact_pdf_endpoint(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...),
    terms: str = Form(...),
):
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(400, "Please upload a .pdf file.")
    term_list = [t for t in terms.split(",")]
    if not any(t.strip() for t in term_list):
        raise HTTPException(400, "Please provide at least one word or phrase to redact.")

    job_dir = make_job_dir()
    input_path = job_dir / "input.pdf"
    output_path = job_dir / "redacted.pdf"
    await save_upload(file, input_path)

    try:
        redact_pdf(str(input_path), str(output_path), term_list)
    except ValueError as e:
        cleanup_job_dir(job_dir)
        raise HTTPException(400, str(e))
    except Exception as e:
        cleanup_job_dir(job_dir)
        raise HTTPException(500, f"Redaction failed: {e}")

    background_tasks.add_task(cleanup_job_dir, job_dir)
    download_name = os.path.splitext(file.filename)[0] + "_redacted.pdf"
    return FileResponse(
        output_path, media_type="application/pdf", filename=download_name,
        background=background_tasks,
    )


@app.post("/api/scan-pdf")
async def scan_pdf_endpoint(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...),
    language: str = Form("eng"),
):
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(400, "Please upload a .pdf file.")

    job_dir = make_job_dir()
    input_path = job_dir / "input.pdf"
    output_path = job_dir / "scanned.pdf"
    await save_upload(file, input_path)

    try:
        scan_pdf(str(input_path), str(output_path), language)
    except ValueError as e:
        cleanup_job_dir(job_dir)
        raise HTTPException(400, str(e))
    except RuntimeError as e:
        cleanup_job_dir(job_dir)
        raise HTTPException(500, str(e))
    except Exception as e:
        cleanup_job_dir(job_dir)
        raise HTTPException(500, f"Scan/OCR failed: {e}")

    background_tasks.add_task(cleanup_job_dir, job_dir)
    download_name = os.path.splitext(file.filename)[0] + "_scanned.pdf"
    return FileResponse(
        output_path, media_type="application/pdf", filename=download_name,
        background=background_tasks,
    )
