"""
converter.py - Core conversion logic used by the API.

Word <-> PDF: LibreOffice does docx->pdf; pdf2docx does pdf->docx.
Excel <-> PDF: LibreOffice does xlsx->pdf (with auto column-widening and
               thin cell borders added first, so the exported PDF has a
               real grid instead of ambiguous plain text); pdfplumber
               extracts tables back out of a PDF into an .xlsx workbook.
"""

import os
import shutil
import subprocess
import tempfile


def find_soffice():
    for name in ("soffice", "libreoffice"):
        path = shutil.which(name)
        if path:
            return path
    candidates = [
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        r"C:\Program Files\LibreOffice\program\soffice.exe",
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    return None


def _office_to_pdf(input_path: str, output_path: str) -> str:
    soffice = find_soffice()
    if not soffice:
        raise RuntimeError("LibreOffice (soffice) not found on this server.")

    with tempfile.TemporaryDirectory() as tmpdir:
        cmd = [
            soffice, "--headless", "--norestore",
            "--convert-to", "pdf",
            "--outdir", tmpdir,
            input_path,
        ]
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=180)
        if result.returncode != 0:
            raise RuntimeError(f"LibreOffice conversion failed: {result.stderr}")

        produced = os.path.join(
            tmpdir, os.path.splitext(os.path.basename(input_path))[0] + ".pdf"
        )
        if not os.path.exists(produced):
            raise RuntimeError("Conversion did not produce a PDF.")
        shutil.move(produced, output_path)

    return output_path


def word_to_pdf(input_path: str, output_path: str) -> str:
    return _office_to_pdf(input_path, output_path)


def pdf_to_word(input_path: str, output_path: str) -> str:
    from pdf2docx import Converter

    cv = Converter(input_path)
    try:
        cv.convert(output_path)
    finally:
        cv.close()

    return output_path


def _autofit_and_border_xlsx(input_path, working_copy_path):
    """Widen columns that would otherwise get clipped on PDF export, and
    add thin borders so the exported PDF has a real grid - which makes
    pdf_to_excel's extraction exact instead of guesswork."""
    import openpyxl
    from openpyxl.styles import Border, Side

    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    wb = openpyxl.load_workbook(input_path)
    for ws in wb.worksheets:
        for col_cells in ws.columns:
            longest = max((len(str(c.value)) for c in col_cells if c.value is not None), default=0)
            if longest == 0:
                continue
            col_letter = col_cells[0].column_letter
            target_width = max(longest + 2, 8)
            current = ws.column_dimensions[col_letter].width
            if current is None or current < target_width:
                ws.column_dimensions[col_letter].width = target_width

        if ws.dimensions and ws.dimensions != "A1:A1":
            for row in ws.iter_rows():
                for cell in row:
                    if cell.value is not None:
                        cell.border = border

    wb.save(working_copy_path)
    return working_copy_path


def excel_to_pdf(input_path: str, output_path: str) -> str:
    with tempfile.TemporaryDirectory() as tmpdir:
        widened = os.path.join(tmpdir, os.path.basename(input_path))
        try:
            _autofit_and_border_xlsx(input_path, widened)
            source = widened
        except Exception:
            source = input_path
        return _office_to_pdf(source, output_path)


def _extract_table_from_words(words, row_tolerance=4):
    """Fallback table reconstruction for PDFs with no visible ruling lines."""
    if not words:
        return []

    rows = {}
    for w in words:
        key = round(w["top"] / row_tolerance)
        rows.setdefault(key, []).append(w)

    row_keys = sorted(rows.keys())
    ordered_rows = [sorted(rows[k], key=lambda w: w["x0"]) for k in row_keys]
    if not ordered_rows:
        return []

    ncols = len(ordered_rows[0])
    if ncols == 0:
        return []

    def row_to_cells(row_words):
        cells = [dict(w) for w in row_words]
        while len(cells) > ncols:
            gaps = [cells[i + 1]["x0"] - cells[i]["x1"] for i in range(len(cells) - 1)]
            i = gaps.index(min(gaps))
            merged_text = cells[i]["text"] + " " + cells[i + 1]["text"]
            cells[i] = {"text": merged_text, "x0": cells[i]["x0"], "x1": cells[i + 1]["x1"]}
            del cells[i + 1]
        while len(cells) < ncols:
            cells.append({"text": "", "x0": None, "x1": None})
        return cells

    rows_of_cells = [row_to_cells(r) for r in ordered_rows]

    if len(rows_of_cells) < 2 or ncols < 2:
        return []

    for col in range(ncols):
        x0s = [row[col]["x0"] for row in rows_of_cells if row[col]["x0"] is not None]
        if len(x0s) < 2:
            continue
        if max(x0s) - min(x0s) > 12:
            return []

    return [[c["text"] for c in row] for row in rows_of_cells]


def pdf_to_excel(input_path: str, output_path: str) -> str:
    import pdfplumber
    import pandas as pd

    page_frames = {}
    with pdfplumber.open(input_path) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            tables = page.extract_tables()
            frames = []

            if tables:
                for table in tables:
                    if not table or len(table) < 1:
                        continue
                    header, *rows = table
                    frames.append(pd.DataFrame(rows, columns=header))
            else:
                words = page.extract_words()
                table = _extract_table_from_words(words)
                if len(table) > 1:
                    header, *rows = table
                    frames.append(pd.DataFrame(rows, columns=header))

            if frames:
                page_frames[page_num] = frames

    if not page_frames:
        raise RuntimeError(
            "No tables were found in this PDF. Excel conversion only works "
            "on PDFs that actually contain tabular data."
        )

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for page_num, frames in page_frames.items():
            sheet_name = f"Page{page_num}"[:31]
            start_row = 0
            for df in frames:
                df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=start_row)
                start_row += len(df) + 2

    return output_path


def merge_pdfs(input_paths: list, output_path: str) -> str:
    """Combine multiple PDFs into one, in the given order."""
    from pypdf import PdfWriter

    writer = PdfWriter()
    for path in input_paths:
        writer.append(path)
    with open(output_path, "wb") as f:
        writer.write(f)
    return output_path


def split_pdf(input_path: str, output_path: str, start_page: int, end_page: int) -> str:
    """Extract a page range (1-indexed, inclusive) into a new PDF."""
    from pypdf import PdfReader, PdfWriter

    reader = PdfReader(input_path)
    total_pages = len(reader.pages)

    if start_page < 1 or end_page < 1:
        raise ValueError("Page numbers must be 1 or greater.")
    if start_page > total_pages:
        raise ValueError(f"This PDF only has {total_pages} page(s).")
    if end_page > total_pages:
        end_page = total_pages
    if start_page > end_page:
        raise ValueError("Start page must be less than or equal to end page.")

    writer = PdfWriter()
    for i in range(start_page - 1, end_page):
        writer.add_page(reader.pages[i])

    with open(output_path, "wb") as f:
        writer.write(f)
    return output_path


def compress_pdf(input_path: str, output_path: str, image_quality: int = 60, max_dimension: int = 1600) -> str:
    """Shrink a PDF's file size by re-encoding its embedded images at a
    lower quality/resolution and compressing internal object streams.
    Text and vector content are left untouched - only raster images (the
    usual cause of PDF bloat) are recompressed."""
    import io
    import pikepdf
    from PIL import Image

    pdf = pikepdf.open(input_path)

    for page in pdf.pages:
        if "/Resources" not in page or "/XObject" not in page.Resources:
            continue
        xobjects = page.Resources.XObject
        for name in list(xobjects.keys()):
            obj = xobjects[name]
            if obj.get("/Subtype") != pikepdf.Name("/Image"):
                continue
            try:
                pdf_image = pikepdf.PdfImage(obj)
                pil_image = pdf_image.as_pil_image()
            except Exception:
                continue  # unsupported image encoding - leave it as-is

            # Downscale oversized images
            w, h = pil_image.size
            if max(w, h) > max_dimension:
                scale = max_dimension / max(w, h)
                pil_image = pil_image.resize((int(w * scale), int(h * scale)), Image.LANCZOS)

            if pil_image.mode not in ("RGB", "L"):
                pil_image = pil_image.convert("RGB")

            buf = io.BytesIO()
            pil_image.save(buf, format="JPEG", quality=image_quality, optimize=True)
            buf.seek(0)

            try:
                new_image = pikepdf.PdfImage.make_image(pdf, buf.read(), pil_image.size)
            except AttributeError:
                # Fallback for pikepdf versions without make_image helper:
                # replace the stream data directly with a DCTDecode (JPEG) filter.
                obj.write(buf.getvalue(), filter=pikepdf.Name("/DCTDecode"))
                obj.Width = pil_image.width
                obj.Height = pil_image.height
                obj.ColorSpace = pikepdf.Name("/DeviceRGB") if pil_image.mode == "RGB" else pikepdf.Name("/DeviceGray")
                obj.BitsPerComponent = 8
                continue

            xobjects[name] = new_image

    pdf.save(output_path, compress_streams=True, object_stream_mode=pikepdf.ObjectStreamMode.generate)
    pdf.close()
    return output_path


def unlock_pdf(input_path: str, output_path: str, password: str) -> str:
    """Remove password protection from a PDF, given the correct password."""
    from pypdf import PdfReader, PdfWriter
    from pypdf.errors import FileNotDecryptedError

    reader = PdfReader(input_path)
    if reader.is_encrypted:
        result = reader.decrypt(password)
        if result == 0:
            raise ValueError("Incorrect password.")

    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)

    with open(output_path, "wb") as f:
        writer.write(f)
    return output_path


def protect_pdf(input_path: str, output_path: str, password: str) -> str:
    """Add password protection to a PDF."""
    from pypdf import PdfReader, PdfWriter

    if not password:
        raise ValueError("Password cannot be empty.")

    reader = PdfReader(input_path)
    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)

    writer.encrypt(user_password=password, algorithm="AES-256")

    with open(output_path, "wb") as f:
        writer.write(f)
    return output_path


def watermark_pdf(input_path: str, output_path: str, text: str, opacity: float = 0.3) -> str:
    """Stamp a diagonal text watermark across every page."""
    import io
    from pypdf import PdfReader, PdfWriter
    from reportlab.pdfgen import canvas
    from reportlab.lib.colors import Color

    reader = PdfReader(input_path)
    writer = PdfWriter()

    for page in reader.pages:
        page_width = float(page.mediabox.width)
        page_height = float(page.mediabox.height)

        buf = io.BytesIO()
        c = canvas.Canvas(buf, pagesize=(page_width, page_height))
        c.saveState()
        c.translate(page_width / 2, page_height / 2)
        c.rotate(45)
        c.setFillColor(Color(0.5, 0.5, 0.5, alpha=opacity))
        font_size = max(24, min(page_width, page_height) / 10)
        c.setFont("Helvetica-Bold", font_size)
        c.drawCentredString(0, 0, text)
        c.restoreState()
        c.save()
        buf.seek(0)

        watermark_reader = PdfReader(buf)
        page.merge_page(watermark_reader.pages[0])
        writer.add_page(page)

    with open(output_path, "wb") as f:
        writer.write(f)
    return output_path


def generate_qr_code(data: str, output_path: str, box_size: int = 10, fill_color: str = "black", back_color: str = "white") -> str:
    """Generate a QR code PNG encoding the given text/URL."""
    import qrcode

    if not data or not data.strip():
        raise ValueError("QR code content cannot be empty.")

    qr = qrcode.QRCode(
        version=None,  # auto-size to fit the data
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=box_size,
        border=4,
    )
    qr.add_data(data)
    qr.make(fit=True)

    img = qr.make_image(fill_color=fill_color, back_color=back_color)
    img.save(output_path)
    return output_path


def extract_pdf_text(input_path: str, max_chars: int = 100_000) -> str:
    """Pull all text out of a PDF for use as AI context. Capped at
    max_chars so a huge PDF can't blow past the model's context window
    or balloon token usage on the free tier."""
    import pdfplumber

    text_parts = []
    total_len = 0
    with pdfplumber.open(input_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text() or ""
            text_parts.append(page_text)
            total_len += len(page_text)
            if total_len >= max_chars:
                break

    full_text = "\n\n".join(text_parts)
    if len(full_text) > max_chars:
        full_text = full_text[:max_chars] + "\n\n[... document truncated for length ...]"

    if not full_text.strip():
        raise ValueError(
            "No text could be extracted from this PDF. It may be a scanned "
            "image without a text layer - Chat with PDF only works on PDFs "
            "with selectable text."
        )
    return full_text


def chat_with_pdf(document_text: str, question: str, api_key: str) -> str:
    """Answer a question about a PDF's contents using Google's Gemini API.
    document_text should already be extracted (see extract_pdf_text) so
    that a multi-turn conversation doesn't need to re-parse the PDF on
    every question."""
    from google import genai

    if not question or not question.strip():
        raise ValueError("Question cannot be empty.")
    if not api_key:
        raise RuntimeError(
            "No Gemini API key configured on the server (GEMINI_API_KEY env var missing)."
        )

    client = genai.Client(api_key=api_key)

    prompt = (
        "You are answering questions about a specific document. Only use "
        "the document text below to answer - if the answer isn't in the "
        "document, say so clearly rather than guessing.\n\n"
        "--- DOCUMENT START ---\n"
        f"{document_text}\n"
        "--- DOCUMENT END ---\n\n"
        f"Question: {question}"
    )

    response = client.models.generate_content(
        model="gemini-2.0-flash",
        contents=prompt,
    )
    return response.text
