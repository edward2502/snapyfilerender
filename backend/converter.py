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
