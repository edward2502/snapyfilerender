#!/usr/bin/env python3
"""
convert.py - Convert Word (.docx) <-> PDF (.pdf)

Usage:
    python convert.py to-pdf input.docx [output.pdf]
    python convert.py to-word input.pdf [output.docx]
    python convert.py auto input_file [output_file]   # picks direction from extension

    python convert.py to-excel input.pdf [output.xlsx]
    python convert.py from-excel input.xlsx [output.pdf]

Requirements:
    pip install pdf2docx docx2pdf pdfplumber pandas openpyxl --break-system-packages

    Word -> PDF uses LibreOffice ("soffice") if it's installed on the system
    (works on Linux/macOS/Windows, no MS Word needed). If LibreOffice isn't
    found, it falls back to `docx2pdf`, which requires MS Word (Windows/macOS).

    PDF -> Word uses `pdf2docx`, which is pure Python and needs no external
    program.

    Excel -> PDF also uses LibreOffice (same engine, just a different
    document type).

    PDF -> Excel uses `pdfplumber` to detect and extract tables from each
    page, then writes each page's table(s) to a separate sheet with
    `pandas`/`openpyxl`. This works well for PDFs that actually contain
    tables (the common case for "convert this PDF to Excel"); a PDF that's
    just prose won't have anything meaningful to extract.

Install LibreOffice if you don't have it:
    Ubuntu/Debian: sudo apt install libreoffice
    macOS:         brew install --cask libreoffice
    Windows:       https://www.libreoffice.org/download/
"""

import argparse
import os
import shutil
import subprocess
import sys
import tempfile


def find_soffice():
    """Locate the LibreOffice/soffice executable, if present."""
    for name in ("soffice", "libreoffice"):
        path = shutil.which(name)
        if path:
            return path
    # Common install locations on macOS/Windows that might not be on PATH
    candidates = [
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    return None


def _office_to_pdf(input_path, output_path=None):
    """Shared LibreOffice-based conversion for any office doc -> PDF
    (Word .docx or Excel .xlsx alike)."""
    input_path = os.path.abspath(input_path)
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"Input file not found: {input_path}")

    if output_path is None:
        output_path = os.path.splitext(input_path)[0] + ".pdf"
    output_path = os.path.abspath(output_path)
    out_dir = os.path.dirname(output_path)
    os.makedirs(out_dir, exist_ok=True)

    soffice = find_soffice()
    if soffice:
        # LibreOffice writes to a directory with the same basename + .pdf,
        # so convert into a temp dir and then move/rename to the requested name.
        with tempfile.TemporaryDirectory() as tmpdir:
            cmd = [
                soffice,
                "--headless",
                "--norestore",
                "--convert-to", "pdf",
                "--outdir", tmpdir,
                input_path,
            ]
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=180)
            if result.returncode != 0:
                raise RuntimeError(
                    f"LibreOffice conversion failed:\n{result.stdout}\n{result.stderr}"
                )
            produced = os.path.join(
                tmpdir, os.path.splitext(os.path.basename(input_path))[0] + ".pdf"
            )
            if not os.path.exists(produced):
                raise RuntimeError("LibreOffice did not produce the expected PDF output.")
            shutil.move(produced, output_path)
        print(f"[LibreOffice] Converted: {input_path} -> {output_path}")
        return output_path

    # Fallback: docx2pdf (needs MS Word/Excel installed; Windows/macOS only)
    try:
        from docx2pdf import convert as docx2pdf_convert
    except ImportError:
        raise RuntimeError(
            "No LibreOffice found and docx2pdf is not installed.\n"
            "Install LibreOffice, or run: pip install docx2pdf"
        )

    docx2pdf_convert(input_path, output_path)
    print(f"[docx2pdf] Converted: {input_path} -> {output_path}")
    return output_path


def word_to_pdf(input_path, output_path=None):
    return _office_to_pdf(input_path, output_path)


def _autofit_xlsx_columns(input_path, working_copy_path):
    """Copy the workbook, widen any column whose content would otherwise
    get clipped in the PDF export, and add thin cell borders around the
    used range. LibreOffice does not auto-wrap/auto-fit on export, so a
    too-narrow column can visually run into the next one - and a PDF with
    no ruling lines at all makes later PDF->Excel extraction unreliable
    (it has to guess column boundaries from text alignment instead of
    reading a real grid). Adding borders here means a PDF produced by our
    own excel_to_pdf can always be read back exactly by pdf_to_excel."""
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


def excel_to_pdf(input_path, output_path=None):
    input_path = os.path.abspath(input_path)
    if output_path is None:
        output_path = os.path.splitext(input_path)[0] + ".pdf"
    output_path = os.path.abspath(output_path)

    with tempfile.TemporaryDirectory() as tmpdir:
        widened = os.path.join(tmpdir, os.path.basename(input_path))
        try:
            _autofit_xlsx_columns(input_path, widened)
            source_for_conversion = widened
        except Exception:
            # If auto-fit fails for any reason, fall back to converting
            # the original file as-is rather than blocking the conversion.
            source_for_conversion = input_path
        return _office_to_pdf(source_for_conversion, output_path)


def _extract_table_from_words(words, row_tolerance=4):
    """Fallback table reconstruction for PDFs with no visible ruling lines.
    Groups words into rows by vertical position, then assumes the header
    row's word count is the true column count, merging extra words in
    later rows (e.g. multi-word cell values) by collapsing the smallest
    horizontal gaps first. This handles the common case where a column is
    right-aligned (numbers) while its header is left-aligned, which trips
    up naive x-position column clustering."""
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
        # Too many words for the column count -> merge closest pair repeatedly
        while len(cells) > ncols:
            gaps = [cells[i + 1]["x0"] - cells[i]["x1"] for i in range(len(cells) - 1)]
            i = gaps.index(min(gaps))
            merged_text = cells[i]["text"] + " " + cells[i + 1]["text"]
            cells[i] = {"text": merged_text, "x0": cells[i]["x0"], "x1": cells[i + 1]["x1"]}
            del cells[i + 1]
        # Too few words (e.g. a blank cell) -> pad so rows stay rectangular
        while len(cells) < ncols:
            cells.append({"text": "", "x0": None, "x1": None})
        return cells

    rows_of_cells = [row_to_cells(r) for r in ordered_rows]

    if len(rows_of_cells) < 2 or ncols < 2:
        return []  # not enough structure to call this a table

    # Reject false positives (e.g. wrapped prose): a real table has each
    # column's left edge landing in roughly the same place across rows.
    # Flowing text won't - its "columns" drift with whatever each line
    # happened to wrap to.
    for col in range(ncols):
        x0s = [row[col]["x0"] for row in rows_of_cells if row[col]["x0"] is not None]
        if len(x0s) < 2:
            continue
        spread = max(x0s) - min(x0s)
        if spread > 12:  # points; real columns line up far more tightly than this
            return []

    return [[c["text"] for c in row] for row in rows_of_cells]


def pdf_to_excel(input_path, output_path=None):
    """Extract tables from each page of a PDF and write them into an
    .xlsx workbook, one sheet per page (pages with multiple tables get
    the tables stacked with a blank row between them)."""
    import pdfplumber
    import pandas as pd

    input_path = os.path.abspath(input_path)
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"Input file not found: {input_path}")

    if output_path is None:
        output_path = os.path.splitext(input_path)[0] + ".xlsx"
    output_path = os.path.abspath(output_path)
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    # First pass: extract everything into memory. Only once we know there's
    # at least one real table do we open the ExcelWriter - otherwise
    # openpyxl errors on trying to save a workbook with zero sheets, which
    # is a confusing error to surface instead of a clear "nothing found".
    page_frames = {}  # page_num -> list[DataFrame]
    with pdfplumber.open(input_path) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            tables = page.extract_tables()  # default: detects ruled/bordered tables
            frames = []

            if tables:
                for table in tables:
                    if not table or len(table) < 1:
                        continue
                    header, *rows = table
                    frames.append(pd.DataFrame(rows, columns=header))
            else:
                # No visible ruling lines (common for plain data exports) -
                # reconstruct rows/columns from word positions instead.
                words = page.extract_words()
                table = _extract_table_from_words(words)
                if len(table) > 1:
                    header, *rows = table
                    frames.append(pd.DataFrame(rows, columns=header))

            if frames:
                page_frames[page_num] = frames

    if not page_frames:
        raise RuntimeError(
            "No tables were found in this PDF. pdf_to_excel only extracts "
            "tabular data; a text-only PDF has nothing to convert."
        )

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for page_num, frames in page_frames.items():
            sheet_name = f"Page{page_num}"[:31]  # Excel sheet name limit
            start_row = 0
            for df in frames:
                df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=start_row)
                start_row += len(df) + 2  # gap between stacked tables

    print(f"[pdfplumber] Converted: {input_path} -> {output_path} "
          f"({len(page_frames)} sheet(s) with tables)")
    return output_path


def pdf_to_word(input_path, output_path=None):
    input_path = os.path.abspath(input_path)
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"Input file not found: {input_path}")

    if output_path is None:
        output_path = os.path.splitext(input_path)[0] + ".docx"
    output_path = os.path.abspath(output_path)
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    from pdf2docx import Converter

    cv = Converter(input_path)
    try:
        cv.convert(output_path)
    finally:
        cv.close()

    print(f"[pdf2docx] Converted: {input_path} -> {output_path}")
    return output_path


def main():
    parser = argparse.ArgumentParser(description="Convert between Word (.docx) and PDF.")
    parser.add_argument(
        "mode",
        choices=["to-pdf", "to-word", "to-excel", "from-excel", "auto"],
        help=(
            "to-pdf: docx->pdf | to-word: pdf->docx | "
            "to-excel: pdf->xlsx | from-excel: xlsx->pdf | "
            "auto: infer from input extension (docx->pdf, pdf->docx, xlsx->pdf)"
        ),
    )
    parser.add_argument("input", help="Path to the input file")
    parser.add_argument("output", nargs="?", default=None, help="Path to the output file (optional)")
    args = parser.parse_args()

    mode = args.mode
    if mode == "auto":
        ext = os.path.splitext(args.input)[1].lower()
        if ext == ".docx":
            mode = "to-pdf"
        elif ext == ".xlsx":
            mode = "from-excel"
        elif ext == ".pdf":
            # A PDF could go to either Word or Excel; default to Word.
            # Use 'to-excel' explicitly if you want a spreadsheet instead.
            mode = "to-word"
        else:
            print(f"Cannot infer direction from extension '{ext}'. "
                  f"Use 'to-pdf', 'to-word', 'to-excel', or 'from-excel' explicitly.",
                  file=sys.stderr)
            sys.exit(1)

    try:
        if mode == "to-pdf":
            word_to_pdf(args.input, args.output)
        elif mode == "to-word":
            pdf_to_word(args.input, args.output)
        elif mode == "to-excel":
            pdf_to_excel(args.input, args.output)
        elif mode == "from-excel":
            excel_to_pdf(args.input, args.output)
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()