"""
excel_to_pdf_chainagewise.py

Chainage-wise Road/Revenue Area Statement -> PDF pipeline.

Given an Excel file with a "Road Chainage in Meters" section (columns
"From" / "To"), this script:

    1.  Detects the header rows automatically.
    2.  Splits the data into chainage-range blocks (consecutive rows that
        share the same From-To pair).
    3.  For every chainage block, rebuilds a standalone workbook that keeps
        the ORIGINAL sheet's formatting -- merged cells, fonts, borders,
        fills, column widths, page setup, print titles and print area --
        and contains only that block's data rows. Column widths are always
        copied exactly as stored in the source sheet and are NEVER
        changed. Every data row's height is instead forced to a FIXED
        value (FIXED_ROW_HEIGHT, 25pt) so wrap-text columns behave
        consistently everywhere. If a wrap-text cell (e.g. in the
        free-text 'Name as in RTC' column) has more text than fits within
        WRAP_TEXT_MAX_LINES (2) lines at that column's existing width and
        the fixed row height, the TEXT is shortened (cut cleanly at a
        word/comma boundary, no '...' marker) and top-aligned, so what's
        shown is always the first complete lines -- never a
        vertically-centered slice clipped from the middle of overflowing
        text.
    4.  Converts every per-chainage workbook to PDF using whichever
        conversion backend is available on the machine:
          - LibreOffice ('soffice', headless) -- works on Windows/Mac/Linux.
          - Microsoft Excel via COM automation ('win32com') -- Windows only,
            used automatically if LibreOffice is not installed but Excel is.
        Either way, this step preserves the exact look of the Excel sheet
        (colours, borders, multi-row headers, etc.) instead of re-drawing
        the table from scratch.
    5.  For chainage blocks with 20 Sl.No. rows or fewer ONLY:
        a.  Crops the converted PDF down to just the printed table (no
            distortion, aspect ratio preserved) and saves it in
            `output/cropped/` as:
                1_cropped_Chainage_<from>_to_<to>.pdf
                2_cropped_Chainage_<from>_to_<to>.pdf
                ...
        b.  Merges the (possibly multi-page, one page per column-block)
            table into a single, tightly packed A3-landscape page and
            saves it in `output/merged/` as:
                1.pdf
                2.pdf
                ...
            with the numbering following chainage order.
        Chainage blocks with MORE than 20 rows are left untouched after
        step 4 -- no cropped copy and no merged copy are produced for
        them, only the main per-chainage PDF.

Usage:
    python excel_to_pdf_chainagewise.py "input_file.xlsx" "output_folder"

Requirements:
    pip install openpyxl pymupdf

    Plus ONE of the following two conversion backends:

      Option A -- LibreOffice (recommended, cross-platform):
          Ubuntu/Debian : sudo apt-get install libreoffice
          macOS         : brew install --cask libreoffice
          Windows       : install LibreOffice and ensure 'soffice.exe' is
                           on PATH.

      Option B -- Microsoft Excel + pywin32 (Windows only):
          Requires Microsoft Excel to already be installed and licensed on
          the machine, plus:
              pip install pywin32
          The script automatically falls back to this option if
          LibreOffice is not found and the machine is running Windows.

    If NEITHER option is available, the script stops immediately with an
    explanation of what to install, before doing any real work.
"""

import argparse
import shutil
import subprocess
import sys
from copy import copy
from pathlib import Path

import fitz  # PyMuPDF
from openpyxl import Workbook, load_workbook


# --------------------------------------------------------------------------
# Configuration
# --------------------------------------------------------------------------

# A chainage block with this many Sl.No. rows or fewer gets a cropped copy
# AND a merged copy produced for it. Blocks with MORE rows are left
# completely alone after the initial conversion -- no cropping, no merging,
# just the main per-chainage PDF exactly as the converter rendered it.
ROW_COUNT_CROP_THRESHOLD = 20

# Every data row is forced to this exact height (in points), regardless of
# whatever height the source sheet had. Combined with WRAP_TEXT_MAX_LINES
# below, this guarantees the "Name as in RTC" (and any other wrap-text)
# column always shows exactly up to 2 full, clean lines -- never a
# vertically-centered, clipped slice from the middle of overflowing text.
FIXED_ROW_HEIGHT = 25
WRAP_TEXT_MAX_LINES = 2

# Layout constants used when cropping / merging pages (in PDF points).
PAGE_MARGIN = 15          # outer margin kept around content on every page
SL_NO_MARGIN = 30         # small extra space reserved on the left (Sl.No. side)
CONTENT_DETECTION_MARGIN = 8  # padding added around auto-detected content bbox


# --------------------------------------------------------------------------
# Excel helpers: locating headers, grouping rows, rebuilding per-block sheets
# --------------------------------------------------------------------------

def find_chainage_header(ws, max_scan_rows=15):
    """Locate the header rows and the 'From' / 'To' chainage columns.

    Returns (header_last_row, from_col, to_col), where header_last_row is
    the last row that is still part of the header block (the row right
    before actual data starts).
    """
    from_col = to_col = header_row = None

    for row in range(1, max_scan_rows + 1):
        for col in range(1, ws.max_column + 1):
            val = ws.cell(row=row, column=col).value
            if isinstance(val, str) and val.strip().lower() == "from":
                next_val = ws.cell(row=row, column=col + 1).value
                if isinstance(next_val, str) and next_val.strip().lower() == "to":
                    header_row, from_col, to_col = row, col, col + 1
                    break
        if header_row:
            break

    if header_row is None:
        raise ValueError(
            f"Could not find 'From' / 'To' chainage header cells in the "
            f"first {max_scan_rows} rows. Please check the sheet layout."
        )

    # The header may continue for a row or two more (e.g. unit sub-headers
    # like Acres/Guntas). The first data row is the first row after
    # header_row whose column A holds a number (the 'Sl No.' value).
    data_start_row = header_row + 1
    while data_start_row <= ws.max_row:
        if isinstance(ws.cell(row=data_start_row, column=1).value, (int, float)):
            break
        data_start_row += 1

    header_last_row = data_start_row - 1
    return header_last_row, from_col, to_col


def group_rows_by_chainage(ws, header_last_row, from_col, to_col):
    """Group data rows into consecutive blocks that share the same
    (From, To) chainage pair.

    Returns a list of (from_val, to_val, [row_indices]) tuples, in the
    order they appear in the sheet.
    """
    data_start_row = header_last_row + 1
    groups = []
    current_key = None
    current_rows = []

    for row in range(data_start_row, ws.max_row + 1):
        # Skip fully blank rows.
        if all(
            ws.cell(row=row, column=c).value in (None, "")
            for c in range(1, ws.max_column + 1)
        ):
            continue

        from_val = ws.cell(row=row, column=from_col).value
        to_val = ws.cell(row=row, column=to_col).value
        key = (from_val, to_val)

        if key != current_key and (from_val is not None or to_val is not None):
            if current_rows:
                groups.append((current_key[0], current_key[1], current_rows))
            current_key = key
            current_rows = [row]
        else:
            current_rows.append(row)

    if current_rows:
        groups.append((current_key[0], current_key[1], current_rows))

    return groups


def copy_cell(src_cell, dst_cell):
    """Copy a cell's value and full style (font, border, fill, number
    format, protection, alignment) from src to dst."""
    dst_cell.value = src_cell.value
    if src_cell.has_style:
        dst_cell.font = copy(src_cell.font)
        dst_cell.border = copy(src_cell.border)
        dst_cell.fill = copy(src_cell.fill)
        dst_cell.number_format = src_cell.number_format
        dst_cell.protection = copy(src_cell.protection)
        dst_cell.alignment = copy(src_cell.alignment)


def copy_sheet_setup(src_ws, dst_ws):
    """Copy page setup, print titles/area, and column widths so the
    generated sheet prints exactly like the source sheet."""
    dst_ws.page_setup.orientation = src_ws.page_setup.orientation
    dst_ws.page_setup.paperSize = src_ws.page_setup.paperSize
    dst_ws.page_setup.fitToWidth = src_ws.page_setup.fitToWidth
    dst_ws.page_setup.fitToHeight = src_ws.page_setup.fitToHeight
    dst_ws.page_setup.scale = src_ws.page_setup.scale
    dst_ws.page_setup.pageOrder = src_ws.page_setup.pageOrder

    dst_ws.sheet_properties.pageSetUpPr.fitToPage = (
        src_ws.sheet_properties.pageSetUpPr.fitToPage
    )

    dst_ws.print_options.horizontalCentered = src_ws.print_options.horizontalCentered
    dst_ws.print_options.verticalCentered = src_ws.print_options.verticalCentered
    dst_ws.page_margins = copy(src_ws.page_margins)

    # Repeat the identifying columns (Sl No. ... Name as in RTC) and the
    # header rows on every printed column-block, exactly like the source.
    if src_ws.print_title_cols:
        dst_ws.print_title_cols = src_ws.print_title_cols
    if src_ws.print_title_rows:
        dst_ws.print_title_rows = src_ws.print_title_rows

    if src_ws.print_area:
        dst_ws.print_area = src_ws.print_area

    for col_letter, dim in src_ws.column_dimensions.items():
        new_dim = dst_ws.column_dimensions[col_letter]
        new_dim.width = dim.width
        new_dim.hidden = dim.hidden


def copy_header_rows(src_ws, dst_ws, header_last_row):
    """Copy the header rows (values, styles, row heights, merged cells)."""
    for row in range(1, header_last_row + 1):
        if src_ws.row_dimensions[row].height:
            dst_ws.row_dimensions[row].height = src_ws.row_dimensions[row].height
        for col in range(1, src_ws.max_column + 1):
            copy_cell(src_ws.cell(row=row, column=col), dst_ws.cell(row=row, column=col))

    for merged_range in src_ws.merged_cells.ranges:
        if merged_range.max_row <= header_last_row:
            dst_ws.merge_cells(str(merged_range))


def find_wrap_columns(ws, header_last_row):
    """Return the set of column letters where data-row cells have
    wrap_text enabled (e.g. the free-text 'Name as in RTC' column). These
    are the columns whose content can overflow the sheet's fixed row
    height if there's too much text for the column's fixed width."""
    wrap_cols = set()
    for row in ws.iter_rows(min_row=header_last_row + 1, max_row=ws.max_row):
        for cell in row:
            if cell.alignment and cell.alignment.wrap_text:
                wrap_cols.add(cell.column_letter)
    return wrap_cols


def _measure_wrapped_height(text, font, width, temp_dir):
    """Render a single bordered cell with the given text/font/column width
    (wrap_text on, no fixed row height) and measure how tall LibreOffice
    auto-fits it -- i.e. how much vertical room this exact text actually
    needs at that column width. Returns None if it can't be measured."""
    from openpyxl.styles import Alignment, Border, Side

    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )

    probe_wb = Workbook()
    probe_ws = probe_wb.active
    probe_ws.column_dimensions["A"].width = width

    cell = probe_ws.cell(row=1, column=1)
    cell.value = text
    cell.font = copy(font)
    cell.alignment = Alignment(wrap_text=True, vertical="top")
    cell.border = thin_border
    probe_ws.row_dimensions[1].height = None  # let LibreOffice auto-fit it

    probe_xlsx = temp_dir / "_probe_wrap.xlsx"
    probe_wb.save(probe_xlsx)
    probe_pdf_path = convert_xlsx_to_pdf(probe_xlsx, temp_dir)

    probe_doc = fitz.open(probe_pdf_path)
    page = probe_doc[0]
    hlines = sorted(
        {round(dr["rect"].y0, 1) for dr in page.get_drawings() if dr["rect"].height < 0.5}
    )
    probe_doc.close()
    probe_xlsx.unlink()
    probe_pdf_path.unlink()

    if len(hlines) >= 2:
        return hlines[1] - hlines[0]
    return None


def _truncate_to_fit(text, max_chars):
    """Shorten `text` to at most `max_chars`, cutting at the last comma or
    space before the limit (so a name/word is never split mid-way). No
    '...' or other marker is added -- it's just cut cleanly."""
    if len(text) <= max_chars:
        return text

    truncated = text[:max_chars]
    cut_at = max(truncated.rfind(","), truncated.rfind(" "))
    if cut_at > max_chars * 0.5:  # only back off to a word boundary if it
        truncated = truncated[:cut_at]  # doesn't throw away too much text
    return truncated.rstrip(", ")


def compute_wrap_text_truncations(
    src_ws,
    group_rows,
    wrap_columns,
    temp_dir,
    max_lines=WRAP_TEXT_MAX_LINES,
    line_spacing=1.3,
):
    """Every data row is forced to FIXED_ROW_HEIGHT (see build_group_workbook),
    so wrap-text columns must fit within exactly `max_lines` lines at that
    fixed height and the column's own EXISTING width. This works out how
    much text actually fits and shortens any cell that has more than that,
    cutting cleanly at a word/comma boundary -- so it displays as complete,
    top-aligned lines instead of an overflowing, vertically-clipped mess.

    Returns {(row, col_letter): truncated_text}, only for cells that
    actually need shortening.
    """
    if not wrap_columns:
        return {}

    from openpyxl.utils import column_index_from_string

    truncations = {}

    for col_letter in wrap_columns:
        col_idx = column_index_from_string(col_letter)
        current_width = src_ws.column_dimensions[col_letter].width or 8.43

        # Measure once, using the LONGEST text in this column, to derive an
        # accurate "characters per line at this exact column width" ratio
        # -- then reuse that ratio for every row in the group without
        # having to re-render each one individually.
        longest_row = max(
            group_rows,
            key=lambda r: len(str(src_ws.cell(row=r, column=col_idx).value or "")),
        )
        longest_value = src_ws.cell(row=longest_row, column=col_idx).value
        if longest_value is None or longest_value == "":
            continue
        longest_text = str(longest_value)  # guard against non-string cell values

        font = src_ws.cell(row=longest_row, column=col_idx).font
        font_size = font.size or 10
        line_height = font_size * line_spacing

        measured_height = _measure_wrapped_height(
            longest_text, font, current_width, temp_dir
        )
        if measured_height is None:
            continue

        lines_for_longest = max(1, round(measured_height / line_height))
        chars_per_line = max(1, len(longest_text) / lines_for_longest)

        # Every row is held to the SAME fixed height and the SAME max line
        # count -- not each row's own (now-irrelevant) source height.
        # A small safety margin (5%) so the estimate leans conservative --
        # slightly under-filling 2 lines is fine, slightly overflowing them
        # is what causes the clipped/garbled look.
        max_chars = int(max_lines * chars_per_line * 0.95)

        for row in group_rows:
            text = src_ws.cell(row=row, column=col_idx).value
            if not isinstance(text, str) or not text:
                continue

            if len(text) > max_chars:
                truncations[(row, col_letter)] = _truncate_to_fit(text, max_chars)

    return truncations


def build_group_workbook(src_ws, header_last_row, group_rows, wrap_text_truncations=None):
    """Build a standalone workbook containing the header rows plus only the
    data rows belonging to one chainage block, preserving all formatting.

    Every data row's height is forced to FIXED_ROW_HEIGHT (not copied from
    the source), so wrap-text columns behave consistently everywhere.
    `wrap_text_truncations` (optional) maps (row, col_letter) to a
    shortened version of that cell's text -- for wrap-text cells whose full
    text would need more than WRAP_TEXT_MAX_LINES lines to display at that
    fixed height (see compute_wrap_text_truncations). ANY wrap-text cell
    (truncated or not) is also switched to top vertical alignment, so what's
    shown is always the first clean lines of text, never a
    vertically-centered slice clipped from the middle of overflowing text.
    """
    wrap_text_truncations = wrap_text_truncations or {}
    wb = Workbook()
    ws = wb.active
    ws.title = src_ws.title

    copy_sheet_setup(src_ws, ws)
    copy_header_rows(src_ws, ws, header_last_row)

    dst_row = header_last_row + 1
    for src_row in group_rows:
        ws.row_dimensions[dst_row].height = FIXED_ROW_HEIGHT
        for col in range(1, src_ws.max_column + 1):
            src_cell = src_ws.cell(row=src_row, column=col)
            dst_cell = ws.cell(row=dst_row, column=col)
            copy_cell(src_cell, dst_cell)

            if dst_cell.alignment and dst_cell.alignment.wrap_text:
                # ANY wrap-text cell (truncated or not) is top-aligned, so
                # if text still overflows the fixed row height for any
                # reason, it clips from the bottom only -- the visible
                # text is always the first complete lines, never a
                # vertically-centered slice clipped top and bottom.
                align = copy(dst_cell.alignment)
                align.vertical = "top"
                dst_cell.alignment = align

            override = wrap_text_truncations.get((src_row, src_cell.column_letter))
            if override is not None:
                dst_cell.value = override
        dst_row += 1

    return wb


# --------------------------------------------------------------------------
# Formatting / naming helpers
# --------------------------------------------------------------------------

def format_chainage_value(val):
    """Format a chainage value (e.g. 0.18) the same way regardless of
    whether openpyxl read it back as a number or a string."""
    if isinstance(val, (int, float)):
        return f"{val:.2f}"
    return str(val).strip()


def safe_filename(text):
    """Strip characters that are unsafe in file names."""
    keep = "-_.() "
    return "".join(c if c.isalnum() or c in keep else "_" for c in text)


# --------------------------------------------------------------------------
# Excel -> PDF conversion backends (LibreOffice, or Excel via win32com)
# --------------------------------------------------------------------------

# Cached once detected, so we don't repeat the detection work for every
# chainage block.
_CONVERSION_BACKEND = None  # "libreoffice" | "win32com" | None


def detect_conversion_backend():
    """Work out which tool is available to do the xlsx -> PDF conversion.

    Preference order:
      1. LibreOffice ('soffice') -- cross-platform, no license needed.
      2. Microsoft Excel via COM automation ('win32com') -- Windows only,
         requires Excel to already be installed and licensed, plus the
         'pywin32' package.

    Returns "libreoffice", "win32com", or None if neither is usable.
    """
    global _CONVERSION_BACKEND
    if _CONVERSION_BACKEND is not None:
        return _CONVERSION_BACKEND

    if shutil.which("soffice") is not None:
        _CONVERSION_BACKEND = "libreoffice"
        return _CONVERSION_BACKEND

    if sys.platform.startswith("win"):
        try:
            import win32com.client  # noqa: F401
            _CONVERSION_BACKEND = "win32com"
            return _CONVERSION_BACKEND
        except ImportError:
            pass

    _CONVERSION_BACKEND = None
    return _CONVERSION_BACKEND


def check_conversion_backend_available():
    """Verify that a usable xlsx -> PDF conversion backend exists. Raises a
    clear, actionable error immediately if neither is available -- rather
    than letting the script fail deep inside the conversion loop after it
    has already done a lot of Excel processing."""
    backend = detect_conversion_backend()

    if backend is None:
        raise EnvironmentError(
            "No Excel-to-PDF conversion tool was found on this machine.\n"
            "This script needs ONE of the following to convert each "
            "chainage workbook to PDF:\n\n"
            "  Option A -- LibreOffice (recommended, works on any OS):\n"
            "    Ubuntu/Debian : sudo apt-get install libreoffice\n"
            "    macOS         : brew install --cask libreoffice\n"
            "    Windows       : install LibreOffice from "
            "https://www.libreoffice.org/download/download/ and make sure "
            "'soffice.exe' is on your PATH.\n\n"
            "  Option B -- Microsoft Excel (Windows only):\n"
            "    Requires Excel to already be installed and licensed, plus:\n"
            "      pip install pywin32\n\n"
            "Install one of the two, then run this script again."
        )

    print(f"Using conversion backend: {backend}")


# Backward-compatible alias: earlier versions of this script only supported
# LibreOffice, so the startup check was named `check_libreoffice_available`.
# The check now covers both LibreOffice and the Excel/win32com fallback, but
# the old name is kept as an alias so existing imports elsewhere in the
# project (e.g. `from src.core.excel_to_pdf_chainagewise_4 import
# check_libreoffice_available`) keep working without modification.
check_libreoffice_available = check_conversion_backend_available


def convert_xlsx_to_pdf(xlsx_path, out_dir):
    """Convert a single .xlsx file to PDF using whichever backend is
    available (LibreOffice or Excel via win32com). Returns the path to the
    resulting PDF."""
    backend = detect_conversion_backend()

    if backend == "libreoffice":
        return _convert_with_libreoffice(xlsx_path, out_dir)
    elif backend == "win32com":
        return _convert_with_win32com(xlsx_path, out_dir)
    else:
        # Should not normally happen if check_conversion_backend_available()
        # was called first, but guard against it anyway.
        check_conversion_backend_available()
        raise EnvironmentError("No conversion backend available.")


def _convert_with_libreoffice(xlsx_path, out_dir):
    """Convert xlsx -> PDF using headless LibreOffice ('soffice')."""
    cmd = [
        "soffice", "--headless",
        "--convert-to", "pdf",
        "--outdir", str(out_dir),
        str(xlsx_path),
    ]
    result = subprocess.run(cmd, capture_output=True, text=True)

    if result.returncode != 0:
        raise RuntimeError(
            f"LibreOffice conversion failed for {Path(xlsx_path).name}:\n"
            f"{result.stdout}\n{result.stderr}"
        )

    pdf_path = Path(out_dir) / (Path(xlsx_path).stem + ".pdf")
    if not pdf_path.exists():
        raise RuntimeError(
            f"Conversion command ran but no PDF was produced for "
            f"{Path(xlsx_path).name}.\nLibreOffice output:\n{result.stdout}"
        )

    return pdf_path


def _convert_with_win32com(xlsx_path, out_dir):
    """Convert xlsx -> PDF by driving Microsoft Excel through COM
    automation. Windows-only; requires Excel to be installed and licensed,
    plus the 'pywin32' package (`pip install pywin32`)."""
    import pythoncom
    import win32com.client

    xlsx_path = Path(xlsx_path).resolve()
    pdf_path = Path(out_dir).resolve() / (xlsx_path.stem + ".pdf")

    # xlTypePDF = 0 (the fixed-format type Excel uses for PDF export)
    XL_TYPE_PDF = 0

    # COM requires its own initialization on the thread that uses it.
    pythoncom.CoInitialize()
    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        workbook = excel.Workbooks.Open(str(xlsx_path))
        workbook.ExportAsFixedFormat(XL_TYPE_PDF, str(pdf_path))
    except Exception as e:
        raise RuntimeError(
            f"Excel (win32com) conversion failed for {xlsx_path.name}: {e}"
        )
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        if excel is not None:
            excel.Quit()
        pythoncom.CoUninitialize()

    if not pdf_path.exists():
        raise RuntimeError(
            f"Excel (win32com) ran but no PDF was produced for {xlsx_path.name}."
        )

    return pdf_path


# --------------------------------------------------------------------------
# PDF post-processing (PyMuPDF): whitespace detection, cropping, merging
# --------------------------------------------------------------------------

def get_content_rects(pdf_path, margin=CONTENT_DETECTION_MARGIN):
    """Open a PDF and return (doc, [content_rect_per_page]).

    content_rect is the bounding box of the actual printed table/text on
    that page (with a small margin added around it), i.e. everything
    excluding blank page whitespace.
    """
    doc = fitz.open(pdf_path)
    rects = []

    for page in doc:
        content_rect = fitz.Rect()

        for block in page.get_text("blocks"):
            content_rect |= fitz.Rect(block[:4])
        for drawing in page.get_drawings():
            content_rect |= drawing["rect"]

        if content_rect.is_empty:
            content_rect = page.rect  # nothing detected: fall back to full page

        content_rect.x0 = max(page.rect.x0, content_rect.x0 - margin)
        content_rect.y0 = max(page.rect.y0, content_rect.y0 - margin)
        content_rect.x1 = min(page.rect.x1, content_rect.x1 + margin)
        content_rect.y1 = min(page.rect.y1, content_rect.y1 + margin)
        rects.append(content_rect)

    return doc, rects


def crop_pdf_to_a3(input_pdf, output_pdf, page_margin=PAGE_MARGIN):
    """Crop every page down to just its printed table content, then place
    that content (scaled up as much as possible WITHOUT distorting its
    proportions) onto a fresh A3-landscape page. The physical page size
    stays A3 throughout, only the wasted blank space is removed."""
    src_doc, rects = get_content_rects(input_pdf)
    a3 = src_doc[0].rect  # source pages are already A3 from LibreOffice

    new_doc = fitz.open()
    for i, content_rect in enumerate(rects):
        page = new_doc.new_page(width=a3.width, height=a3.height)

        avail_w = a3.width - 2 * page_margin
        avail_h = a3.height - 2 * page_margin

        # Single uniform scale -- fills as much of the page as possible
        # without stretching the table out of proportion.
        scale = min(avail_w / content_rect.width, avail_h / content_rect.height)
        w = content_rect.width * scale
        h = content_rect.height * scale

        # Center the (undistorted) content in the available space.
        x0 = page_margin + (avail_w - w) / 2
        y0 = page_margin + (avail_h - h) / 2
        target = fitz.Rect(x0, y0, x0 + w, y0 + h)

        page.show_pdf_page(target, src_doc, i, clip=content_rect)  # keep_proportion=True

    new_doc.save(output_pdf)
    new_doc.close()
    src_doc.close()


def merge_pages_to_single(
    input_pdf,
    output_pdf,
    sl_no_margin=SL_NO_MARGIN,
    page_margin=PAGE_MARGIN,
):
    """Stack every page's table content vertically and fit the whole thing
    onto ONE single A3-landscape page.

    - A single uniform scale factor is used for every block, so nothing is
      stretched or distorted -- proportions match the original Excel sheet.
    - Every block is placed at the exact same x0 and given the exact same
      target width (based on the widest block). This guarantees the left
      and right edges of every stacked section line up perfectly with each
      other instead of drifting a few points left/right (which is what
      caused the 'misaligned boxes' look when each block was centered
      independently based on its own slightly different natural width).
    - Left-over vertical space (whatever is left after the table content is
      scaled to fit) is split EVENLY across the top margin, every gap
      between stacked blocks, and the bottom margin. This means a chainage
      block with fewer Sl.No. rows -- and therefore more left-over space --
      gets proportionally MORE whitespace everywhere, including the gap(s)
      between sections, instead of a fixed small gap. The result reads as
      one clean, evenly-balanced page rather than a stitched-together one.
    - A small amount of extra space is reserved on the left (Sl.No. side)
      of the table, per the requested layout.
    """
    src_doc, rects = get_content_rects(input_pdf)
    a3 = src_doc[0].rect

    comp_width = max(r.width for r in rects)
    raw_heights = [r.height for r in rects]

    avail_w = a3.width - 2 * page_margin - sl_no_margin
    # Reserve at least `page_margin` at the very top/bottom when computing
    # the scale, so we never end up with zero (or negative) outer margin.
    avail_h_for_scale = a3.height - 2 * page_margin

    # Single uniform scale for the whole composite -- keeps every block's
    # proportions correct and keeps every block's edges aligned.
    scale = min(avail_w / comp_width, avail_h_for_scale / sum(raw_heights))
    scaled_heights = [h * scale for h in raw_heights]

    # Evenly distribute ALL left-over vertical space (top margin + every
    # inter-block gap + bottom margin) across (n_blocks + 1) equal slots.
    n_gaps = len(rects) + 1
    leftover = a3.height - sum(scaled_heights)
    gap = leftover / n_gaps

    # Every block shares the SAME horizontal position and width (based on
    # the widest block), centered once as a whole -- not per block.
    target_w = comp_width * scale
    x0 = page_margin + sl_no_margin + (avail_w - target_w) / 2

    new_doc = fitz.open()
    page = new_doc.new_page(width=a3.width, height=a3.height)

    y = gap
    for i, r in enumerate(rects):
        h = scaled_heights[i]
        target = fitz.Rect(x0, y, x0 + target_w, y + h)
        # keep_proportion=False here is safe: block widths differ from each
        # other by well under 1% naturally, so forcing them to the exact
        # same target width causes no visible distortion, and guarantees
        # every block's left/right edges align exactly with every other.
        page.show_pdf_page(target, src_doc, i, clip=r, keep_proportion=False)
        y += h + gap

    new_doc.save(output_pdf)
    new_doc.close()
    src_doc.close()


# --------------------------------------------------------------------------
# Main pipeline
# --------------------------------------------------------------------------

def process_chainage_group(
    index,
    from_val,
    to_val,
    group_rows,
    src_ws,
    header_last_row,
    output_dir,
    cropped_dir,
    merged_dir,
    temp_dir,
):
    """Run the full pipeline for a single chainage block:
    build workbook -> convert to PDF -> (crop + merge, only if <= 20 rows).

    `index` is the 1-based chainage order, used for file naming.

    Chainage blocks with MORE than ROW_COUNT_CROP_THRESHOLD (20) Sl.No.
    rows are not cropped or repacked -- but a plain copy of the as-exported
    PDF is still placed in merged/ (under its original exported filename)
    so every chainage block has a file present there.

    Returns (main_pdf_path, cropped_pdf_path, merged_pdf_path).
    For blocks with more than 20 rows, cropped_pdf_path is None and
    merged_pdf_path points to the as-exported copy (not a repacked file).
    """
    from_str = format_chainage_value(from_val)
    to_str = format_chainage_value(to_val)
    chainage_label = safe_filename(f"Chainage_{from_str}_to_{to_str}")
    # Every exported file for this block -- the main per-chainage PDF, and
    # (for >20 row blocks) its as-exported copy in merged/ -- carries this
    # same "<index>_" prefix, e.g. "1_Chainage_0.00_to_0.25".
    exported_label = f"{index}_{chainage_label}"

    # 1. Neither row height (now fixed) nor column width are changed by
    #    guesswork -- work out, using the wrap-text column's own EXISTING
    #    width and the FIXED row height, how much text actually fits within
    #    2 lines, and shorten any cell that has more than that. Then build
    #    the standalone workbook for this chainage block.
    wrap_columns = find_wrap_columns(src_ws, header_last_row)
    wrap_text_truncations = compute_wrap_text_truncations(
        src_ws, group_rows, wrap_columns, temp_dir
    )
    group_wb = build_group_workbook(
        src_ws, header_last_row, group_rows, wrap_text_truncations
    )
    temp_xlsx = temp_dir / f"{exported_label}.xlsx"
    group_wb.save(temp_xlsx)

    # 2. Convert it to PDF via LibreOffice / Excel (preserves all formatting).
    main_pdf_path = convert_xlsx_to_pdf(temp_xlsx, output_dir)

    row_count = len(group_rows)
    if row_count > ROW_COUNT_CROP_THRESHOLD:
        # More than 20 rows: skip cropping AND merging (no distortion/
        # repacking of these larger tables). A plain copy of the
        # as-exported PDF is still placed in the merged/ folder -- under
        # its original exported name, NOT the "<index>.pdf" naming used
        # for actually-merged files -- so every chainage block has some
        # file present in merged/ for downstream consumers.
        skipped_copy_path = merged_dir / main_pdf_path.name
        shutil.copy2(main_pdf_path, skipped_copy_path)
        return main_pdf_path, None, skipped_copy_path

    # 3. Crop whitespace (20 rows or fewer only).
    cropped_name = f"{index}_cropped_{chainage_label}.pdf"
    cropped_pdf_path = cropped_dir / cropped_name
    crop_pdf_to_a3(main_pdf_path, cropped_pdf_path)

    # 4. Merge all column-block pages into a single tightly packed A3 page.
    merged_pdf_path = merged_dir / f"{index}.pdf"
    merge_pages_to_single(main_pdf_path, merged_pdf_path)

    return main_pdf_path, cropped_pdf_path, merged_pdf_path


def split_workbook_by_chainage(input_file, output_folder):
    """Main entry point: read the Excel file, split it chainage-wise, and
    produce the main / cropped / merged PDFs for every chainage block."""
    input_path = Path(input_file).expanduser().resolve()
    output_dir = Path(output_folder).expanduser().resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    cropped_dir = output_dir / "cropped"
    merged_dir = output_dir / "merged"
    cropped_dir.mkdir(exist_ok=True)
    merged_dir.mkdir(exist_ok=True)

    temp_dir = output_dir / "_temp_xlsx"
    temp_dir.mkdir(exist_ok=True)

    wb = load_workbook(input_path, data_only=True)
    ws = wb.active

    header_last_row, from_col, to_col = find_chainage_header(ws)
    groups = group_rows_by_chainage(ws, header_last_row, from_col, to_col)

    results = []
    for index, (from_val, to_val, group_rows) in enumerate(groups, start=1):
        paths = process_chainage_group(
            index=index,
            from_val=from_val,
            to_val=to_val,
            group_rows=group_rows,
            src_ws=ws,
            header_last_row=header_last_row,
            output_dir=output_dir,
            cropped_dir=cropped_dir,
            merged_dir=merged_dir,
            temp_dir=temp_dir,
        )
        results.append(paths)

    # Clean up intermediate per-chainage xlsx files.
    for f in temp_dir.glob("*.xlsx"):
        f.unlink()
    temp_dir.rmdir()

    return results


# --------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description=(
            "Split a chainage-wise Excel area statement into per-chainage "
            "PDFs, crop whitespace, and merge each block into a single "
            "clean A3-landscape PDF."
        )
    )
    parser.add_argument("input_file", help="Path to the input Excel file")
    parser.add_argument("output_folder", help="Folder where results should be saved")
    args = parser.parse_args()

    try:
        check_conversion_backend_available()
        results = split_workbook_by_chainage(args.input_file, args.output_folder)
        print(f"Processed {len(results)} chainage block(s):")
        for i, (main_pdf, cropped_pdf, merged_pdf) in enumerate(results, start=1):
            if cropped_pdf is None:
                print(f"  {i}. more than {ROW_COUNT_CROP_THRESHOLD} rows -- "
                      f"kept as-is -> {main_pdf}")
                print(f"     merged  -> {merged_pdf}  (as-exported copy, not repacked)")
            else:
                print(f"  {i}. cropped -> {cropped_pdf}")
                print(f"     merged  -> {merged_pdf}")
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
