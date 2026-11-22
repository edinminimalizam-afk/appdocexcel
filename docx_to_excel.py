#!/usr/bin/env python3
"""
docx_to_excel.py - Extract the content of a Word (.docx) file into an Excel file.
If the Excel file already exists, the content is added to the end of it.
If it doesn't exist, it is created.

The install command below only works if you have homebrew, python, and pip installed. let me know if
you'd like some help with that or you can find some guides online.


# HOW TO USE VIA THE CLI

Install:
    pip install python-docx openpyxl

Usage:
    python docx_to_excel.py my_document.docx output.xlsx # specify the file you want the contentextracted to
    python docx_to_excel.py my_document.docx            # writes to extracted_content.xlsx if no file name specified

    python docx_to_excel.py                             # pick both files in dialogs

"""

import sys
from datetime import datetime
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font

DEFAULT_OUTPUT = "extracted_content.xlsx"
SHEET_NAME = "Content"
HEADERS = ["File", "Extracted At", "#", "Type", "Style", "Content"]


class ExtractError(Exception):
    pass


# --------------------------------------------------------------------------- #
# Extraction
# --------------------------------------------------------------------------- #
def read_docx(path):
    """Yield (type, style, text) for each paragraph and table row, in document order."""
    doc = Document(path)
    for child in doc.element.body.iterchildren():
        if child.tag == qn("w:p"):
            para = Paragraph(child, doc)
            text = para.text.strip()
            if text:
                style = para.style.name if para.style is not None else ""
                kind = (
                    "Heading" if style.startswith(("Heading", "Title")) else "Paragraph"
                )
                yield kind, style, text
        elif child.tag == qn("w:tbl"):
            for row in Table(child, doc).rows:
                cells, seen = [], set()
                for cell in row.cells:  # merged cells repeat; keep each once
                    if id(cell._tc) not in seen:
                        seen.add(id(cell._tc))
                        cells.append(cell.text.strip())
                if any(cells):
                    yield "Table row", "", " | ".join(cells)


def open_workbook(path):
    """Open the Excel file if it exists, or create it with a header row."""
    if Path(path).exists():
        wb = load_workbook(path)
        if SHEET_NAME in wb.sheetnames:
            return wb, wb[SHEET_NAME]
        ws = wb.create_sheet(SHEET_NAME)
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = SHEET_NAME

    ws.append(HEADERS)
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for col, width in zip("ABCDEF", [30, 20, 6, 12, 16, 100]):
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "A2"
    return wb, ws


def extract(input_path, output_path):
    """Append the Word file's content to the Excel file. Returns a status message."""
    if not input_path or not Path(input_path).is_file():
        raise ExtractError(f"File not found: {input_path}")
    if Path(input_path).suffix.lower() != ".docx":
        raise ExtractError(
            "Only .docx files are supported. Save .doc or Google Docs as .docx first."
        )
    if not str(output_path).lower().endswith(".xlsx"):
        output_path = f"{output_path}.xlsx"

    wb, ws = open_workbook(output_path)
    file_name = Path(input_path).name
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M")

    count = 0
    for count, (kind, style, text) in enumerate(read_docx(input_path), start=1):
        ws.append([file_name, timestamp, count, kind, style, text])
        ws.cell(row=ws.max_row, column=6).alignment = Alignment(
            wrap_text=True, vertical="top"
        )

    try:
        wb.save(output_path)
    except PermissionError:
        raise ExtractError(
            f"Can't write {output_path}. Close it in Excel and try again."
        )

    return f"Added {count} rows from {file_name} to {output_path}"


# --------------------------------------------------------------------------- #
# Dialog mode (double-clicked app, or no arguments)
# --------------------------------------------------------------------------- #
def run_dialogs():
    import tkinter as tk
    from tkinter import filedialog, messagebox

    root = tk.Tk()
    root.withdraw()

    input_path = filedialog.askopenfilename(
        title="Step 1 of 2: Select the Word document",
        filetypes=[("Word documents", "*.docx"), ("All files", "*.*")],
    )
    if not input_path:
        return

    output_path = filedialog.asksaveasfilename(
        title="Step 2 of 2: Choose an existing Excel file or type a new name",
        initialdir=str(Path(input_path).parent),
        initialfile=DEFAULT_OUTPUT,
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")],
        confirmoverwrite=False,  # existing files are added to, not replaced
    )
    if not output_path:
        return

    try:
        messagebox.showinfo("Done", extract(input_path, output_path))
    except Exception as exc:
        messagebox.showerror("Something went wrong", str(exc))
    finally:
        root.destroy()


def main():
    if len(sys.argv) < 2:
        run_dialogs()
        return
    input_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else DEFAULT_OUTPUT
    try:
        print(extract(input_path, output_path))
    except ExtractError as exc:
        sys.exit(str(exc))


if __name__ == "__main__":
    main()


# ---------------------------------------------------------------------------
# HOW TO BUILD THE WINDOWS APP (DocxToExcel.exe)
# ---------------------------------------------------------------------------
# You only need Python on the PC that builds the .exe. People who use the
# .exe don't need Python.
#
# 1. Install Python from https://www.python.org/downloads/
#    On the first installer screen, tick "Add python.exe to PATH".
#
# 2. Open Command Prompt in the folder that has this file:
#    in File Explorer, click the address bar, type cmd and press Enter.
#
# 3. Install the libraries and PyInstaller (only needed once):
#      python -m pip install python-docx openpyxl pyinstaller
#
# 4. Build the app:
#      python -m PyInstaller --onefile --windowed --name DocxToExcel docx_to_excel.py
#
#    --onefile   packs everything into a single .exe
#    --windowed  stops a console window from opening (only the dialogs show)
#    Optional:   add --icon=myicon.ico to use your own icon
#
# 5. The app is created at dist\DocxToExcel.exe. That one file is all you
#    need to share. You can delete the build folder and DocxToExcel.spec.
#
# 6. Run it by double-clicking. Pick the Word file, then pick an existing
#    Excel file or type a new name. A popup shows how many rows were added.
#
# Notes:
#  - The first time it runs, Windows may show "Windows protected your PC",
#    because the app isn't signed. Click "More info" -> "Run anyway".
#  - Some antivirus programs flag new PyInstaller .exe files. If that
#    happens, add an exception or send the file inside a .zip.
#  - After changing this script, run step 4 again to rebuild the .exe.
#  - If the app fails without showing an error, build it without --windowed
#    and run dist\DocxToExcel.exe from Command Prompt to see the error.
#  - When built as an app/.exe and double-clicked, it runs in dialog mode:
#    pick the Word file, pick (or name) the Excel file, and a popup confirms the result.
# ---------------------------------------------------------------------------
