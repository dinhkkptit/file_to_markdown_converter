# Prompt Guide
You are a senior Python engineer. Write a robust CLI script named `excel_sheets_to_markdown.py` that converts files from `input/` to Markdown in `output/`.

Rules:
- Scan `input/` recursively; write outputs into `output/`. Default args: `python excel_sheets_to_markdown.py input output` (defaults to `input`/`output` if omitted).
- Support: `.xlsx`, `.csv`, `.txt`, `.docx`, `.pdf`.
- `.xlsx`:
  - If multiple sheets → `output/<xlsx_stem>/<sheet_name>.md` (one file per sheet).
  - If single sheet → `output/<xlsx_stem>.md`.
  - Convert each sheet to a Markdown table (preserve values as strings; empty cells become empty).
- `.csv`: like single-sheet workbook → `output/<csv_stem>.md` as Markdown table.
- `.txt` and `.docx`: `output/<stem>.md` with heading and extracted text (simple paragraphs ok).
- `.pdf`: `output/<stem>.md` using text extraction only (PyMuPDF). If no extractable text, write a clear note like “No extractable text found (likely scanned PDF).”
- Slugify filenames/sheet names safely for Windows/Linux.
- Print OK/FAIL per file and continue on errors.

Dependencies: `pandas`, `openpyxl`, `tabulate`, `python-docx`, `pymupdf`. **Do NOT include OCR/pytesseract.**

Deliver: output the full script in one code block.
