# Excel/CSV/TXT/DOCX/PDF → Markdown Converter (No OCR)

This small project converts files in `input/` into Markdown in `output/`.

## Supported formats
- **.xlsx**
  - **Multiple sheets** → `output/<xlsx_name>/<sheet>.md`
  - **Single sheet** → `output/<xlsx_name>.md`
  - Sheets are exported as **Markdown tables**
- **.csv**
  - Exported as **Markdown table** → `output/<csv_name>.md`
- **.txt**
  - Exported as text → `output/<name>.md`
- **.docx**
  - Paragraph text extracted → `output/<name>.md`
- **.pdf**
  - Text extraction only (no OCR) → `output/<name>.md`
  - If the PDF is scanned/image-based, the output will contain a note.

## Project structure
```
.
├─ input/                   # put your files here (can be nested)
├─ output/                  # generated markdown appears here
└─ markdown.py
```

## Install dependencies
```bash
pip install pandas openpyxl tabulate python-docx pymupdf
```

## Run
```bash
python markdown.py input output
python markdown.py "anysource" "anydest"
```

If you omit args, it defaults to `input/` and `output/`:
```bash
python markdown.py
```

## Notes
- Filenames and sheet names are **slugified** to be safe on Windows/Linux.
- The tool prints `OK` / `FAIL` for each file and continues even if one file fails.
