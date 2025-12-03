#!/usr/bin/env python3
import argparse
import os
import re
import sys
from pathlib import Path

import pandas as pd
from docx import Document

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None


def slugify(name: str, max_len: int = 120) -> str:
    """
    Safe filename component for Windows/Linux.
    """
    name = str(name).strip()
    name = name.replace(os.sep, "_").replace("/", "_").replace("\\", "_")
    name = re.sub(r"[^\w\-. ]+", "_", name, flags=re.UNICODE)
    name = re.sub(r"\s+", "_", name).strip("_")
    return (name[:max_len] or "untitled")


def write_text_md(out_path: Path, title: str, body: str) -> None:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    content = f"# {title}\n\n{body.rstrip()}\n"
    out_path.write_text(content, encoding="utf-8")


def df_to_markdown_table(df: pd.DataFrame) -> str:
    if df is None or (df.shape[0] == 0 and df.shape[1] == 0):
        return "_(Empty table)_\n"

    df = df.copy().fillna("")
    df.columns = ["" if c is None else str(c) for c in df.columns]

    try:
        return df.to_markdown(index=False)
    except Exception:
        # Fallback if tabulate has trouble
        return "```\n" + df.to_string(index=False) + "\n```"


def convert_xlsx(in_path: Path, out_dir: Path) -> None:
    base = slugify(in_path.stem)
    xl = pd.ExcelFile(in_path, engine="openpyxl")
    sheets = xl.sheet_names

    # Single sheet -> output/<file>.md
    if len(sheets) == 1:
        sheet = sheets[0]
        df = pd.read_excel(xl, sheet_name=sheet, dtype=str).fillna("")
        write_text_md(out_dir / f"{base}.md", title=in_path.name, body=df_to_markdown_table(df))
        return

    # Multiple sheets -> output/<file>/<sheet>.md ...
    book_out = out_dir / base
    book_out.mkdir(parents=True, exist_ok=True)
    for sheet in sheets:
        df = pd.read_excel(xl, sheet_name=sheet, dtype=str).fillna("")
        write_text_md(book_out / f"{slugify(sheet)}.md", title=sheet, body=df_to_markdown_table(df))


def convert_csv(in_path: Path, out_dir: Path) -> None:
    base = slugify(in_path.stem)
    df = pd.read_csv(
        in_path,
        dtype=str,
        keep_default_na=False,
        na_filter=False,
        encoding="utf-8",
        engine="python",
    )
    write_text_md(out_dir / f"{base}.md", title=in_path.name, body=df_to_markdown_table(df))


def convert_txt(in_path: Path, out_dir: Path) -> None:
    base = slugify(in_path.stem)
    text = in_path.read_text(encoding="utf-8", errors="replace")
    write_text_md(out_dir / f"{base}.md", title=in_path.name, body=text)


def convert_docx(in_path: Path, out_dir: Path) -> None:
    base = slugify(in_path.stem)
    doc = Document(str(in_path))
    parts = [p.text.rstrip() for p in doc.paragraphs]
    body = "\n\n".join(parts).strip() or "_(Empty document)_"
    write_text_md(out_dir / f"{base}.md", title=in_path.name, body=body)


def convert_pdf(in_path: Path, out_dir: Path) -> None:
    """
    PDF -> Markdown via text extraction ONLY (no OCR).
    If the PDF is scanned/image-based, output a clear note.
    """
    if fitz is None:
        raise RuntimeError("Missing dependency: pymupdf. Install with: pip install pymupdf")

    base = slugify(in_path.stem)
    doc = fitz.open(str(in_path))

    chunks = []
    for i in range(doc.page_count):
        page = doc.load_page(i)
        text = page.get_text("text").strip()
        if text:
            chunks.append(f"## Page {i+1}\n\n{text}")

    body = "\n\n".join(chunks).strip()
    if not body:
        body = "_(No extractable text found. This PDF may be scanned; OCR is not enabled.)_"

    write_text_md(out_dir / f"{base}.md", title=in_path.name, body=body)


def iter_inputs(input_dir: Path):
    exts = {".xlsx", ".csv", ".txt", ".docx", ".pdf"}
    for p in sorted(input_dir.rglob("*")):
        if p.is_file() and p.suffix.lower() in exts:
            yield p


def main():
    parser = argparse.ArgumentParser(
        description="Convert .xlsx/.csv/.txt/.docx/.pdf in an input folder to Markdown in an output folder."
    )
    parser.add_argument("input_dir", nargs="?", default="input", help="Input folder (default: input)")
    parser.add_argument("output_dir", nargs="?", default="output", help="Output folder (default: output)")
    args = parser.parse_args()

    input_dir = Path(args.input_dir)
    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    if not input_dir.exists():
        print(f"ERROR: input folder not found: {input_dir}", file=sys.stderr)
        sys.exit(2)

    total = 0
    for in_path in iter_inputs(input_dir):
        try:
            suf = in_path.suffix.lower()
            if suf == ".xlsx":
                convert_xlsx(in_path, output_dir)
            elif suf == ".csv":
                convert_csv(in_path, output_dir)
            elif suf == ".txt":
                convert_txt(in_path, output_dir)
            elif suf == ".docx":
                convert_docx(in_path, output_dir)
            elif suf == ".pdf":
                convert_pdf(in_path, output_dir)

            total += 1
            print(f"OK   {in_path}")
        except Exception as e:
            print(f"FAIL {in_path} -> {e}", file=sys.stderr)

    print(f"\nDone. Converted {total} file(s). Output: {output_dir.resolve()}")


if __name__ == "__main__":
    main()
