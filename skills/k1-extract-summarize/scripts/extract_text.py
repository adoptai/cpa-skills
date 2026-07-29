#!/usr/bin/env python3
"""
Extract the text layer and word geometry from PDFs, fully locally.

Bundled with this skill so a single-skill install works standalone. Identical to
the extractor in the bank-statement-to-excel skill.

Writes, per input PDF:
    <stem>.text.txt    page-delimited text
    <stem>.words.jsonl one JSON object per word with coordinates
    <stem>.meta.json   page count, per-page char counts, scanned flag

Usage:
    python3 extract_text.py --input statements/ --out work/
    python3 extract_text.py --input jan.pdf --out work/ --ocr

No network access. Originals are opened read-only and never modified.
"""

from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

try:
    import pdfplumber
except ImportError:
    sys.exit("pdfplumber is required.  pip install pdfplumber")

# A page with fewer than this many extractable characters is treated as a scan.
SCAN_CHAR_THRESHOLD = 60


def ocr_pdf(src: Path) -> Path:
    """Run local OCR, returning a path to a new searchable PDF. Never touches src."""
    ocrmypdf = shutil.which("ocrmypdf")
    if not ocrmypdf:
        sys.exit(
            "This PDF has no text layer and 'ocrmypdf' is not installed.\n"
            "Install it locally, then re-run with --ocr:\n"
            "  macOS:  brew install ocrmypdf\n"
            "  Debian: sudo apt install ocrmypdf tesseract-ocr\n"
            "Do not upload client tax documents to a cloud OCR service."
        )
    # Resolve to an absolute input path and confirm it is a real PDF before handing
    # it to another process. No shell is involved (list argv, shell=False), and the
    # binary is the absolute path resolved above rather than a PATH lookup at spawn.
    src = src.resolve(strict=True)
    if not src.is_file() or src.suffix.lower() != ".pdf":
        sys.exit(f"Not a PDF file: {src}")

    out = Path(tempfile.mkdtemp(prefix="ocr_")) / f"{src.stem}.ocr.pdf"
    print(f"  OCR (local): {src.name} -> {out.name}", flush=True)
    subprocess.run(
        [
            ocrmypdf,
            "--force-ocr",
            "--deskew",
            "--optimize", "0",
            "--quiet",
            str(src),
            str(out),
        ],
        check=True,
    )
    return out


def extract(pdf_path: Path, out_dir: Path, use_ocr: bool) -> dict:
    stem = pdf_path.stem
    source = pdf_path

    if use_ocr:
        source = ocr_pdf(pdf_path)

    text_lines: list[str] = []
    words_out: list[str] = []
    page_meta: list[dict] = []

    with pdfplumber.open(str(source)) as pdf:
        for pageno, page in enumerate(pdf.pages, start=1):
            txt = page.extract_text() or ""
            text_lines.append(f"\n===== PAGE {pageno} =====\n{txt}")

            for w in page.extract_words(
                use_text_flow=False,
                keep_blank_chars=False,
                extra_attrs=["size"],
            ):
                words_out.append(
                    json.dumps(
                        {
                            "page": pageno,
                            "text": w["text"],
                            "x0": round(float(w["x0"]), 2),
                            "x1": round(float(w["x1"]), 2),
                            "top": round(float(w["top"]), 2),
                            "bottom": round(float(w["bottom"]), 2),
                            "size": round(float(w.get("size", 0)), 1),
                        },
                        ensure_ascii=False,
                    )
                )

            page_meta.append(
                {
                    "page": pageno,
                    "chars": len(txt),
                    "scanned": len(txt.strip()) < SCAN_CHAR_THRESHOLD,
                    "width": round(float(page.width), 1),
                    "height": round(float(page.height), 1),
                }
            )

    out_dir.mkdir(parents=True, exist_ok=True)
    (out_dir / f"{stem}.text.txt").write_text("\n".join(text_lines), encoding="utf-8")
    (out_dir / f"{stem}.words.jsonl").write_text(
        "\n".join(words_out) + "\n", encoding="utf-8"
    )

    scanned_pages = [p["page"] for p in page_meta if p["scanned"]]
    meta = {
        "source_file": pdf_path.name,
        "ocr_applied": use_ocr,
        "page_count": len(page_meta),
        "scanned_pages": scanned_pages,
        "scanned": bool(scanned_pages),
        "pages": page_meta,
    }
    (out_dir / f"{stem}.meta.json").write_text(
        json.dumps(meta, indent=2), encoding="utf-8"
    )
    return meta


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--input", required=True, help="PDF file or folder of PDFs")
    ap.add_argument("--out", default="work", help="output directory (default: work)")
    ap.add_argument(
        "--ocr",
        action="store_true",
        help="run local OCR first (required for scanned statements)",
    )
    args = ap.parse_args()

    src = Path(args.input).expanduser()
    out_dir = Path(args.out).expanduser()

    if src.is_dir():
        pdfs = sorted(src.glob("*.pdf")) + sorted(src.glob("*.PDF"))
    elif src.is_file():
        pdfs = [src]
    else:
        sys.exit(f"Not found: {src}")

    if not pdfs:
        sys.exit(f"No PDFs found in {src}")

    print(f"Extracting {len(pdfs)} PDF(s) -> {out_dir}/")
    needs_ocr: list[str] = []

    for p in pdfs:
        print(f"- {p.name}")
        meta = extract(p, out_dir, args.ocr)
        flag = ""
        if meta["scanned"] and not args.ocr:
            needs_ocr.append(p.name)
            flag = f"  [SCANNED pages {meta['scanned_pages']} - re-run with --ocr]"
        print(f"    {meta['page_count']} page(s){flag}")

    if needs_ocr:
        print(
            "\nThese files have pages with no usable text layer and MUST be re-run "
            "with --ocr before you read any figures from them:"
        )
        for n in needs_ocr:
            print(f"  - {n}")
        print(
            "Reading amounts off a scanned page without OCR produces silent blanks, "
            "not errors."
        )
        return 2

    print("\nDone. Text layer looks usable on every page.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
