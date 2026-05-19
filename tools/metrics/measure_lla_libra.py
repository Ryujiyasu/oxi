"""Build LibreOffice-side per-page, per-line text from a libra-rendered PDF.

Mirror of measure_lla_oxi.py / measure_lla_word.py. Reuses the exact same
`lines_from_pdf` clustering logic from measure_lla_word.py so the Libra
side is symmetric with the Word side (both go through pymupdf with the
same Y-bucketing tolerance and whitespace normalisation).

Output schema (identical to Word/Oxi LLA JSON):
{
  "doc_id":  "<basename>",
  "n_pages": N,
  "pages":   { "1": ["line1", "line2", ...], ... }
}

Pre-req: pipeline_data/libra_pdf/<doc_id>.pdf must already exist (from
render_libra.py). If --pdf is given, that path is used instead.

Usage:
    python tools/metrics/measure_lla_libra.py <doc_id> -o out.json
    python tools/metrics/measure_lla_libra.py --pdf path/to/libra.pdf -o out.json
"""
from __future__ import annotations

import argparse
import json
import os
import sys
from pathlib import Path

# Reuse the canonical PDF -> lines function from measure_lla_word.py.
SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))
from measure_lla_word import lines_from_pdf  # noqa: E402

REPO_ROOT = SCRIPT_DIR.parent.parent
LIBRA_PDF_DIR = REPO_ROOT / "pipeline_data" / "libra_pdf"


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("doc_id", nargs="?", default=None,
                    help="doc_id (matches libra_pdf/<doc_id>.pdf)")
    ap.add_argument("--pdf", default=None, help="explicit PDF path (overrides doc_id lookup)")
    ap.add_argument("-o", "--output", required=True)
    args = ap.parse_args()

    if args.pdf:
        pdf_path = Path(args.pdf)
        doc_id = pdf_path.stem
    elif args.doc_id:
        pdf_path = LIBRA_PDF_DIR / f"{args.doc_id}.pdf"
        doc_id = args.doc_id
    else:
        sys.exit("must give either doc_id or --pdf")

    if not pdf_path.is_file():
        sys.exit(f"PDF not found: {pdf_path}")

    pages = lines_from_pdf(str(pdf_path))
    result = {
        "doc_id": doc_id,
        "n_pages": max((int(k) for k in pages), default=0),
        "pages": pages,
    }
    out = Path(args.output)
    out.parent.mkdir(parents=True, exist_ok=True)
    with out.open("w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    total = sum(len(v) for v in result["pages"].values())
    print(f"# wrote {out}  ({result['n_pages']} pages, {total} lines)")


if __name__ == "__main__":
    main()
