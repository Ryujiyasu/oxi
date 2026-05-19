"""Build LibreOffice per-page paragraph records from a libra-rendered PDF,
in the same shape as measure_pagination_oxi.py's output, so that
pagination_diff_libra.py can run the exact diff_doc() logic from
pagination_diff.py.

Per-page record format (matches Oxi):
{
  "doc_id":  "<basename>",
  "n_pages": N,
  "pages":   { "1": [{"text": "..."}, ...], "2": [...], ... }
}

Each "record" is a visual line on the PDF page (clustered the same way
as measure_lla_word.lines_from_pdf, but we keep the line list directly
instead of joining cells). This is sufficient for the pagination diff
since it only matches by text prefix.

Pre-req: pipeline_data/libra_pdf/<doc_id>.pdf must exist.

Usage:
    python tools/metrics/measure_pagination_libra.py            # all libra PDFs
    python tools/metrics/measure_pagination_libra.py 04b88e     # prefix
    python tools/metrics/measure_pagination_libra.py --limit 5
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))
from measure_lla_word import lines_from_pdf  # noqa: E402

REPO_ROOT = SCRIPT_DIR.parent.parent
LIBRA_PDF_DIR = REPO_ROOT / "pipeline_data" / "libra_pdf"
OUT_DIR = REPO_ROOT / "pipeline_data" / "pagination_libra"


def build_for(doc_id: str) -> Path | None:
    pdf = LIBRA_PDF_DIR / f"{doc_id}.pdf"
    if not pdf.is_file():
        return None
    pages = lines_from_pdf(str(pdf))
    pages_with_records = {
        page: [{"text": line} for line in lines]
        for page, lines in pages.items()
    }
    result = {
        "doc_id": doc_id,
        "n_pages": max((int(k) for k in pages_with_records), default=0),
        "pages": pages_with_records,
    }
    out = OUT_DIR / f"{doc_id}.json"
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    return out


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("prefix", nargs="?", default=None)
    ap.add_argument("--limit", type=int, default=0)
    args = ap.parse_args()

    pdfs = sorted(LIBRA_PDF_DIR.glob("*.pdf"))
    doc_ids = [p.stem for p in pdfs]
    if args.prefix:
        doc_ids = [d for d in doc_ids if d.startswith(args.prefix)]
    if args.limit > 0:
        doc_ids = doc_ids[: args.limit]
    if not doc_ids:
        sys.exit("no libra PDFs matched")

    print(f"# building pagination data for {len(doc_ids)} libra PDF(s)")
    n_ok = 0
    for i, doc_id in enumerate(doc_ids, start=1):
        out = build_for(doc_id)
        if out:
            n_ok += 1
            print(f"[{i:3}/{len(doc_ids)}] {doc_id:55.55s} -> {out.name}")
        else:
            print(f"[{i:3}/{len(doc_ids)}] {doc_id:55.55s} MISSING PDF")
    print(f"\n# done. {n_ok}/{len(doc_ids)} ok -> {OUT_DIR}")


if __name__ == "__main__":
    main()
