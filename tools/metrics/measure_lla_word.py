"""Measure Word's per-page, per-line text via PDF export + pymupdf parsing.

This is the third attempt at the Word side of the LLA metric. The previous
two attempts hit fundamental Word-COM model limits:

1. `GoTo(wdGoToLine, N)` + `wdStatisticLines` skipped table-cell lines
   entirely (a1d6 p.1: 33 lines vs Oxi's 48).
2. `Selection.MoveDown(wdLine)` with `Visible=True` reordered cells in
   weird ways and inserted phantom empty lines (a1d6 p.1: 29 lines, wrong
   order).
3. `Range.Information(wdFirstCharacterLineNumber)` per-char was correct
   but 5+ minutes per doc, plus polluted output with `\x07` cell markers
   (a1d6 p.1: 94 lines).

PDF export bypasses Word's COM line iterator entirely. We render to PDF
once (Word's "ExportAsFixedFormat" matches the on-screen layout), then
parse with pymupdf which gives us per-span (x, y, text) tuples. Grouping
by Y is then symmetric with the Oxi side — both sides bucket the same way,
so any divergence is genuine layout disagreement.

Output schema (same as the v1/v2 measurers):
{
  "doc_id":  "<basename>",
  "n_pages": N,
  "pages": { "1": ["line1", "line2", ...], ... }
}
"""
from __future__ import annotations

import argparse
import json
import os
import sys
import tempfile

Y_BUCKET_TOL_PT = 3.0   # same tolerance as Oxi-side measure_lla_oxi.py


def _strip_line(text: str) -> str:
    while text and text[-1] in "\r\n\v\x07":
        text = text[:-1]
    return text


def _normalise(text: str) -> str:
    # Same normalisation as measure_lla_oxi.py: drop ALL whitespace
    # including U+3000 ideographic spaces. Word's PDF export converts
    # docx U+3000 padding to ASCII spaces, so we can't trust whitespace
    # encoding either way.
    out = []
    for ch in text:
        if ch in (" ", "\t", "　"):
            continue
        out.append(ch)
    return "".join(out)


def export_docx_to_pdf(docx_abs: str, pdf_abs: str) -> None:
    import win32com.client

    word = win32com.client.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        doc = word.Documents.Open(docx_abs, ReadOnly=True, AddToRecentFiles=False)
        try:
            doc.Repaginate()
            # wdExportFormatPDF = 17, wdExportOptimizeForPrint = 0
            doc.ExportAsFixedFormat(
                OutputFileName=pdf_abs,
                ExportFormat=17,
                OpenAfterExport=False,
                OptimizeFor=0,
                Range=0,
                Item=0,
                IncludeDocProps=False,
                KeepIRM=False,
                CreateBookmarks=0,
                DocStructureTags=False,
                BitmapMissingFonts=True,
                UseISO19005_1=False,
            )
        finally:
            doc.Close(SaveChanges=False)
    finally:
        word.Quit()


def lines_from_pdf(pdf_abs: str) -> dict[str, list[str]]:
    import fitz  # pymupdf

    pages_out: dict[str, list[str]] = {}
    with fitz.open(pdf_abs) as pdf:
        for page_idx, page in enumerate(pdf, start=1):
            # `get_text("dict")` returns blocks → lines → spans, where each
            # line already has its own bbox. Use that directly — pymupdf's
            # line clustering matches the visual line structure of the PDF.
            page_dict = page.get_text("dict")
            entries: list[tuple[float, float, str]] = []  # (y, x, text)
            for block in page_dict.get("blocks", []):
                if block.get("type") != 0:  # 0 = text block; 1 = image
                    continue
                for line in block.get("lines", []):
                    bbox = line.get("bbox") or [0, 0, 0, 0]
                    y = float(bbox[1])
                    text = "".join(s.get("text", "") for s in line.get("spans", []))
                    if not line.get("spans"):
                        continue
                    x = float(line["spans"][0].get("bbox", [0, 0, 0, 0])[0])
                    entries.append((y, x, text))

            # Re-cluster across pymupdf's lines: sometimes a single visual
            # line is split into multiple pymupdf "lines" with very close Y
            # values. Merge if |dy| <= Y_BUCKET_TOL_PT.
            entries.sort()
            clusters: list[tuple[float, list[tuple[float, str]]]] = []
            for y, x, t in entries:
                if clusters and abs(y - clusters[-1][0]) <= Y_BUCKET_TOL_PT:
                    clusters[-1][1].append((x, t))
                else:
                    clusters.append((y, [(x, t)]))

            lines: list[str] = []
            for _y, chars in clusters:
                chars.sort(key=lambda c: c[0])
                lines.append(_normalise(_strip_line("".join(c[1] for c in chars))))
            # Drop empty lines — Word's PDF export emits them for empty
            # paragraphs in cells which the Oxi IR doesn't always retain.
            # Both sides agree to ignore them.
            lines = [ln for ln in lines if ln != ""]
            pages_out[str(page_idx)] = lines
    return pages_out


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("docx")
    ap.add_argument("-o", "--output", required=True)
    ap.add_argument("--pdf-cache", default=None,
                    help="reuse / cache the exported PDF at this path")
    args = ap.parse_args()

    docx_abs = os.path.abspath(args.docx)
    if not os.path.isfile(docx_abs):
        print(f"docx not found: {docx_abs}", file=sys.stderr)
        sys.exit(2)

    if args.pdf_cache:
        pdf_abs = os.path.abspath(args.pdf_cache)
        os.makedirs(os.path.dirname(pdf_abs), exist_ok=True)
        if not os.path.isfile(pdf_abs):
            export_docx_to_pdf(docx_abs, pdf_abs)
    else:
        tmp = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
        tmp.close()
        pdf_abs = tmp.name
        try:
            export_docx_to_pdf(docx_abs, pdf_abs)
        except Exception:
            os.unlink(pdf_abs)
            raise

    try:
        pages = lines_from_pdf(pdf_abs)
    finally:
        if not args.pdf_cache:
            try:
                os.unlink(pdf_abs)
            except OSError:
                pass

    result = {
        "doc_id": os.path.splitext(os.path.basename(docx_abs))[0],
        "n_pages": max((int(k) for k in pages), default=0),
        "pages": pages,
    }
    with open(args.output, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    total = sum(len(v) for v in result["pages"].values())
    print(f"# wrote {args.output}  ({result['n_pages']} pages, {total} lines)")


if __name__ == "__main__":
    main()
