# -*- coding: utf-8 -*-
"""How many of Word's lines does Oxi break in the same place?

A line-for-line count, the diagnostic the cell-budget thread has been scored on
(kojin 532/706 and so on). Per page the two line-text sequences are aligned with
difflib, and a line counts only when its text matches Word's exactly -- so a
wrap decision one character out costs the line and usually its neighbour.

    python _cw_lines.py kojin 04b88e 34140b            # current environment
    OXI_CELLLAW=1 python _cw_lines.py kojin            # under the derived law

Word truth is the cached export beside each corpus document (`*_rt.pdf`), or
whatever `PDFS` names for the ones held elsewhere.
"""
import difflib
import glob
import unicodedata
import json
import os
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(os.path.dirname(HERE))
sys.path.insert(0, HERE)
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
import _cb_cell_linecount as CL  # noqa: E402
from _cb_cell_linecount import word_lines  # noqa: E402


def oxi_lines(path):
    """As _cb_cell_linecount.oxi_lines, but table-cell text carries a null
    para_idx, which its `sorted(set(...))` cannot order."""
    d = json.load(open(path, encoding="utf-8"))
    pages = []
    for pg in d["pages"]:
        buckets = {}
        for e in pg["elements"]:
            if e.get("type") != "text":
                continue
            buckets.setdefault(round(e["y"] * 2) / 2, []).append(e)
        lines = []
        for y in sorted(buckets):
            els = sorted(buckets[y], key=lambda e: e["x"])
            lines.append((y, min(e["x"] for e in els),
                          max(e["x"] + e["w"] for e in els),
                          "".join(CL.decode(e["text"] or "") for e in els),
                          sorted({e["para_idx"] for e in els}, key=lambda v: (v is None, v))))
        pages.append(lines)
    return pages

DOCS = os.path.join(REPO, "tools", "golden-test", "documents", "docx")
RENDERER = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                        "oxi-gdi-renderer.exe")
PDFS = {"kojin": os.path.join(REPO, "pipeline_data", "_cb_budget", "kojin_000505.pdf")}


def find(prefix):
    docx = sorted(glob.glob(os.path.join(DOCS, prefix + "*.docx")))
    if not docx:
        return None, None
    pdf = PDFS.get(prefix) or docx[0][:-5] + "_rt.pdf"
    return docx[0], (pdf if os.path.exists(pdf) else None)


def dump(docx):
    with tempfile.TemporaryDirectory() as td:
        out = os.path.join(td, "d.json")
        subprocess.run([RENDERER, docx, os.path.join(td, "p"), "96",
                        "--dump-layout=" + out], check=True, capture_output=True)
        return oxi_lines(out)


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    total_m = total_w = 0
    for prefix in args:
        docx, pdf = find(prefix)
        if not pdf:
            print(f"{prefix:<10} no cached Word PDF -- skipped")
            continue
        ox, wd = dump(docx), word_lines(pdf)

        def norm(t):
            """Whitespace is not a wrap decision.

            Word's PDF export puts a trailing space on most lines, renders U+3000 as
            an ordinary space, and MuPDF invents one wherever a gap is wide enough.
            Comparing raw strings scored kojin 63/605 when the wrap decisions largely
            agree; what the metric is asking is only which characters landed on which
            line."""
            return "".join(c for c in unicodedata.normalize("NFKC", t) if not c.isspace())

        m = w = 0
        per_page = []
        for i in range(max(len(ox), len(wd))):
            o = [norm(t[3]) for t in ox[i]] if i < len(ox) else []
            r = [norm(t[3]) for t in wd[i]] if i < len(wd) else []
            o = [t for t in o if t]
            r = [t for t in r if t]
            k = sum(b.size for b in difflib.SequenceMatcher(None, o, r, autojunk=False)
                    .get_matching_blocks())
            m += k
            w += len(r)
            per_page.append((i + 1, k, len(r)))
        total_m += m
        total_w += w
        worst = sorted(per_page, key=lambda p: p[1] - p[2])[:5]
        print(f"{prefix:<10} {m:>5}/{w:<5} lines match  "
              f"(worst pages: {', '.join(f'p{p}{k - n:+d}' for p, k, n in worst if k != n)})")
    if len(args) > 1:
        print(f"{'TOTAL':<10} {total_m:>5}/{total_w:<5}")


if __name__ == "__main__":
    main()
