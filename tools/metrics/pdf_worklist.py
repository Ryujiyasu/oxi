# -*- coding: utf-8 -*-
"""Which documents to fix next, and where each one first goes wrong.

  python tools/metrics/pdf_worklist.py <word_pdf_dir> <oxi_pdf_dir> [--out=FILE.json]
         [--tol=2.0] [--jobs=4] [--top=30]

For every <stem>.pdf in the Word directory with a same-named Oxi PDF:
  * the scores of score_engine_column (150 dpi SSIM per common page,
    penalized = sum / max(pages)) and the page-count delta;
  * the FIRST DIVERGENCE: both PDFs' text lines are read in reading order and
    aligned by their text (difflib over the whole document); the first aligned
    line that sits on another page or more than --tol pt away vertically is
    reported with its page, both y values and the text.  A line Word has and Oxi
    lost (or the reverse) before that point is reported instead.
The list is ordered as the fix loop works: page-count misses first, then the
lowest penalized score.  The first-divergence text is where to open the
document; everything after it is often only the consequence.
"""
import difflib
import json
import os
import sys
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import fitz
import numpy as np
from PIL import Image
from skimage.metrics import structural_similarity

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
DPI = 150


def rgb(pdf, i):
    pix = pdf.load_page(i).get_pixmap(dpi=DPI, colorspace=fitz.csRGB, alpha=False)
    return np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, 3)


def ssim(ref, cand):
    if cand.shape[:2] != ref.shape[:2]:
        h, w = ref.shape[:2]
        cand = np.array(Image.fromarray(cand).resize((w, h), Image.LANCZOS))
    return float(structural_similarity(ref, cand, channel_axis=2))


def lines(pdf):
    """[(page, y_top, text)] for every text line, reading order per page."""
    out = []
    for pi, page in enumerate(pdf):
        for b in page.get_text("dict")["blocks"]:
            for ln in b.get("lines", []):
                t = "".join(s["text"] for s in ln["spans"]).strip()
                if t:
                    out.append((pi + 1, round(ln["bbox"][1], 2), " ".join(t.split())))
    return out


def first_divergence(wl, ol, tol):
    sm = difflib.SequenceMatcher(None, [x[2] for x in wl], [x[2] for x in ol], autojunk=False)
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            for k in range(i2 - i1):
                w, o = wl[i1 + k], ol[j1 + k]
                if w[0] != o[0] or abs(w[1] - o[1]) > tol:
                    return {"kind": "moved", "word_page": w[0], "word_y": w[1], "oxi_page": o[0],
                            "oxi_y": o[1], "dy": round(o[1] - w[1], 2), "text": w[2][:60]}
        else:
            w = wl[i1] if i1 < len(wl) else None
            o = ol[j1] if j1 < len(ol) else None
            # a re-broken line reads differently on each side: same page, close y
            if w and o and w[0] == o[0] and abs(w[1] - o[1]) <= tol:
                return {"kind": "rebroken", "word_page": w[0], "word_y": w[1], "oxi_y": o[1],
                        "text": w[2][:60], "oxi_text": o[2][:60]}
            return {"kind": tag, "word_page": w[0] if w else None, "word_y": w[1] if w else None,
                    "oxi_page": o[0] if o else None, "oxi_y": o[1] if o else None,
                    "text": (w or o)[2][:60] if (w or o) else "", "oxi_text": o[2][:60] if o else ""}
    return None


def one(wp, op, tol):
    W, O = fitz.open(wp), fitz.open(op)
    nw, no = W.page_count, O.page_count
    s = [ssim(rgb(W, i), rgb(O, i)) for i in range(min(nw, no))]
    div = first_divergence(lines(W), lines(O), tol)
    W.close(); O.close()
    d = max(nw, no)
    return {"word_pages": nw, "pages": no, "page_delta": no - nw,
            "penalized_mean": round(sum(s) / d, 6) if d else None,
            "worst_page": (1 + int(np.argmin(s))) if s else None,
            "worst_ssim": round(min(s), 4) if s else None, "first_div": div}


def main(argv):
    opts = {"tol": 2.0, "jobs": 4, "top": 30, "out": None}
    pos = []
    for a in argv:
        if a.startswith("--") and "=" in a:
            k, v = a[2:].split("=", 1)
            opts[k] = v
        else:
            pos.append(a)
    if len(pos) != 2:
        print(__doc__); return 2
    wd, od = Path(pos[0]), Path(pos[1])
    tol, jobs, top = float(opts["tol"]), int(opts["jobs"]), int(opts["top"])
    stems = [p.stem for p in sorted(wd.glob("*.pdf")) if (od / p.name).is_file()]

    def job(stem):
        try:
            return stem, one(wd / f"{stem}.pdf", od / f"{stem}.pdf", tol)
        except Exception as e:  # a broken PDF is a finding, not a crash
            return stem, {"error": str(e)[:200]}

    with ThreadPoolExecutor(jobs) as ex:
        res = dict(ex.map(job, stems))
    if opts["out"]:
        Path(opts["out"]).write_text(json.dumps(res, ensure_ascii=False, indent=1), encoding="utf-8")
    ok = [(k, r) for k, r in res.items() if "error" not in r]
    ok.sort(key=lambda kr: (kr[1]["page_delta"] == 0, kr[1]["penalized_mean"] or 0))
    pen = [r["penalized_mean"] or 0 for _, r in ok]
    print(f"docs {len(res)}  penalized mean {sum(pen) / max(1, len(pen)):.4f}  "
          f"page-count match {sum(1 for _, r in ok if r['page_delta'] == 0)}  errors {len(res) - len(ok)}")
    for k, r in ok[:top]:
        d = r["first_div"]
        where = "-"
        if d:
            if d["kind"] == "moved":
                where = f"p{d['word_page']} y{d['word_y']} -> p{d['oxi_page']} y{d['oxi_y']} ({d['kind']})"
            else:
                where = f"p{d.get('word_page')} y{d.get('word_y')} ({d['kind']})"
        print(f"{r['penalized_mean']:.3f} pd={r['page_delta']:+d} worst p{r['worst_page']}={r['worst_ssim']}  {k}")
        if d:
            print(f"      first: {where}  «{d['text']}»" + (f" / oxi «{d['oxi_text']}»" if d.get("oxi_text") and d["kind"] != "moved" else ""))
    for k, r in res.items():
        if "error" in r:
            print(f"ERROR {k}: {r['error']}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
