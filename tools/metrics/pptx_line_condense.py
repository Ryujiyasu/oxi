# -*- coding: utf-8 -*-
"""Is the pen pullback inside a TJ array a UNIFORM scale on the line?

`pptx_tj_census.py` measures how much narrower than its /Widths sum PowerPoint
draws a run, and recorded that the adjustments "scatter (-11..+29 thousandths)
rather than repeating one value per character" -- read as evidence that the
pullback is not one rule.

That reading confuses the correction with the thing corrected. A TJ adjustment
is the CUMULATIVE error the exporter has to give back at that point, not a
per-glyph rate, so a perfectly uniform scale still produces adjustments that
scatter: they are the running remainder of `s * design_cum` against a position
built from integer /Widths. The question has to be asked of the POSITIONS.

So for every run this fits

    drawn_cum(k) = s * design_cum(k)

by least squares through the origin, and reports the worst point's residual
beside the worst residual of the same run at s = 1. When the ratio of those two
is large, one number -- the run's own horizontal scale -- explains the whole
line and the scatter was the instrument's.

    python tools/metrics/pptx_line_condense.py [deck ...] [--all] [--min-chars N]
"""
from __future__ import annotations

import argparse
import glob
import re
import sys
from pathlib import Path

import pymupdf

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
BACKSLASH = chr(92)


def parse_tj(body):
    """-> [(text, adjustment_after)] for one TJ array body.

    Kept per piece rather than flattened, because the position of each
    adjustment is the whole point here.

    ★PDF strings come in two spellings, literal `(...)` and hex `<...>`. A
    parser that handles only one finds no adjustments on half the corpus and
    invites the conclusion that there are none -- which is how this finding was
    once retracted in error.
    """
    pieces, i, pending = [], 0, None
    while i < len(body):
        ch = body[i]
        if ch == "(":
            depth, i, buf = 1, i + 1, []
            while i < len(body) and depth:
                c = body[i]
                if c == BACKSLASH:
                    buf.append(body[i + 1])
                    i += 2
                    continue
                if c == "(":
                    depth += 1
                elif c == ")":
                    depth -= 1
                    if not depth:
                        i += 1
                        break
                buf.append(c)
                i += 1
            pieces.append(["".join(buf), 0.0])
        elif ch == "<":
            j = body.find(">", i)
            if j < 0:
                break
            h = re.sub(r"[^0-9A-Fa-f]", "", body[i + 1:j])
            if len(h) % 2:
                h += "0"
            try:
                pieces.append([bytes.fromhex(h).decode("latin-1"), 0.0])
            except ValueError:
                pass
            i = j + 1
        elif ch == "-" or ch.isdigit() or ch == ".":
            m = re.match(r"-?\d+\.?\d*", body[i:])
            if pieces:
                pieces[-1][1] += float(m.group())
            i += len(m.group())
        else:
            i += 1
    return pieces


def widths_for(doc, page):
    """resource name (F1..) -> (FirstChar, [widths], BaseFont)"""
    out = {}
    for xref, _, _, base, name, _, _ in page.get_fonts(full=True):
        obj = doc.xref_object(xref)
        fc = re.search(r"/FirstChar\s+(\d+)", obj)
        wm = re.search(r"/Widths\s+(\d+)\s+0\s+R", obj)
        if wm:
            body = doc.xref_object(int(wm.group(1)))
        else:
            inline = re.search(r"/Widths\s*\[(.*?)\]", obj, re.S)
            body = inline.group(1) if inline else None
        if fc and body:
            out[name] = (int(fc.group(1)),
                         [float(x) for x in re.findall(r"-?\d+\.?\d*", body)],
                         base)
    return out


def fit(design_cum, drawn_cum):
    """The one scale that best carries design positions onto drawn ones."""
    den = sum(a * a for a in design_cum)
    if den <= 0:
        return None
    s = sum(a * b for a, b in zip(design_cum, drawn_cum)) / den
    scaled = max(abs(s * a - b) for a, b in zip(design_cum, drawn_cum))
    plain = max(abs(a - b) for a, b in zip(design_cum, drawn_cum))
    return s, scaled, plain


def runs_of(pdf: Path, min_chars: int):
    doc = pymupdf.open(pdf)
    for pno in range(len(doc)):
        page = doc[pno]
        wtab = widths_for(doc, page)
        raw = page.read_contents().decode("latin-1", "replace")
        cur = None
        for m in re.finditer(r"/(\w+)\s+([\d.]+)\s+Tf|\[(.*?)\]\s*TJ", raw, re.S):
            if m.group(1):
                cur = (m.group(1), float(m.group(2)))
                continue
            if cur is None or cur[0] not in wtab:
                continue
            first, widths, base = wtab[cur[0]]
            size = cur[1]
            design_cum, drawn_cum = [0.0], [0.0]
            d = w = 0.0
            text, ok = [], True
            for piece, adj in parse_tj(m.group(3)):
                for c in piece:
                    code = ord(c)
                    if not (first <= code < first + len(widths)):
                        ok = False
                        break
                    d += widths[code - first] / 1000 * size
                    w += widths[code - first] / 1000 * size
                    design_cum.append(d)
                    drawn_cum.append(w)
                if not ok:
                    break
                text.append(piece)
                # The adjustment moves the pen BACK by adj/1000 of the em.
                w -= adj / 1000 * size
                drawn_cum[-1] = w
            if not ok or len("".join(text)) < min_chars or d <= 0:
                continue
            got = fit(design_cum, drawn_cum)
            if got:
                yield pno + 1, base.split("+")[-1], size, "".join(text), got, d


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("decks", nargs="*", default=[])
    ap.add_argument("--all", action="store_true", help="every deck with a truth PDF")
    ap.add_argument("--min-chars", type=int, default=20)
    ap.add_argument("--report", type=int, default=15)
    args = ap.parse_args()

    pdfs = []
    root = REPO / "pipeline_data" / "pptx_benchmark" / "dev" / "pdf"
    if args.all or not args.decks:
        pdfs = [Path(p) for p in sorted(glob.glob(str(root / "*.pdf")))]
    else:
        for d in args.decks:
            pdfs += [Path(p) for p in sorted(glob.glob(str(root / (d + "__*.pdf"))))]

    rows = []
    for pdf in pdfs:
        stem = pdf.name.split("__")[0]
        for pno, base, size, text, (s, scaled, plain), width in runs_of(pdf, args.min_chars):
            rows.append({"deck": stem, "page": pno, "font": base, "size": size,
                         "scale": s, "resid": scaled, "plain": plain,
                         "width": width, "text": text})
    if not rows:
        print("no runs")
        return

    pulled = [r for r in rows if r["plain"] > 0.5]
    print(f"{len(rows)} runs of >= {args.min_chars} characters; "
          f"{len(pulled)} are drawn more than 0.5pt off their /Widths sum")
    if pulled:
        better = [r for r in pulled if r["resid"] < r["plain"] / 3]
        print(f"of those, {len(better)} ({100*len(better)/len(pulled):.0f}%) are "
              f"explained to within a third by ONE scale on the line")
        med = sorted(r["scale"] for r in pulled)[len(pulled) // 2]
        print(f"median scale on a pulled run: {med:.5f}  "
              f"({(med-1)*100:+.3f}%)")
    pulled.sort(key=lambda r: r["scale"])
    print(f"\ntightest {args.report}:")
    print(f"{'scale':>9}{'resid':>8}{'plain':>8}{'width':>8}  deck  font / size / text")
    for r in pulled[: args.report]:
        print(f"{r['scale']:9.5f}{r['resid']:8.3f}{r['plain']:8.3f}{r['width']:8.1f}  "
              f"{r['deck']:4} {r['font'][:18]:19} {r['size']:g}  {r['text'][:34]!r}")

    by_font: dict[str, list] = {}
    for r in pulled:
        by_font.setdefault(r["font"], []).append(r["scale"])
    print("\nby face (faces with 5+ pulled runs):")
    for f, v in sorted(by_font.items(), key=lambda kv: sum(kv[1]) / len(kv[1])):
        if len(v) < 5:
            continue
        v = sorted(v)
        print(f"   {len(v):4}  mean {sum(v)/len(v):.5f}  "
              f"min {v[0]:.5f}  max {v[-1]:.5f}   {f}")


if __name__ == "__main__":
    main()
