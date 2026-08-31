# -*- coding: utf-8 -*-
"""Does `b="0"` on a run turn OFF a bold that the level turned on?

`SlideRun.bold` is a plain `bool`, so "explicitly not bold" and "said nothing"
arrive at the renderer as the same value and it resolves them with
`run.bold || default_bold`. The 2026-08-31 audit found 3675 candidate runs
across 54 decks but no case where the drawing was actually wrong, and it left a
note: an example must be identified PER SLIDE, because the same word carries
different formatting on different slides and matching by string alone once
produced three imaginary ones.

So this asks the truth PDF, one slide at a time, for shapes that carry BOTH a
run with explicit `b="0"` AND a run with no `b` at all. Those are the only
shapes that answer the question on their own evidence: same shape, same level,
one difference. If PowerPoint draws the first in an upright face and the second
in a bold one, `b="0"` overrides the level and the IR cannot say so.

    python tools/metrics/pptx_bzero_census.py [--decks all]
"""
from __future__ import annotations

import argparse
import glob
import re
import sys
import zipfile
from pathlib import Path

import pymupdf

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]

RUN = re.compile(r"<a:r>(.*?)</a:r>", re.S)
RPR = re.compile(r"<a:rPr\b([^>]*?)/?>", re.S)
TEXT = re.compile(r"<a:t>(.*?)</a:t>", re.S)
SP = re.compile(r"<p:sp>.*?</p:sp>", re.S)


def unescape(s: str) -> str:
    return (s.replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">")
             .replace("&quot;", '"').replace("&apos;", "'"))


def shapes_of(xml: str):
    """-> [[(text, bold_attr)]] per shape, bold_attr in {'1', '0', None}."""
    out = []
    for m in SP.finditer(xml):
        runs = []
        for r in RUN.finditer(m.group(0)):
            body = r.group(1)
            t = TEXT.search(body)
            if not t:
                continue
            pr = RPR.search(body)
            b = None
            if pr:
                got = re.search(r'\bb="(\d)"', pr.group(1))
                if got:
                    b = got.group(1)
            runs.append((unescape(t.group(1)), b))
        if runs:
            out.append(runs)
    return out


def page_faces(page):
    """-> [(text, font_name)] for every span PowerPoint drew."""
    out = []
    for block in page.get_text("dict")["blocks"]:
        for line in block.get("lines", []):
            for span in line["spans"]:
                out.append((span["text"], span["font"]))
    return out


def drawn_bold(text: str, spans) -> bool | None:
    """Whether the face PowerPoint drew `text` with names itself bold."""
    t = text.strip()
    if len(t) < 3:
        return None
    hits = {("bold" in f.lower()) for s, f in spans if t[:24] in s or s.strip() and s.strip() in t}
    if len(hits) != 1:
        return None
    return hits.pop()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--decks", nargs="*", default=[])
    ap.add_argument("--all", action="store_true",
                    help="the wider corpus too, not just dev")
    args = ap.parse_args()

    bench = REPO / "pipeline_data" / "pptx_benchmark"
    roots = [(bench / "dev" / "pptx", bench / "dev" / "pdf")]
    if args.all:
        roots.append((bench / "pptx", bench / "ssim_pptx" / "ppt_pdf"))
    pairs = []
    for pdir, fdir in roots:
        for pp in sorted(pdir.glob("*.pptx")):
            stem = pp.name.split("__")[0]
            if args.decks and stem not in args.decks:
                continue
            fp = next(iter(sorted(fdir.glob(stem + "*.pdf"))), None)
            if fp:
                pairs.append((stem, pp, fp))
    decks = [p[0] for p in pairs]
    total_pairs = 0
    overrides = []
    ignored = []
    for deck, pptx, pdf in pairs:
        z = zipfile.ZipFile(pptx)
        doc = pymupdf.open(pdf)
        for name in z.namelist():
            m = re.fullmatch(r"ppt/slides/slide(\d+)\.xml", name)
            if not m:
                continue
            sno = int(m.group(1))
            if sno > len(doc):
                continue
            spans = page_faces(doc[sno - 1])
            # A bold face has to be on the page at all, or "not bold" says
            # nothing -- the deck may simply have no bold to serve.
            if not any("bold" in f.lower() for _, f in spans):
                continue
            for runs in shapes_of(z.read(name).decode("utf-8", "replace")):
                zeros = [t for t, b in runs if b == "0"]
                silent = [t for t, b in runs if b is None]
                if not zeros or not silent:
                    continue
                total_pairs += 1
                for zt in zeros:
                    zb = drawn_bold(zt, spans)
                    for st in silent:
                        sb = drawn_bold(st, spans)
                        if zb is None or sb is None:
                            continue
                        if sb and not zb:
                            overrides.append((deck, sno, zt[:34], st[:34]))
                        elif sb and zb:
                            ignored.append((deck, sno, zt[:34]))
        doc.close()

    print(f"{total_pairs} shapes carry both an explicit b=\"0\" run and a run "
          f"that says nothing, on a page that has a bold face")
    print(f"\n{len(overrides)} of them are drawn as an OVERRIDE "
          f"(the silent run bold, the b=\"0\" run not):")
    seen = set()
    for deck, sno, zt, st in overrides:
        if (deck, sno, zt) in seen:
            continue
        seen.add((deck, sno, zt))
        print(f"   {deck} s{sno:<3} b=0 {zt!r}   beside bold {st!r}")
    print(f"\n{len(ignored)} are drawn BOLD anyway (the attribute ignored):")
    for deck, sno, zt in ignored[:20]:
        print(f"   {deck} s{sno:<3} {zt!r}")
    print(f"\ndecks with an override: "
          f"{sorted({d for d, _, _, _ in overrides})}")


if __name__ == "__main__":
    main()
