# -*- coding: utf-8 -*-
"""Which embedded SLOT does PowerPoint actually draw for a bold run?

`slot_face_name` selects an embedded part by its `p:embeddedFont` slot; its
companion rule then rejects any part whose internal family is not the typeface it
is filed under (55 of 262 corpus parts, 21%). d04 falsifies that rejection: its
"Inria Sans Light" BOLD slot is byte-for-byte "Inria Sans"'s REGULAR part, an
internal-family mismatch, and PowerPoint draws it for every `b="1"` run on
slide 2 -- even though a genuine InriaSans-Bold sits in the same deck.

One deck is not a re-derivation, so this asks all fifty, and it asks without
rendering anything: the parts are EOT-compressed and cannot be read directly, but
they can be compared to EACH OTHER by hash, and PowerPoint's PDF names the
PostScript face of every span it drew. Together those pin the slot:

  * `bold slot sha == some other family's regular slot sha` says what the bold
    slot HOLDS, without decompressing it.
  * a `b="1"` run whose PDF span carries a DIFFERENT PostScript name from its
    plain-text neighbours says PowerPoint changed face for the bold.

A deck where the two agree is a deck that took the slot. A deck where the bold
run keeps the plain face, or jumps to a name neither slot holds, is one the
rejection rule may be right about -- and that is the set worth reading.

Usage:
    python tools/metrics/pptx_slot_census.py [--deck 04]
"""
from __future__ import annotations

import argparse
import hashlib
import io
import json
import re
import sys
import zipfile
from collections import defaultdict
from pathlib import Path

import pymupdf
from fontTools.ttLib import TTFont

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
ROOT = REPO / "pipeline_data" / "pptx_benchmark"
SS = ROOT / "ssim_pptx"
SLOTS = ("regular", "bold", "italic", "boldItalic")


def embedded_parts(z: zipfile.ZipFile) -> dict:
    """family -> {slot: (sha, bytes_len)} from p:embeddedFont."""
    try:
        pres = z.read("ppt/presentation.xml").decode("utf-8", "replace")
        rels = z.read("ppt/_rels/presentation.xml.rels").decode("utf-8", "replace")
    except KeyError:
        return {}
    rmap = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
    out = {}
    for m in re.finditer(r"<p:embeddedFont>.*?</p:embeddedFont>", pres, re.S):
        blk = m.group(0)
        fam = re.search(r'typeface="([^"]+)"', blk)
        if not fam:
            continue
        slots = {}
        for slot, rid in re.findall(r'<p:(regular|bold|italic|boldItalic) r:id="([^"]+)"', blk):
            t = rmap.get(rid, "").lstrip("/")
            t = t if t.startswith("ppt/") else "ppt/" + t
            try:
                d = z.read(t)
            except KeyError:
                continue
            slots[slot] = (hashlib.sha256(d).hexdigest()[:12], len(d))
        out[fam.group(1)] = slots
    return out


def ps_names(pdf, page, cache: dict) -> dict:
    """PDF font resource name -> PostScript name of the embedded subset."""
    out = {}
    for xref, _ext, _typ, name, *_ in page.get_fonts(full=True):
        if xref in cache:
            ps = cache[xref]
            out[name] = ps
            out[name.split("+", 1)[-1]] = ps
            continue
        try:
            buf = pdf.extract_font(xref)
            data = buf[3] if isinstance(buf, tuple) and len(buf) > 3 else None
            ps = TTFont(io.BytesIO(data), lazy=True,
                        checkChecksums=0)["name"].getDebugName(6)
        except Exception:
            ps = None
        # A span reports the font WITHOUT the subset prefix that `get_fonts`
        # keeps ("36" against "BCDFEE+36"), so file it under both spellings --
        # keying on one alone silently falls back to the resource name and the
        # census then compares numbers instead of typefaces.
        cache[xref] = ps
        out[name] = ps
        out[name.split("+", 1)[-1]] = ps
    return out


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--deck", default="")
    args = ap.parse_args()
    man = json.loads((ROOT / "manifest.json").read_text(encoding="utf-8"))
    tally = defaultdict(int)
    for item in man:
        doc = f"{item['idx']:02d}"
        if args.deck and doc != args.deck.zfill(2):
            continue
        src = ROOT / "pptx" / item["local"]
        pdfp = SS / "ppt_pdf" / f"{doc}.pdf"
        if not src.exists() or not pdfp.exists():
            continue
        try:
            z = zipfile.ZipFile(src)
        except Exception:
            continue
        parts = embedded_parts(z)
        if not parts:
            continue
        # what each bold slot HOLDS, named by the family whose regular it equals
        holds = {}
        for fam, slots in parts.items():
            b = slots.get("bold", (None,))[0]
            r = slots.get("regular", (None,))[0]
            if b is None:
                continue
            if b == r:
                holds[fam] = f"same file as its own regular"
            else:
                other = [f2 for f2, s2 in parts.items()
                         if f2 != fam and s2.get("regular", (None,))[0] == b]
                holds[fam] = f"= {other[0]}'s regular" if other else "a distinct part"
        pdf = pymupdf.open(pdfp)
        fcache: dict = {}
        rows = []
        for n in z.namelist():
            m = re.fullmatch(r"ppt/slides/slide(\d+)\.xml", n)
            if not m:
                continue
            sl = int(m.group(1))
            if sl > len(pdf):
                continue
            x = z.read(n).decode("utf-8", "replace")
            if 'b="1"' not in x:
                continue
            page = pdf[sl - 1]
            names = ps_names(pdf, page, fcache)
            spans = []
            for b in page.get_text("rawdict")["blocks"]:
                if b["type"] != 0:
                    continue
                for l in b["lines"]:
                    for s in l["spans"]:
                        t = "".join(c["c"] for c in s["chars"]).strip()
                        if len(t) >= 6:
                            spans.append((t, names.get(s["font"], s["font"])))
            bold, plain = set(), set()
            for rm in re.finditer(r"<a:r>.*?</a:r>", x, re.S):
                r = rm.group(0)
                t = re.search(r"<a:t>([^<]*)</a:t>", r)
                if not t or len(t.group(1).strip()) < 6:
                    continue
                head = t.group(1).strip()[:16]
                hit = next((ps for st, ps in spans if st.startswith(head[:12])), None)
                if not hit:
                    continue
                (bold if re.search(r'\bb="1"', r) else plain).add(hit)
            if bold and plain:
                rows.append((sl, sorted(bold), sorted(plain)))
        pdf.close()
        if not rows:
            continue
        changed = sum(1 for _, b, p in rows if set(b) - set(p))
        tally["decks"] += 1
        tally["slides"] += len(rows)
        tally["bold_changed_face"] += changed
        print(f"\n{doc}: {len(rows)} slides mixing bold and plain, "
              f"{changed} where the bold run changes face")
        for fam, what in holds.items():
            print(f"    bold slot of {fam!r}: {what}")
        for sl, b, p in rows[:3]:
            print(f"    s{sl:<3} bold={b}  plain={p}")
    print(f"\n{tally['decks']} decks, {tally['slides']} mixed slides, "
          f"{tally['bold_changed_face']} where bold uses a face the plain runs do not")


if __name__ == "__main__":
    main()
