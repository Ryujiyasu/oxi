"""Strip the panel one part at a time and ask Excel each time.

The panel takes no shortfall where every synthetic arm takes one, and it keeps
refusing when transplanted alone into the same workbook — so whatever decides
it is written in the shape. Each arm removes one thing and is compared against
our own render, which never applies a shortfall: matching means Excel did not
apply one either.
"""
import re
import subprocess
import zipfile
from pathlib import Path

import numpy as np
from PIL import Image

SRC = Path("tools/golden-test/documents/xlsx/"
           "33ac9f9d7afc_20230411_resources_standard_guidelines_glossary_05.xlsx")
OUT = Path(r"C:\tmp\bisect")
RENDER = Path("tools/oxi-xlsx-renderer/target/release/oxi-xlsx-renderer.exe")
OUT.mkdir(parents=True, exist_ok=True)

source = zipfile.ZipFile(SRC)
drawing = source.read("xl/drawings/drawing1.xml").decode("utf-8")
head = drawing[: drawing.index("<xdr:twoCellAnchor")]
panel = next(m.group(0) for m in
             re.finditer(r"<xdr:twoCellAnchor[^>]*>.*?</xdr:twoCellAnchor>", drawing, re.S)
             if len("".join(re.findall(r"<a:t>([^<]*)</a:t>", m.group(0)))) > 120)

first = re.search(r"<a:p>.*?</a:p>", panel, re.S).group(0)
one_line = re.sub(r"<a:t>[^<]*</a:t>", "<a:t>AAAA</a:t>", first)


def arms():
    yield "control", panel
    # Which EDGE of this box is wrong. A block anchored `t` hangs from the head
    # and one anchored `b` hangs from the foot, so each arm reads one edge on
    # its own — where the centred control mixes the two and can only say that
    # the box wants to be a pixel taller.
    for where in ("t", "b"):
        body = re.sub(r'anchor="ctr"', f'anchor="{where}"', panel)
        yield f"anchor-{where}", body
    # The same, one line, so nothing overflows either end.
    for where in ("t", "b"):
        body = re.sub(r"<a:p>.*</a:p>", one_line, panel, flags=re.S)
        body = re.sub(r'anchor="ctr"', f'anchor="{where}"', body)
        yield f"one-{where}", body
    # The spacing swept across 100% inside the real panel, one paragraph so the
    # block stays inside the box. Below 100% the paragraph asks for LESS room
    # than the face's own line box and above it asks for more; if the shortfall
    # is one-sided, it shows up on one side of this sweep and not the other.
    for pct in (70000, 80000, 90000, 100000, 110000, 150000, 200000, 300000):
        body = re.sub(r"<a:p>.*</a:p>", one_line, panel, flags=re.S)
        body = re.sub(r'<a:spcPct val="\d+"/>', f'<a:spcPct val="{pct}"/>', body)
        yield f"pct-{pct // 1000}", body
    # And one at 80% in a SHORT box, since every synthetic arm that implied a
    # shortfall below 100% had a box under 62 pixels where this panel has 173.
    short = re.sub(r"<a:p>.*</a:p>", one_line, panel, flags=re.S)
    short = re.sub(r"<xdr:rowOff>2286000</xdr:rowOff>",
                   "<xdr:rowOff>1000000</xdr:rowOff>", short)
    yield "pct-80-short", short
    # The panel's own seven lines at percentages below 100, where the one-line
    # arms could not tell a shortfall from none because the shift never crossed
    # a rounding boundary. At seven lines it does.
    for pct in (70000, 90000):
        body = re.sub(r'<a:spcPct val="\d+"/>', f'<a:spcPct val="{pct}"/>', panel)
        yield f"seven-{pct // 1000}", body
    # One short paragraph, everything else as written.
    yield "one-para", re.sub(r"<a:p>.*</a:p>", one_line, panel, flags=re.S)
    # The body says nothing but how to wrap and where to anchor.
    yield "plain-body", re.sub(r"<a:bodyPr[^>]*/?>",
                               '<a:bodyPr wrap="square" anchor="ctr"/>', panel)
    # The runs say only face and size.
    yield "plain-runs", re.sub(
        r"<a:rPr[^>]*>.*?</a:rPr>",
        '<a:rPr lang="ja-JP" sz="1200"><a:latin typeface="Yu Gothic UI"/>'
        '<a:ea typeface="Yu Gothic UI"/></a:rPr>', panel, flags=re.S)
    # No paragraph properties at all.
    yield "no-pPr", re.sub(r"<a:pPr[^>]*>.*?</a:pPr>|<a:pPr[^>]*/>", "", panel, flags=re.S)
    # No shape transform: let the anchor place it.
    yield "no-xfrm", re.sub(r"<a:xfrm>.*?</a:xfrm>", "", panel, flags=re.S)
    # The same box written as a ONE-cell anchor — a corner and a size — which
    # is what every synthetic arm has been. The panel is a two-cell anchor and
    # that is the last thing separating them.
    fr = re.search(r"<xdr:from>.*?</xdr:from>", panel, re.S).group(0)
    to = re.search(r"<xdr:to>.*?</xdr:to>", panel, re.S).group(0)
    def at(part):
        col = int(re.search(r"<xdr:col>(\d+)</xdr:col>", part).group(1))
        coff = int(re.search(r"<xdr:colOff>(-?\d+)</xdr:colOff>", part).group(1))
        row = int(re.search(r"<xdr:row>(\d+)</xdr:row>", part).group(1))
        roff = int(re.search(r"<xdr:rowOff>(-?\d+)</xdr:rowOff>", part).group(1))
        return col, coff, row, roff
    fc, fco, frw, fro = at(fr)
    tc, tco, trw, tro = at(to)
    assert fc == tc and frw == trw, "the panel sits inside one cell both ways"
    body = panel.replace("<xdr:twoCellAnchor", "<xdr:oneCellAnchor")
    body = body.replace("</xdr:twoCellAnchor>", "</xdr:oneCellAnchor>")
    body = body.replace(to, f'<xdr:ext cx="{tco - fco}" cy="{tro - fro}"/>')
    yield "one-cell", body
    # The reverse test. Every arm so far kept `glossary_05`'s workbook, so
    # none of them could tell a property of the SHAPE from a property of the
    # BOOK. This drops a shape of the kind that demonstrably DOES take a
    # shortfall — one short line at 300%, centred, overflow — into the panel's
    # own anchor. If it refuses here, the book is what decides.
    made = (
        '<a:p><a:pPr algn="l"><a:lnSpc><a:spcPct val="300000"/></a:lnSpc></a:pPr>'
        '<a:r><a:rPr lang="ja-JP" sz="1200">'
        '<a:latin typeface="Yu Gothic UI"/><a:ea typeface="Yu Gothic UI"/>'
        '</a:rPr><a:t>A</a:t></a:r></a:p>'
    )
    body = re.sub(r"<a:p>.*</a:p>", made, panel, flags=re.S)
    body = re.sub(r"<a:bodyPr[^>]*/?>",
                  '<a:bodyPr wrap="square" anchor="ctr"/>', body)
    yield "synthetic", body
    # Excel drew nothing for the two arms that rewrote the RUNS, so those said
    # nothing. This keeps the panel's own runs and changes only the spacing to
    # 300%, where the shortfall is seven pixels and cannot be missed.
    body = re.sub(r'<a:spcPct val="80000"/>', '<a:spcPct val="300000"/>', panel)
    yield "at-300", body
    # And the same, cut to one paragraph, so the block is one line like the
    # arms that take the shortfall.
    body = re.sub(r"<a:p>.*</a:p>", one_line, panel, flags=re.S)
    body = re.sub(r'<a:spcPct val="80000"/>', '<a:spcPct val="300000"/>', body)
    yield "one-at-300", body


made = []
for name, body in arms():
    book = OUT / f"{name}.xlsx"
    book.unlink(missing_ok=True)
    # A fresh handle per book: reading one zip repeatedly while writing others
    # came back with a bad CRC part way through.
    with zipfile.ZipFile(SRC) as again,             zipfile.ZipFile(book, "w", zipfile.ZIP_DEFLATED) as out:
        for item in again.infolist():
            data = again.read(item.filename)
            if item.filename == "xl/drawings/drawing1.xml":
                data = (head + body + "</xdr:wsDr>").encode("utf-8")
            out.writestr(item, data)
    made.append((name, book, OUT / f"{name}.excel.png", OUT / f"{name}.oxi.png"))

listing = OUT / "_batch.txt"
listing.write_text("\n".join(f"{b}\t{p}" for _n, b, p, _o in made), encoding="utf-8-sig")
subprocess.run(["powershell", "-NoProfile", "-File",
                r"tools\metrics\_xlsx_screen_shot.ps1", "-ListFile", str(listing)],
               capture_output=True, text=True, encoding="utf-8",
               errors="replace", timeout=3600)

print(f"{'arm':<12}{'Excel tops':<30}{'ours tops':<30}")
for name, book, shot, ours in made:
    subprocess.run([str(RENDER), str(book), str(ours), "96"],
                   capture_output=True, timeout=1800)
    if not shot.exists() or not ours.exists():
        print(f"{name:<12}no picture")
        continue
    held = []
    for path in (shot, ours):
        a = np.asarray(Image.open(path).convert("L")).astype(int)
        sub = a[660:845, 50:730] < 128
        rows = np.where(sub.sum(axis=1) > 0)[0]
        bands, run = [], None
        for step in rows:
            if run is None or step > run[1] + 2:
                if run:
                    bands.append(run[0] + 660)
                run = [step, step]
            else:
                run[1] = step
        if run:
            bands.append(run[0] + 660)
        held.append(bands[:4])
    same = held[0] == held[1]
    print(f"{name:<12}{str(held[0]):<30}{str(held[1]):<30}"
          f"{'same -> no shortfall' if same else 'DIFFERS -> shortfall'}")
