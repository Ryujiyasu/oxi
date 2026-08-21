# -*- coding: utf-8 -*-
"""Minimal repro for "does an inline PICTURE share the line of another INLINE
OBJECT in the same paragraph?"

The witness is correspondence__000407cd (Word 1 page / Oxi 2): one paragraph
holds a real inline OLE (v:shape 106x49.3pt, ProgID RM.ColourMagic.2), 66
whitespace chars, and a wp:inline picture (45.9x44.2pt). Word draws them on ONE
line; Oxi's S854 routing only counts PICTURES, so with n_pictures == 1 the
picture fell to SPLIT-BLOCK and got its own line (+44pt -> +1 page).

The repro reuses that document's OWN package (media + embeddings + rels +
styles + theme) so the object is a REAL OLE, and swaps in a body of isolated
arms.  Each arm is bracketed by marker paragraphs so the Word PDF gives the
arm's total advance and the picture/object x positions.

  python tools/metrics/_pb_inlineobj_gen.py            # write the repro docx
  python tools/metrics/_pb_inlineobj_gen.py --measure  # + Word PDF truth
  python tools/metrics/_pb_inlineobj_gen.py --oxi      # + Oxi's answer
"""
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
SRC = REPO / "pipeline_data" / "docx_corpus" / "en" / "correspondence" / "000407cd6c79442c.docx"
OUT = Path(os.environ.get("OXI_SCRATCH", tempfile.gettempdir())) / "pb_inlineobj.docx"

PIC_W_EMU, PIC_H_EMU = 582930, 561340          # 45.90 x 44.20 pt
PIC_W_PT, PIC_H_PT = PIC_W_EMU / 12700.0, PIC_H_EMU / 12700.0
OBJ_W_PT, OBJ_H_PT = 106.0, 49.3

SHAPETYPE = (
    '<v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75"'
    ' o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f">'
    '<v:stroke joinstyle="miter"/><v:formulas>'
    '<v:f eqn="if lineDrawn pixelLineWidth 0"/><v:f eqn="sum @0 1 0"/>'
    '<v:f eqn="sum 0 0 @1"/><v:f eqn="prod @2 1 2"/>'
    '<v:f eqn="prod @3 21600 pixelWidth"/><v:f eqn="prod @3 21600 pixelHeight"/>'
    '<v:f eqn="sum @0 0 1"/><v:f eqn="prod @6 1 2"/>'
    '<v:f eqn="prod @7 21600 pixelWidth"/><v:f eqn="sum @8 21600 0"/>'
    '<v:f eqn="prod @7 21600 pixelHeight"/><v:f eqn="sum @10 21600 0"/>'
    "</v:formulas>"
    '<v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/>'
    '<o:lock v:ext="edit" aspectratio="t"/></v:shapetype>'
)


def obj_run(idx: int, w=OBJ_W_PT, h=OBJ_H_PT, ole=True) -> str:
    """A real inline OLE object (v:shape + o:OLEObject) drawn from image4.emf.

    `ole=False` drops <o:OLEObject> = the S851 "OLE-less" bare picture shape.
    """
    shape_id = "_x0000_i10%02d" % idx
    body = (
        '<v:shape id="%s" type="#_x0000_t75" style="width:%gpt;height:%gpt" o:ole="">'
        '<v:imagedata r:id="rId12" o:title=""/></v:shape>' % (shape_id, w, h)
    )
    if ole:
        body += (
            '<o:OLEObject Type="Embed" ProgID="RM.ColourMagic.2" ShapeID="%s"'
            ' DrawAspect="Content" ObjectID="_179318304%d" r:id="rId13"/>' % (shape_id, idx)
        )
    return (
        "<w:r><w:object w:dxaOrig=\"11280\" w:dyaOrig=\"6255\">"
        + (SHAPETYPE if idx == 0 else "")
        + body
        + "</w:object></w:r>"
    )


def pic_run(idx: int, w_emu=PIC_W_EMU, h_emu=PIC_H_EMU) -> str:
    """A wp:inline picture (media/image5.png)."""
    return (
        '<w:r><w:rPr><w:noProof/></w:rPr><w:drawing>'
        '<wp:inline distT="0" distB="0" distL="0" distR="0">'
        '<wp:extent cx="%d" cy="%d"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
        '<wp:docPr id="%d" name="Picture %d"/><wp:cNvGraphicFramePr/>'
        '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<pic:nvPicPr><pic:cNvPr id="0" name="Picture %d"/><pic:cNvPicPr/></pic:nvPicPr>'
        '<pic:blipFill><a:blip r:embed="rId14"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        '<pic:spPr bwMode="auto"><a:xfrm><a:off x="0" y="0"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
        "</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>"
        % (w_emu, h_emu, 900 + idx, 900 + idx, 900 + idx, w_emu, h_emu)
    )


def txt_run(s: str, sz: int = 16) -> str:
    return (
        '<w:r><w:rPr><w:b/><w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr>'
        '<w:t xml:space="preserve">%s</w:t></w:r>' % (sz, sz, s)
    )


SPACES = " " * 66

# (name, run-builder) — the run builder gets a per-arm index for unique ids.
ARMS = [
    ("pic_alone", lambda i: pic_run(i)),
    ("obj_alone", lambda i: obj_run(i)),
    ("obj_pic", lambda i: obj_run(i) + pic_run(i)),
    ("obj_spaces_pic", lambda i: obj_run(i) + txt_run(SPACES) + pic_run(i)),
    ("pic_obj", lambda i: pic_run(i) + obj_run(i)),
    ("pic_text", lambda i: pic_run(i) + txt_run("TEXT")),
    ("obj_text", lambda i: obj_run(i) + txt_run("TEXT")),
    ("pic_pic", lambda i: pic_run(i) + pic_run(i + 50)),
    ("obj_pic_pic", lambda i: obj_run(i) + pic_run(i) + pic_run(i + 50)),
    # width sweep around the 451.3pt column (11906 - 2*1440 tw):
    #   400.0 + 45.9 = 445.9 <= 451.3  -> one line
    #   380.0 + 45.9 = 425.9 <= 451.3  -> one line
    #   440.0 + 45.9 = 485.9 >  451.3  -> the picture must wrap to line 2
    ("objwide_pic", lambda i: obj_run(i, w=400.0, h=20.0) + pic_run(i)),
    ("objfit_pic", lambda i: obj_run(i, w=380.0, h=20.0) + pic_run(i)),
    ("objover_pic", lambda i: obj_run(i, w=440.0, h=20.0) + pic_run(i)),
    # OLE-less <w:object> (the S851 shape) + picture
    ("objnoole_pic", lambda i: obj_run(i, ole=False) + pic_run(i)),
    ("pic_spaces_pic", lambda i: pic_run(i) + txt_run(SPACES) + pic_run(i + 50)),
]

HDR = None


def mark(i: int, side: str) -> str:
    return "ZMARK%s%02dZ" % (side, i)


def build() -> Path:
    with zipfile.ZipFile(SRC) as z:
        names = z.namelist()
        parts = {n: z.read(n) for n in names}
    doc = parts["word/document.xml"].decode("utf8")
    global HDR
    HDR = doc[: doc.index("<w:body>")]
    sect = re.search(r"<w:sectPr.*?</w:sectPr>", doc, re.S).group(0)

    body = []
    for i, (name, mk) in enumerate(ARMS):
        # ★Marker text must be HYPHEN-FREE: Oxi (like Word) breaks a Latin word
        # after a hyphen, so "MARK-arm-A" arrives as three fragments and the
        # arm lookup silently misses every row.
        body.append("<w:p>" + txt_run(mark(i, "A"), sz=22) + "</w:p>")
        body.append("<w:p>" + mk(i) + "</w:p>")
        body.append("<w:p>" + txt_run(mark(i, "B"), sz=22) + "</w:p>")
    # ★sectPr belongs to the BODY, not inside a bare <w:p> (a paragraph-level
    # sectPr must sit in <w:pPr>). The invalid form made Word fall back to its
    # own page setup and every x was 13.1pt off the margin the file states.
    out_doc = HDR + "<w:body>" + "".join(body) + sect + "</w:body></w:document>"

    OUT.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
        for n in names:
            z.writestr(n, out_doc.encode("utf8") if n == "word/document.xml" else parts[n])
    print("wrote", OUT)
    return OUT


def measure_word(path: Path):
    pdf = Path(tempfile.gettempdir()) / (path.stem + ".truth.pdf")
    import win32com.client as win32
    w = win32.DispatchEx("Word.Application")
    w.Visible = False
    try:
        d = w.Documents.Open(str(path), ReadOnly=True)
        d.ExportAsFixedFormat(str(pdf), 17)
        d.Close(False)
    finally:
        w.Quit()
    return report(pdf, "WORD")


def measure_oxi(path: Path):
    exe = REPO / "tools" / "oxi-gdi-renderer" / "target" / "release" / "oxi-gdi-renderer.exe"
    tmp = Path(tempfile.mkdtemp())
    dump = tmp / "d.json"
    subprocess.run([str(exe), str(path), str(tmp / "p"), "110", "--dump-layout=%s" % dump],
                   check=True, capture_output=True)
    import json
    d = json.load(open(dump, encoding="utf-8"))
    rows = []
    for pi, pg in enumerate(d["pages"]):
        for e in pg["elements"]:
            y = e.get("y", 0.0) + pi * 10000
            if e.get("type") == "text":
                rows.append((y, "T", e.get("x", 0.0), (e.get("text") or "")))
            elif e.get("type") == "image":
                rows.append((y, "I", e.get("x", 0.0), "%.1fx%.1f" % (e.get("w", 0), e.get("h", 0))))
    return summarize(rows, "OXI")


def report(pdf: Path, tag: str):
    import fitz
    doc = fitz.open(pdf)
    rows = []
    for pi in range(doc.page_count):
        pg = doc[pi]
        for blk in pg.get_text("dict")["blocks"]:
            if blk.get("type", 0) != 0:
                continue
            for ln in blk.get("lines", []):
                t = "".join(s["text"] for s in ln["spans"])
                y = min(s["bbox"][1] for s in ln["spans"]) + pi * 10000
                rows.append((y, "T", min(s["bbox"][0] for s in ln["spans"]), t))
        for im in pg.get_image_info():
            b = im["bbox"]
            rows.append((b[1] + pi * 10000, "I", b[0], "%.1fx%.1f" % (b[2] - b[0], b[3] - b[1])))
    return summarize(rows, tag)


def summarize(rows, tag):
    rows.sort()
    out = {}
    print("--- %s ---" % tag)
    for idx, (name, _) in enumerate(ARMS):
        ya = [r[0] for r in rows if r[1] == "T" and mark(idx, "A") in r[3]]
        yb = [r[0] for r in rows if r[1] == "T" and mark(idx, "B") in r[3]]
        if not ya or not yb:
            print("  %-16s MISSING" % name)
            continue
        a, b = min(ya), min(yb)
        mid = [r for r in rows if a < r[0] < b]
        # distinct y bands inside the arm, merged under 3pt
        ys = sorted({round(r[0], 1) for r in mid})
        bands = []
        for y in ys:
            if not bands or y - bands[-1] > 3.0:
                bands.append(y)
        objs = ["%s@%.0f,%.0f" % (r[3], r[2], r[0]) for r in mid if r[1] == "I"]
        print("  %-16s advance %7.2f  bands %d  %s" % (name, b - a, len(bands), " ".join(objs)))
        out[name] = (b - a, len(bands))
    return out


if __name__ == "__main__":
    p = build()
    w = measure_word(p) if "--measure" in sys.argv else None
    o = measure_oxi(p) if "--oxi" in sys.argv else None
    if w and o:
        print("--- DIFF (oxi - word) ---")
        for name, _mk in ARMS:
            if name in w and name in o:
                print("  %-16s d_advance %+8.2f  bands %d vs %d %s"
                      % (name, o[name][0] - w[name][0], o[name][1], w[name][1],
                         "MISMATCH" if abs(o[name][0] - w[name][0]) > 1.0 else ""))
