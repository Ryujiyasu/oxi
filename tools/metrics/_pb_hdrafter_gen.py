# -*- coding: utf-8 -*-
"""Does Word apply the document default space-after (after=160 line=278, as this
machine's Normal.dotm gives a pPrDefault-less package) to HEADER paragraphs --
between them, below the last one, and to an inline-picture paragraph?

2026-10-13 probeqhdrimg (header = inline picture 113.4x85pt + one text line,
re-measured under those defaults): Word's header text line sits 6.55pt lower
than Oxi's and the body 12.5pt lower -- about one 8pt after at each of the two
paragraph ends. Arms (explicit pPrDefault, ＭＳ 明朝 10.5pt, lines 360, header
distance 851):
    text1   one header paragraph
    text2   two header paragraphs
    img     picture paragraph + text paragraph (the probeqhdrimg shape)
    img_only picture paragraph alone
    text2_noafter  two paragraphs, pPrDefault after=0 (control)
Read: header line baselines and the first body baseline (Word PDF via fitz;
Oxi via --dump-glyphs).

    python tools/metrics/_pb_hdrafter_gen.py gen
    python tools/metrics/_pb_hdrafter_gen.py word
    python tools/metrics/_pb_hdrafter_gen.py oxi <renderer.exe>
"""
import json, os, shutil, struct, subprocess, sys, tempfile, zipfile, zlib
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_txbxfit_gen as T  # noqa: E402
import line_census as L  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_hdrafter"
RPR = '<w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:sz w:val="21"/></w:rPr>'
# only the namespaces T.NS does not already declare (a duplicate xmlns is "file appears corrupt")
NSW = " ".join(f'xmlns:{p}="{u}"' for p, u in (("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"),
                                              ("a", "http://schemas.openxmlformats.org/drawingml/2006/main"),
                                              ("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture"),
                                              ("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")) if f"xmlns:{p}=" not in T.NS)
ARMS = {"text1": (1, False, 160), "text2": (2, False, 160), "img": (1, True, 160), "img_only": (0, True, 160), "text2_noafter": (2, False, 0),
        # footer arms: the same two paragraphs in the FOOTER (distance 992); read the footer lines and the last body line
        "ftr_text2": (2, False, 160), "ftr_text2_noafter": (2, False, 0)}


def styles(after):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + T.NS + '><w:docDefaults><w:rPrDefault>' + RPR +
            f'</w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="{after}" w:line="278" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
            '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:pPr><w:widowControl w:val="0"/></w:pPr></w:style>'
            '</w:styles>')


def png(w=120, h=90):
    raw = b"".join(b"\x00" + bytes([40, 80, 160] * w) for _ in range(h))
    def chunk(t, d):
        return struct.pack(">I", len(d)) + t + d + struct.pack(">I", zlib.crc32(t + d) & 0xffffffff)
    return b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0)) + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b"")


def para(text):
    return f'<w:p><w:pPr><w:jc w:val="left"/>{RPR}</w:pPr><w:r>{RPR}<w:t xml:space="preserve">{text}</w:t></w:r></w:p>'


def img_para():
    return ('<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="1440000" cy="1080000"/>'
            '<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="300" name="Pic300"/><wp:cNvGraphicFramePr/>'
            '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>'
            '<pic:nvPicPr><pic:cNvPr id="300" name="Pic300"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1440000" cy="1080000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
            '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>')


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    body = "".join(para(f"第{i}条　本項に定める事項については、関係法令及び本規程の趣旨に照らし、善良なる管理者の注意をもって。") for i in range(1, 6))
    for arm, (ntext, with_img, after) in ARMS.items():
        is_ftr = arm.startswith("ftr_")
        tag, part, reltype = ("ftr", "footer1.xml", "footer") if is_ftr else ("hdr", "header1.xml", "header")
        hdr_paras = (img_para() if with_img else "") + "".join(para(f"{'フッタ' if is_ftr else 'ヘッダ'}行{i}") for i in range(1, ntext + 1))
        header = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:' + tag + ' ' + T.NS + " " + NSW + ">" + hdr_paras + "</w:" + tag + ">"
        body_here = body * 3 if is_ftr else body  # footer arms: fill the page so the last body line meets the footer
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + T.NS + " " + NSW + "><w:body>" + body_here +
               f'<w:sectPr><w:{reltype}Reference w:type="default" r:id="rIdHdr"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1418" w:right="1418" w:bottom="1418"'
               ' w:left="1418" w:header="851" w:footer="992" w:gutter="0"/><w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr></w:body></w:document>')
        ct = T.CT.replace("</Types>", f'<Default Extension="png" ContentType="image/png"/><Override PartName="/word/{part}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.{reltype}+xml"/></Types>')
        drels = T.DRELS.replace("</Relationships>", f'<Relationship Id="rIdHdr" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/{reltype}" Target="{part}"/></Relationships>')
        hrels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                 '<Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/></Relationships>')
        with zipfile.ZipFile(OUT / f"{arm}.docx", "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", T.RELS)
            z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/styles.xml", styles(after))
            z.writestr("word/" + part, header); z.writestr("word/_rels/" + part + ".rels", hrels)
            z.writestr("word/media/image1.png", png()); z.writestr("word/document.xml", doc)
    print("ok", len(ARMS))


def lines_of(items):
    ls = [(round(l["y"], 2), l["t"][:10]) for l in L.group_lines(items)]
    return ls if len(ls) <= 4 else ls[:2] + ls[-3:]


def word_lines(pdf):
    import fitz
    page = fitz.open(pdf)[0]; items = []
    for b in page.get_text("rawdict")["blocks"]:
        for l in b.get("lines", []):
            for s in l["spans"]:
                for c in s["chars"]:
                    if c["c"].strip():
                        items.append((c["origin"][1], c["origin"][0], c["bbox"][2], c["c"]))
    return lines_of(items)


def oxi_lines(exe, docx):
    t = tempfile.mkdtemp(); out = os.path.join(t, "g.json")
    subprocess.run([exe, str(docx), os.path.join(t, "p"), "--dump-glyphs=" + out], capture_output=True, timeout=600)
    d = json.load(open(out, encoding="utf-8")); items = []
    for g in d["pages"][0]["glyphs"]:
        fs = g.get("font_size", 0) or 0
        base = g.get("baseline")
        if base is None:
            base = g["top"] + L.win_ascent(g.get("font_family", "")) * fs
        items.append((base, g["x"], g["x"] + fs, g["char"]))
    return lines_of(items)


def word():
    import win32com.client
    w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
    try:
        for arm in ARMS:
            t = os.path.join(tempfile.mkdtemp(), arm + ".docx"); shutil.copy(OUT / f"{arm}.docx", t)
            d = w.Documents.Open(t, ReadOnly=False, AddToRecentFiles=False)
            d.SaveAs2(str((OUT / f"{arm}.pdf").resolve()), 17); d.Close(0)
    finally:
        w.Quit()
    res = {arm: word_lines(OUT / f"{arm}.pdf") for arm in ARMS}
    for arm, v in res.items():
        print(f"{arm:14s} word {v}", flush=True)
    (OUT / "word.json").write_text(json.dumps(res, ensure_ascii=False), encoding="utf-8")


def oxi(exe):
    wd = json.loads((OUT / "word.json").read_text(encoding="utf-8"))
    for arm in ARMS:
        o = oxi_lines(os.path.abspath(exe), OUT / f"{arm}.docx")
        print(f"{arm:14s} oxi {o}\n{'':14s} word {wd[arm]}", flush=True)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()
