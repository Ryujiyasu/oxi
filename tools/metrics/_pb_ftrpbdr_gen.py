# -*- coding: utf-8 -*-
"""Does a footer paragraph's top/bottom border (pBdr width + space) count in
the footer's height, i.e. move the text up from the footer line?

blind-G en/creative__03bf8c5f: footer = one Calibri 9 paragraph with
<w:pBdr><w:top sz=6 space=3/> (after 0). Word's baseline is 804.24 = line
bottom ON the footer line (841.9 - 35.4); Oxi's is 807.83: the 3.75pt border
room was added below the stack instead of inside it.

Arms: <bdr>_<where>: bdr = none / top (sz6 space3) / bot / both / topbig
(sz 24 space 12); where = ftr / hdr. Footer text «Footer line»; header text
«Header line»; body 40 lines.  Read: the footer/header baseline, the first
and last body baselines (page 1).

    python tools/metrics/_pb_ftrpbdr_gen.py gen
    python tools/metrics/_pb_ftrpbdr_gen.py word
    python tools/metrics/_pb_ftrpbdr_gen.py oxi <renderer.exe>
"""
import json, os, shutil, sys, tempfile, zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_txbxfit_gen as T  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_ftrpbdr"
TOP = '<w:top w:val="single" w:sz="{sz}" w:space="{sp}" w:color="365F91"/>'
BOT = '<w:bottom w:val="single" w:sz="{sz}" w:space="{sp}" w:color="365F91"/>'
BDRS = {
    "none": "",
    "top": "<w:pBdr>" + TOP.format(sz=6, sp=3) + "</w:pBdr>",
    "bot": "<w:pBdr>" + BOT.format(sz=6, sp=3) + "</w:pBdr>",
    "both": "<w:pBdr>" + TOP.format(sz=6, sp=3) + BOT.format(sz=6, sp=3) + "</w:pBdr>",
    "topbig": "<w:pBdr>" + TOP.format(sz=24, sp=12) + "</w:pBdr>",
}
WHERES = ("ftr", "hdr")
ARMS = [f"{b}_{w}" for b in BDRS for w in WHERES]
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + T.NS + '><w:docDefaults>'
          '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/>'
          '<w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="de-CH"/></w:rPr></w:rPrDefault>'
          '<w:pPrDefault><w:pPr><w:spacing w:after="200" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
          '</w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>'
          '<w:style w:type="paragraph" w:styleId="Fuzeile"><w:name w:val="footer"/><w:basedOn w:val="Normal"/>'
          '<w:pPr><w:tabs><w:tab w:val="center" w:pos="4536"/><w:tab w:val="right" w:pos="9072"/></w:tabs>'
          '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:style>'
          '</w:styles>')


def part(root, bdr, text):
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:{root} {T.NS}><w:p><w:pPr><w:pStyle w:val="Fuzeile"/>'
            f'{BDRS[bdr]}<w:rPr><w:sz w:val="18"/></w:rPr></w:pPr><w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:t>{text}</w:t></w:r></w:p></w:{root}>')


def doc_xml(where):
    body = "".join(f'<w:p><w:r><w:t>Body line {i + 1} of the filler text that keeps the page full.</w:t></w:r></w:p>' for i in range(40))
    ref = ('<w:headerReference w:type="default" r:id="rIdH"/>' if where == "hdr" else '<w:footerReference w:type="default" r:id="rIdF"/>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + T.NS + "><w:body>" + body +
            '<w:sectPr>' + ref + '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1134"'
            ' w:left="1417" w:header="708" w:footer="708" w:gutter="0"/><w:cols w:space="708"/><w:docGrid w:linePitch="360"/>'
            '</w:sectPr></w:body></w:document>')


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    for arm in ARMS:
        bdr, where = arm.split("_")
        name, root, rid, ctype, text = (("header1", "hdr", "rIdH", "header", "Header line") if where == "hdr"
                                        else ("footer1", "ftr", "rIdF", "footer", "Footer line"))
        ct = T.CT.replace("</Types>", f'<Override PartName="/word/{name}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.{ctype}+xml"/></Types>')
        drels = T.DRELS.replace("</Relationships>", f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/{ctype}" Target="{name}.xml"/></Relationships>')
        with zipfile.ZipFile(OUT / f"{arm}.docx", "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", T.RELS)
            z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/styles.xml", STYLES)
            z.writestr(f"word/{name}.xml", part(root, bdr, text)); z.writestr("word/document.xml", doc_xml(where))
    print("ok", len(ARMS))


def _read(lines):
    hf = [l["y"] for l in lines if l["t"].startswith("Footer") or l["t"].startswith("Header")]
    body = sorted(l["y"] for l in lines if l["t"].startswith("Body"))
    return {"hf": round(hf[0], 2) if hf else None, "body_first": round(body[0], 2) if body else None,
            "body_last": round(body[-1], 2) if body else None, "n_body": len(body)}


def word():
    import win32com.client, line_census as L
    w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
    try:
        for arm in ARMS:
            t = os.path.join(tempfile.mkdtemp(), arm + ".docx"); shutil.copy(OUT / f"{arm}.docx", t)
            d = w.Documents.Open(t, ReadOnly=False, AddToRecentFiles=False)
            d.SaveAs2(str((OUT / f"{arm}.pdf").resolve()), 17); d.Close(0)
    finally:
        w.Quit()
    res = {}
    for arm in ARMS:
        res[arm] = _read(L.word_lines(OUT / f"{arm}.pdf"))
        print(f"{arm:12s} word {res[arm]}", flush=True)
    (OUT / "word.json").write_text(json.dumps(res), encoding="utf-8")


def oxi(exe):
    import line_census as L
    word = json.loads((OUT / "word.json").read_text(encoding="utf-8"))
    for arm in ARMS:
        o = _read(L.oxi_lines(os.path.abspath(exe), OUT / f"{arm}.docx"))
        wd = word.get(arm) or {}
        flag = "" if o["hf"] is not None and wd.get("hf") is not None and abs(o["hf"] - wd["hf"]) < 0.6 else "  <-- differs"
        print(f"{arm:12s} oxi {o}  word {wd}{flag}", flush=True)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()
