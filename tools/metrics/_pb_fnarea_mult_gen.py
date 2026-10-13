# -*- coding: utf-8 -*-
"""Word's footnote AREA under document defaults after=160 line=278 on a lines-360
grid (the pPrDefault-less golden probefn re-measured under this machine's
Normal.dotm): Oxi's area is 25.35pt taller than Word's (first note 585.87 vs
611.71) although the note pitch (28.85 = 18 x 278/240 + 8) agrees.

Arms: n footnotes {1, 2, 3, 6} x fnPos {pageBottom (default), beneathText};
the separator paragraph keeps the golden shape (after=0 line=240), notes are
ＭＳ 明朝 9pt (sz 18) with no direct spacing (= the defaults). Body: 40
three-line 10.5pt paragraphs so page 1 overflows; notes are referenced from
paragraphs 1..n (always on page 1).
Read (Word PDF via fitz): the last body baseline on page 1, the separator rule
(a drawn line) y, each note's baseline; Oxi via --dump-layout (the same items).

    python tools/metrics/_pb_fnarea_mult_gen.py gen
    python tools/metrics/_pb_fnarea_mult_gen.py word
    python tools/metrics/_pb_fnarea_mult_gen.py oxi <renderer.exe>
"""
import json, os, shutil, subprocess, sys, tempfile, zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_txbxfit_gen as T  # noqa: E402
import line_census as L  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_fnarea_mult"
RPR = '<w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:sz w:val="21"/></w:rPr>'
RPR9 = RPR.replace('w:val="21"', 'w:val="18"')
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + T.NS + '><w:docDefaults><w:rPrDefault>' + RPR +
          '</w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="278" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:pPr><w:widowControl w:val="0"/></w:pPr></w:style>'
          '</w:styles>')
TEXT = ("本項に定める事項については、関係法令及び本規程の趣旨に照らし、善良なる管理者の注意をもって、誠実かつ適切に取り扱うものとし、"
        "これに関連して生じた一切の事項は、別に定めるところにより、遅滞なく、かつ、公正に処理しなければならない。")
# (n, fnPos, pPrDefault spacing): the after / multiple roles separated by single-spacing control arms
SPACINGS = {"bottom": (160, 278), "beneath": (160, 278), "single": (0, 240), "after0": (0, 278), "mult1": (160, 240)}
ARMS = [(n, pos) for pos in ("bottom", "beneath") for n in (1, 2, 3, 6)] + [(n, pos) for pos in ("single", "after0", "mult1") for n in (1, 3, 6)]


def arm_name(n, pos):
    return f"{pos}_{n}"


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    for n, pos in ARMS:
        paras = []
        for i in range(1, 41):
            ref = (f'<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="{i + 1}"/></w:r>' if i <= n else "")
            paras.append(f'<w:p><w:pPr><w:jc w:val="both"/>{RPR}</w:pPr><w:r>{RPR}<w:t xml:space="preserve">第{i}条　{TEXT}</w:t></w:r>{ref}</w:p>')
        fnpr = '<w:footnotePr><w:pos w:val="beneathText"/></w:footnotePr>' if pos == "beneath" else ""
        after, line = SPACINGS[pos]
        styles = STYLES.replace('w:after="160" w:line="278"', f'w:after="{after}" w:line="{line}"')
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + T.NS + "><w:body>" + "".join(paras) +
               '<w:sectPr>' + fnpr + '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1418" w:right="1418" w:bottom="1418"'
               ' w:left="1418" w:header="851" w:footer="992" w:gutter="0"/><w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr></w:body></w:document>')
        notes = ['<w:footnote w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:footnote>',
                 '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>']
        for i in range(1, n + 1):
            notes.append(f'<w:footnote w:id="{i + 1}"><w:p><w:pPr>{RPR9}</w:pPr><w:r>{RPR9.replace("</w:rPr>", "<w:vertAlign w:val=\"superscript\"/></w:rPr>")}<w:footnoteRef/></w:r>'
                         f'<w:r>{RPR9}<w:t xml:space="preserve">注記{i}：本条の適用に関する補足説明であって、実務上の取扱いを示すものである。</w:t></w:r></w:p></w:footnote>')
        footnotes = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes ' + T.NS + ">" + "".join(notes) + "</w:footnotes>"
        ct = T.CT.replace("</Types>", '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>')
        drels = T.DRELS.replace("</Relationships>", '<Relationship Id="rIdFn" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>')
        with zipfile.ZipFile(OUT / f"{arm_name(n, pos)}.docx", "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", T.RELS)
            z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/styles.xml", styles)
            z.writestr("word/footnotes.xml", footnotes); z.writestr("word/document.xml", doc)
    print("ok", len(ARMS))


def read_word(pdf):
    import fitz
    page = fitz.open(pdf)[0]; items = []
    for b in page.get_text("rawdict")["blocks"]:
        for l in b.get("lines", []):
            for s in l["spans"]:
                for c in s["chars"]:
                    if c["c"].strip():
                        items.append((c["origin"][1], c["origin"][0], c["bbox"][2], c["c"], round(s["size"], 1)))
    lines = L.group_lines([it[:4] for it in items])
    small = sorted({round(it[0], 2) for it in items if it[4] < 9.5 and it[3] not in "0123456789"})
    body = [l["y"] for l in lines if l["y"] not in small]
    rules = sorted({round(d["rect"].y0, 2) for d in page.get_drawings() if d["rect"].width > 50 and d["rect"].height < 3})
    return {"last_body": round(max(body), 2) if body else None, "rule": rules, "notes": small}


def read_oxi(exe, docx):
    """Baselines from --dump-glyphs (the renderer's own drawn baseline); the rule from --dump-layout."""
    t = tempfile.mkdtemp(); out = os.path.join(t, "g.json"); lay = os.path.join(t, "l.json")
    subprocess.run([exe, str(docx), os.path.join(t, "p"), "--dump-glyphs=" + out], capture_output=True, timeout=600)
    subprocess.run([exe, str(docx), os.path.join(t, "q"), "--dump-layout=" + lay], capture_output=True, timeout=600)
    g = json.load(open(out, encoding="utf-8"))["pages"][0]["glyphs"]
    items = [(gg.get("baseline") or gg["top"] + L.win_ascent(gg.get("font_family", "")) * gg["font_size"], gg["x"], gg["x"] + gg["font_size"], gg["char"], gg["font_size"]) for gg in g]
    lines = L.group_lines([it[:4] for it in items])
    small = sorted({round(it[0], 2) for it in items if it[4] < 9.5 and it[3] not in "0123456789"})
    body = [l["y"] for l in lines if round(l["y"], 2) not in small]
    pg = json.load(open(lay, encoding="utf-8"))["pages"][0]
    rules = sorted({round(e["y"], 2) for e in pg["elements"] if e["type"] not in ("text",) and (e.get("width") or 0) > 50 and (e.get("height") or 0) <= 2})
    return {"last_body": round(max(body), 2) if body else None, "rule": rules, "notes": small}


def word():
    import win32com.client
    w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
    try:
        for n, pos in ARMS:
            arm = arm_name(n, pos)
            if (OUT / f"{arm}.pdf").exists():
                continue
            t = os.path.join(tempfile.mkdtemp(), arm + ".docx"); shutil.copy(OUT / f"{arm}.docx", t)
            d = w.Documents.Open(t, ReadOnly=False, AddToRecentFiles=False)
            d.SaveAs2(str((OUT / f"{arm}.pdf").resolve()), 17); d.Close(0)
    finally:
        w.Quit()
    res = {}
    for n, pos in ARMS:
        arm = arm_name(n, pos); res[arm] = read_word(OUT / f"{arm}.pdf")
        print(f"{arm:10s} word {res[arm]}", flush=True)
    (OUT / "word.json").write_text(json.dumps(res), encoding="utf-8")


def oxi(exe):
    wd = json.loads((OUT / "word.json").read_text(encoding="utf-8"))
    for n, pos in ARMS:
        arm = arm_name(n, pos)
        print(f"{arm:10s} oxi  {read_oxi(os.path.abspath(exe), OUT / f'{arm}.docx')}\n{'':10s} word {wd[arm]}", flush=True)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()
