# -*- coding: utf-8 -*-
"""Which rFonts slot does Word draw each character from?

The face census (tools/metrics/face_census.py) leaves mostly slot choices:
Roman numerals Ⅰ-Ⅷ and “ drawn from eastAsia by Word, from ascii by Oxi
(ja/educational__20d9968b and two more); the en dash – from ascii by Word, from
eastAsia by Oxi (ja/reports__5823d5a8); ※ (ja/forms__01e60007); Latin words in
ja/reports__597cb86c.

Arms h<hint>_l<lang>: one paragraph per test character, each in a run with
ascii/hAnsi = Arial and eastAsia = ＭＳ 明朝 (two faces whose PDF programs are
told apart), hint = none / eastAsia; the run's w:lang eastAsia = ja-JP /
en-US.  Characters: Roman numerals, ※, dashes, quotes, circled numbers, Greek,
Cyrillic, ×÷°±§¶, arrows, box drawing, ・, …, ‰, €, fullwidth/halfwidth forms,
a Latin letter control and a kanji control.  Read: the face of every character
(Word PDF via its embedded programs; Oxi glyph dump).

    python tools/metrics/_pb_slot_class_gen.py gen
    python tools/metrics/_pb_slot_class_gen.py word
    python tools/metrics/_pb_slot_class_gen.py oxi <renderer.exe>
"""
import json, os, shutil, subprocess, sys, tempfile, zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_txbxfit_gen as T  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_slot_class"
CHARS = ("A", "漢", "Ⅰ", "Ⅷ", "ⅰ", "※", "–", "—", "―", "‐", "“", "”", "‘", "’", "…", "‥", "・", "①", "⑳",
         "α", "Ω", "Б", "я", "×", "÷", "°", "±", "§", "¶", "→", "←", "↑", "⇒", "─", "│", "■", "□", "○", "●",
         "◆", "★", "☆", "‰", "€", "£", "¥", "©", "®", "™", "℃", "№", "Å", "é", "ü", "ﾃ", "Ａ", "１", "！")
ARMS = {f"h{h}_l{l}": (h, l) for h in ("none", "ea") for l in ("ja", "en")}
MS = "ＭＳ 明朝"


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    for arm, (h, l) in ARMS.items():
        hint = ' w:hint="eastAsia"' if h == "ea" else ""
        lang = "ja-JP" if l == "ja" else "en-US"
        rpr = (f'<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="{MS}" w:cs="Arial"{hint}/>'
               f'<w:sz w:val="22"/><w:lang w:val="en-US" w:eastAsia="{lang}"/></w:rPr>')
        body = "".join(f'<w:p><w:r>{rpr}<w:t>{c}</w:t></w:r></w:p>' for c in CHARS)
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + T.NS + '><w:body>' + body +
               '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="720" w:right="1440" w:bottom="720" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>'
               '<w:cols w:num="3" w:space="708"/></w:sectPr></w:body></w:document>')
        styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + T.NS + '><w:docDefaults><w:rPrDefault><w:rPr>'
                  '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="' + MS + '" w:cs="Arial"/><w:sz w:val="22"/><w:lang w:val="en-US" w:eastAsia="ja-JP"/>'
                  '</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
                  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>')
        settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + T.NS + '><w:compat><w:compatSetting w:name="compatibilityMode" '
                    'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
        ct = T.CT.replace("</Types>", '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
        drels = T.DRELS.replace("</Relationships>", '<Relationship Id="rIdSet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
        with zipfile.ZipFile(OUT / f"{arm}.docx", "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", T.RELS)
            z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/styles.xml", styles)
            z.writestr("word/settings.xml", settings); z.writestr("word/document.xml", doc)
    print("ok", len(ARMS))


def short(fam):
    f = (fam or "?").lstrip("=")
    return "A" if f.startswith("Arial") else "M" if ("Mincho" in f or "明朝" in f) else f[:10]


def word():
    import win32com.client, fitz
    import face_census as F
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
        got = {}
        for pi, c, x, y, face in F.word_chars(str(OUT / f"{arm}.pdf"), 3):
            got.setdefault(c, short(face))
        res[arm] = got
    (OUT / "word.json").write_text(json.dumps(res, ensure_ascii=False), encoding="utf-8")
    table(res, None)


def table(word, oxi):
    print("char  " + "  ".join(f"{a:>8s}" for a in ARMS))
    for c in CHARS:
        cells = []
        for a in ARMS:
            wv = word[a].get(c, "-")
            if oxi is None:
                cells.append(f"{wv:>8s}")
            else:
                ov = oxi[a].get(c, "-")
                cells.append(f"{wv + ('' if ov == wv else '/' + ov):>8s}")
        print(f"U+{ord(c):04X} {c}  " + "  ".join(cells))


def oxi(exe):
    word = json.loads((OUT / "word.json").read_text(encoding="utf-8"))
    res = {}
    for arm in ARMS:
        t = tempfile.mkdtemp(); go = os.path.join(t, "g.json")
        subprocess.run([os.path.abspath(exe), str(OUT / f"{arm}.docx"), os.path.join(t, "p"), "--dump-glyphs=" + go], capture_output=True, timeout=600)
        got = {}
        for pg in json.load(open(go, encoding="utf-8"))["pages"]:
            for g in pg["glyphs"]:
                got.setdefault(g["char"], short(g.get("font_family")))
        res[arm] = got
    print("cells: word face, or word/oxi where they differ (A = Arial slot, M = ＭＳ 明朝 slot)")
    table(word, res)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()
