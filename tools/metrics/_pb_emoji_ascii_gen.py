# -*- coding: utf-8 -*-
"""Does Word draw from an ascii face that is a COLOUR emoji font?

ja/forms__02281020: a ✔ run with ascii/hAnsi = Segoe UI Emoji and eastAsia =
BIZ UD明朝 Medium; Word drew it in Segoe UI Symbol.  Word's slot table (layout/
word_slot.rs) sends an unhinted ✔ to the ascii face when that face has the
glyph -- Segoe UI Emoji does -- so either Word skips colour fonts or the table
is wrong for this face.

Arms <ascii>_<hint>: ascii = Segoe UI Emoji / Segoe UI Symbol / Arial, eastAsia
= ＭＳ 明朝, hint none / eastAsia; characters ✔ ★ ☐ ☑ ① ⇒ ◆ ♪ ☎ ✉ ❤ and the
letter A.  Read: the embedded program per character (Word PDF; Oxi glyphs).

    python tools/metrics/_pb_emoji_ascii_gen.py gen
    python tools/metrics/_pb_emoji_ascii_gen.py word
    python tools/metrics/_pb_emoji_ascii_gen.py oxi <renderer.exe>
"""
import json, os, shutil, subprocess, sys, tempfile, zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_txbxfit_gen as T  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_emoji_ascii"
CHARS = ("A", "✔", "★", "☐", "☑", "①", "⇒", "◆", "♪", "☎", "✉", "❤")
FACES = {"emoji": "Segoe UI Emoji", "symbol": "Segoe UI Symbol", "arial": "Arial"}
ARMS = {f"{f}_{h}": (f, h) for f in FACES for h in ("none", "ea")}


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    for arm, (f, h) in ARMS.items():
        face = FACES[f]
        hint = ' w:hint="eastAsia"' if h == "ea" else ""
        rpr = f'<w:rPr><w:rFonts w:ascii="{face}" w:hAnsi="{face}" w:eastAsia="ＭＳ 明朝" w:cs="{face}"{hint}/><w:sz w:val="24"/></w:rPr>'
        body = "".join(f'<w:p><w:r><w:t xml:space="preserve">x </w:t></w:r><w:r>{rpr}<w:t>{c}</w:t></w:r></w:p>' for c in CHARS)
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + T.NS + '><w:body>' + body +
               '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>')
        styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + T.NS + '><w:docDefaults><w:rPrDefault><w:rPr>'
                  '<w:rFonts w:ascii="Century" w:hAnsi="Century" w:eastAsia="ＭＳ 明朝" w:cs="Times New Roman"/><w:sz w:val="21"/><w:lang w:val="en-US" w:eastAsia="ja-JP"/>'
                  '</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>')
        with zipfile.ZipFile(OUT / f"{arm}.docx", "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", T.CT); z.writestr("_rels/.rels", T.RELS)
            z.writestr("word/_rels/document.xml.rels", T.DRELS); z.writestr("word/styles.xml", styles); z.writestr("word/document.xml", doc)
    print("ok", len(ARMS))


def short(f):
    f = (f or "?").lstrip("=")
    return {"Segoe UI Emoji": "Emoji", "Segoe UI Symbol": "Symbol", "Arial": "Arial", "MS Mincho": "Mincho"}.get(f, f[:12])


def word():
    import win32com.client
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
            if c != "x":
                got.setdefault(c, short(face))
        res[arm] = got
    (OUT / "word.json").write_text(json.dumps(res, ensure_ascii=False), encoding="utf-8")
    for arm in ARMS:
        print(f"{arm:12s}", " ".join(f"{c}:{res[arm].get(c, '-')}" for c in CHARS))


def oxi(exe):
    wd = json.loads((OUT / "word.json").read_text(encoding="utf-8"))
    for arm in ARMS:
        t = tempfile.mkdtemp(); go = os.path.join(t, "g.json")
        subprocess.run([os.path.abspath(exe), str(OUT / f"{arm}.docx"), os.path.join(t, "p"), "--dump-glyphs=" + go], capture_output=True, timeout=600, env=os.environ)
        got = {}
        for pg in json.load(open(go, encoding="utf-8"))["pages"]:
            for g in pg["glyphs"]:
                if g["char"] != "x":
                    got.setdefault(g["char"], short(g.get("font_family")))
        print(f"{arm:12s}", " ".join(f"{c}:{wd[arm].get(c, '-')}" + ("" if got.get(c) == wd[arm].get(c) else f"/{got.get(c, '-')}") for c in CHARS))


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()
