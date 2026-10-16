# -*- coding: utf-8 -*-
"""Word's character -> rFonts slot table, block by block.

`_pb_slot_class_gen.py` showed Word's slot choice is a per-character table, not
a script guess (lang eastAsia ja / en makes no difference): with no hint the
general punctuation, ×÷, arrows, box drawing and geometric shapes come from the
ascii font while ※ ‥ ・ Roman numerals and fullwidth forms come from eastAsia;
with hint="eastAsia" most symbols move to eastAsia but © ® £ ¥ and letters stay.

Here every assigned code point of the symbol blocks a Japanese document meets is
put in a run whose ascii/hAnsi = Meiryo (wide coverage, so a glyph missing from
the slot font rarely hides the slot) and eastAsia = ＭＳ 明朝; Word's PDF tells,
per character, Meiryo (ascii slot) / MS Mincho (eastAsia slot) / anything else
(the slot font lacked it: a fallback).  Arms: ascii face Meiryo (wide coverage) / Arial (a Latin face: Oxi's slot
choice changed with the face's script) x hint none / eastAsia.

    python tools/metrics/_pb_slot_table_gen.py gen
    python tools/metrics/_pb_slot_table_gen.py word      # writes word.json
    python tools/metrics/_pb_slot_table_gen.py oxi <renderer.exe>
"""
import json, os, shutil, subprocess, sys, tempfile, unicodedata, zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_txbxfit_gen as T  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_slot_table"
BLOCKS = [(0x00A1, 0x00FF), (0x0100, 0x017F), (0x02B0, 0x02FF), (0x0370, 0x03FF), (0x0400, 0x045F),
          (0x2000, 0x206F), (0x2070, 0x209F), (0x20A0, 0x20C0), (0x2100, 0x214F), (0x2150, 0x218F),
          (0x2190, 0x21FF), (0x2200, 0x22FF), (0x2300, 0x23FF), (0x2460, 0x24FF), (0x2500, 0x257F),
          (0x2580, 0x259F), (0x25A0, 0x25FF), (0x2600, 0x26FF), (0x2700, 0x27BF), (0x3000, 0x303F),
          (0x3200, 0x32FF), (0x3300, 0x33FF), (0xFE30, 0xFE4F), (0xFF00, 0xFFEF)]
CPS = [c for a, b in BLOCKS for c in range(a, b + 1)
       if unicodedata.category(chr(c))[0] in "LNPS" and not unicodedata.category(chr(c)).startswith("Z")]
ARMS = {"mei_none": ("Meiryo", ""), "mei_ea": ("Meiryo", ' w:hint="eastAsia"'),
        "ari_none": ("Arial", ""), "ari_ea": ("Arial", ' w:hint="eastAsia"')}
EA_F = "ＭＳ 明朝"
PER_PARA = 40


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    for arm, (ASCII_F, hint) in ARMS.items():
        rpr = (f'<w:rPr><w:rFonts w:ascii="{ASCII_F}" w:hAnsi="{ASCII_F}" w:eastAsia="{EA_F}" w:cs="{ASCII_F}"{hint}/>'
               '<w:sz w:val="20"/><w:lang w:val="en-US" w:eastAsia="ja-JP"/></w:rPr>')
        paras = []
        for i in range(0, len(CPS), PER_PARA):
            chunk = CPS[i:i + PER_PARA]
            runs = "".join(f'<w:r>{rpr}<w:t xml:space="preserve">{chr(c)} </w:t></w:r>' for c in chunk)
            paras.append(f'<w:p>{runs}</w:p>')
        body = "".join(paras).replace("&", "&amp;").replace("<w:t xml:space=\"preserve\">&amp;", "<w:t xml:space=\"preserve\">&amp;")
        # escape the XML specials inside the text runs
        body = "".join(paras)
        for bad, good in (("<w:t xml:space=\"preserve\">& ", "<w:t xml:space=\"preserve\">&amp; "),
                          ("<w:t xml:space=\"preserve\">< ", "<w:t xml:space=\"preserve\">&lt; "),
                          ("<w:t xml:space=\"preserve\">> ", "<w:t xml:space=\"preserve\">&gt; ")):
            body = body.replace(bad, good)
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + T.NS + '><w:body>' + body +
               '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="708" w:footer="708" w:gutter="0"/>'
               '</w:sectPr></w:body></w:document>')
        styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + T.NS + '><w:docDefaults><w:rPrDefault><w:rPr>'
                  f'<w:rFonts w:ascii="{ASCII_F}" w:hAnsi="{ASCII_F}" w:eastAsia="{EA_F}" w:cs="{ASCII_F}"/><w:sz w:val="20"/><w:lang w:val="en-US" w:eastAsia="ja-JP"/>'
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
    print("ok", len(ARMS), "chars", len(CPS))


def cls(fam, ascii_f):
    f = (fam or "?").lstrip("=")
    if f.startswith(ascii_f) or (ascii_f == "Meiryo" and "メイリオ" in f):
        return "A"
    if "Mincho" in f or "明朝" in f:
        return "E"
    return "F:" + f[:16]


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
        for pi, c, x, y, face in F.word_chars(str(OUT / f"{arm}.pdf"), 99):
            got.setdefault("%04X" % ord(c), cls(face, ARMS[arm][0]))
        res[arm] = got
    (OUT / "word.json").write_text(json.dumps(res, ensure_ascii=False, indent=0), encoding="utf-8")
    summarize(res)


def summarize(res, oxi=None):
    for arm in ARMS:
        cnt = {}
        for v in res[arm].values():
            k = v if not v.startswith("F:") else "F"
            cnt[k] = cnt.get(k, 0) + 1
        print(arm, cnt, "of", len(CPS))
    if oxi:
        for arm in ARMS:
            diff = [(k, res[arm].get(k), oxi[arm].get(k)) for k in res[arm]
                    if res[arm].get(k) in ("A", "E") and oxi[arm].get(k) in ("A", "E") and res[arm][k] != oxi[arm][k]]
            print(f"{arm}: slot differs on {len(diff)} chars")
            runs, prev = [], None
            for k, a, b in sorted(diff):
                c = int(k, 16)
                if prev and c == prev[1] + 1 and (a, b) == prev[2]:
                    prev[1] = c
                else:
                    prev = [c, c, (a, b)]; runs.append(prev)
            for lo, hi, (a, b) in runs:
                print(f"   U+{lo:04X}-U+{hi:04X} {chr(lo)}..{chr(hi)}  word {a} oxi {b}")


def oxi(exe):
    word = json.loads((OUT / "word.json").read_text(encoding="utf-8"))
    res = {}
    for arm in ARMS:
        t = tempfile.mkdtemp(); go = os.path.join(t, "g.json")
        subprocess.run([os.path.abspath(exe), str(OUT / f"{arm}.docx"), os.path.join(t, "p"), "--dump-glyphs=" + go], capture_output=True, timeout=900)
        got = {}
        for pg in json.load(open(go, encoding="utf-8"))["pages"]:
            for g in pg["glyphs"]:
                got.setdefault("%04X" % ord(g["char"]), cls(g.get("font_family"), ARMS[arm][0]))
        res[arm] = got
    summarize(word, res)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()
