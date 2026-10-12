# -*- coding: utf-8 -*-
"""Where does Word put a ruby annotation's baseline when the line spacing is
EXACT, and does the face (Yu Mincho vs MS Mincho) or a centred cell change it?

blind-H ja/legal__0f67d319: 游明朝 12pt base «団体名» with rt «（ふりがな）»
(hps 12, hpsRaise 26) in a vAlign=center cell with line 440 exact: Word's
ruby baseline is 12.96 above the base baseline (= hpsRaise), Oxi's 15.53.

Arms: <font>_<line>_<where>: font = yumin / msmin; line = auto / ex440 / ex300;
where = body / cell (one-cell table, vAlign center, trHeight 497 atLeast).
Each arm: plain «前の行», the ruby line (base 24 half-pt, hps 12, raise 26),
plain «後の行».  Read the baselines of «前», the base «団», the ruby «（» and
«後» as offsets from «前».

    python tools/metrics/_pb_rubyexact_gen.py gen
    python tools/metrics/_pb_rubyexact_gen.py word
    python tools/metrics/_pb_rubyexact_gen.py oxi <renderer.exe>
"""
import json, os, shutil, sys, tempfile, zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_txbxfit_gen as T  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_rubyexact"
FONTS = {"yumin": "游明朝", "msmin": "ＭＳ 明朝"}
LINES = {"auto": "", "ex440": '<w:spacing w:line="440" w:lineRule="exact"/>', "ex300": '<w:spacing w:line="300" w:lineRule="exact"/>'}
WHERES = ("body", "cell")
ARMS = [f"{f}_{l}_{w}" for f in FONTS for l in LINES for w in WHERES]


def rpr(font, sz):
    return f'<w:rPr><w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}"/><w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr>'


def paras(font, line):
    ppr = f"<w:pPr>{LINES[line]}{rpr(font, 24)}</w:pPr>"
    ruby = (f'<w:r>{rpr(font, 24)}<w:ruby><w:rubyPr><w:rubyAlign w:val="distributeSpace"/><w:hps w:val="12"/>'
            f'<w:hpsRaise w:val="26"/><w:hpsBaseText w:val="24"/><w:lid w:val="ja-JP"/></w:rubyPr>'
            f'<w:rt><w:r>{rpr(font, 12)}<w:t>（ふりがな）</w:t></w:r></w:rt>'
            f'<w:rubyBase><w:r>{rpr(font, 24)}<w:t>団体名</w:t></w:r></w:rubyBase></w:ruby></w:r>')
    return (f'<w:p>{ppr}<w:r>{rpr(font, 24)}<w:t>前の行</w:t></w:r></w:p>'
            f'<w:p>{ppr}{ruby}</w:p>'
            f'<w:p>{ppr}<w:r>{rpr(font, 24)}<w:t>後の行</w:t></w:r></w:p>')


def doc_xml(font, line, where):
    inner = paras(font, line)
    if where == "cell":
        inner = ('<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/><w:tblBorders><w:top w:val="single" w:sz="4"/>'
                 '<w:bottom w:val="single" w:sz="4"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>'
                 '<w:tr><w:trPr><w:trHeight w:val="497"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/>'
                 '<w:vAlign w:val="center"/></w:tcPr>' + inner + '</w:tc></w:tr></w:tbl><w:p/>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + T.NS + "><w:body>" + inner +
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134"'
            ' w:left="1134" w:header="720" w:footer="720" w:gutter="0"/><w:docGrid w:type="lines" w:linePitch="360"/>'
            '</w:sectPr></w:body></w:document>')


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    ct = T.CT.replace("</Types>", '<Override PartName="/word/settings.xml" ContentType="application/vnd.'
                      'openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
    drels = T.DRELS.replace("</Relationships>", '<Relationship Id="rIdSet" Type="http://schemas.openxmlformats.org/'
                            'officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
    settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + T.NS + '><w:compat>'
                '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>'
                '</w:compat></w:settings>')
    for arm in ARMS:
        f, l, w = arm.split("_")
        with zipfile.ZipFile(OUT / f"{arm}.docx", "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", T.RELS)
            z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/styles.xml", T.STYLES)
            z.writestr("word/settings.xml", settings); z.writestr("word/document.xml", doc_xml(FONTS[f], l, w))
    print("ok", len(ARMS))


KEYS = ("前の行", "団体名", "（ふりがな）", "後の行")


def _offsets(lines):
    ys = {}
    for l in lines:
        for k in KEYS:
            if l["t"].startswith(k) and k not in ys:
                ys[k] = l["y"]
    if "前の行" not in ys:
        return None
    return {k: round(ys[k] - ys["前の行"], 2) for k in KEYS if k in ys}


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
        res[arm] = _offsets(L.word_lines(OUT / f"{arm}.pdf"))
        print(f"{arm:18s} word {res[arm]}", flush=True)
    (OUT / "word.json").write_text(json.dumps(res, ensure_ascii=False), encoding="utf-8")


def oxi(exe):
    import line_census as L
    word = json.loads((OUT / "word.json").read_text(encoding="utf-8"))
    for arm in ARMS:
        o = _offsets(L.oxi_lines(os.path.abspath(exe), OUT / f"{arm}.docx"))
        wd = word.get(arm) or {}
        flag = "" if o and all(abs(o.get(k, 99) - wd.get(k, -99)) < 0.6 for k in wd) else "  <-- differs"
        print(f"{arm:18s} oxi {o}  word {wd}{flag}", flush=True)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()
