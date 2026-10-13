# -*- coding: utf-8 -*-
"""Where does Word put an n-ary operator's limits (limLoc undOvr vs subSup) in
INLINE math (m:oMath in a text paragraph) versus DISPLAY math (m:oMathPara)?

blind-G ja/educational__20d9968b: an inline «∑_{k=0}^{n} k r^{k-1}» with
limLoc=undOvr. Word's PDF puts «n» and «k=0» at x 124.1 = right after the ∑
(x 116.7, advance 7.4) as super/subscripts, baselines 4.2 above / 2.6 below
the text baseline; Oxi stacks them over / under the operator (8 above / 8
below). Arms: one docx each, Cambria Math 10.5, the same expression:
    inline_undovr, inline_subsup, display_undovr, display_subsup,
    inline_undovr_setting (settings mathPr naryLim=undOvr), inline_nolim (no limLoc)
Read: x of «n» and «k» (limits) relative to the ∑'s x, and their baselines
relative to the base «k» baseline.

    python tools/metrics/_pb_narylim_gen.py gen
    python tools/metrics/_pb_narylim_gen.py word
    python tools/metrics/_pb_narylim_gen.py oxi <renderer.exe>
"""
import json, os, shutil, sys, tempfile, zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_txbxfit_gen as T  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_narylim"
ARMS = {"inline_undovr": (False, "undOvr", None), "inline_subsup": (False, "subSup", None), "display_undovr": (True, "undOvr", None),
        "display_subsup": (True, "subSup", None), "inline_undovr_setting": (False, "undOvr", "undOvr"), "inline_nolim": (False, None, None),
        "display_nolim": (True, None, None), "display_nolim_subsup_setting": (True, None, "subSup")}
NSM = ''  # T.NS already declares xmlns:m
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + T.NS + '><w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝" w:cs="Times New Roman"/><w:kern w:val="2"/><w:sz w:val="21"/>'
          '<w:lang w:val="en-US" w:eastAsia="ja-JP"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr></w:style>'
          '</w:styles>')
RPR = '<w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr>'


def mrun(t):
    return f'<m:r><m:rPr><m:sty m:val="p"/></m:rPr>{RPR}<m:t>{t}</m:t></m:r>'


def omath(limloc):
    ll = f'<m:limLoc m:val="{limloc}"/>' if limloc else ""
    return ('<m:oMath><m:nary><m:naryPr><m:chr m:val="∑"/>' + ll + f'<m:ctrlPr>{RPR}</m:ctrlPr></m:naryPr>'
            f'<m:sub>{mrun("k=0")}</m:sub><m:sup>{mrun("n")}</m:sup><m:e>{mrun("k")}'
            f'<m:sSup><m:sSupPr><m:ctrlPr>{RPR}</m:ctrlPr></m:sSupPr><m:e>{mrun("r")}</m:e><m:sup>{mrun("k-1")}</m:sup></m:sSup></m:e></m:nary></m:oMath>')


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    for arm, (display, limloc, setting) in ARMS.items():
        m = omath(limloc)
        if display:
            body = f'<w:p><w:r><w:t xml:space="preserve">Ｑ式 </w:t></w:r></w:p><w:p><m:oMathPara>{m}</m:oMathPara></w:p><w:p><w:r><w:t>Ｒ次の段落</w:t></w:r></w:p>'
        else:
            body = f'<w:p><w:r><w:t xml:space="preserve">Ｑ式 </w:t></w:r>{m}<w:r><w:t xml:space="preserve"> を考える</w:t></w:r></w:p><w:p><w:r><w:t>Ｒ次の段落</w:t></w:r></w:p>'
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + T.NS + NSM + "><w:body>" + body +
               '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1985" w:right="1701" w:bottom="1701"'
               ' w:left="1701" w:header="851" w:footer="992" w:gutter="0"/><w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr></w:body></w:document>')
        ct = T.CT.replace("</Types>", '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
        drels = T.DRELS.replace("</Relationships>", '<Relationship Id="rIdSet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
        mathpr = (f'<m:mathPr><m:mathFont m:val="Cambria Math"/><m:naryLim m:val="{setting}"/></m:mathPr>' if setting else "")
        settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + T.NS + NSM + '><w:defaultTabStop w:val="840"/>'
                    '<w:characterSpacingControl w:val="compressPunctuation"/><w:compat><w:useFELayout/>'
                    '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>' + mathpr + '</w:settings>')
        with zipfile.ZipFile(OUT / f"{arm}.docx", "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", T.RELS)
            z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/styles.xml", STYLES)
            z.writestr("word/settings.xml", settings); z.writestr("word/document.xml", doc)
    print("ok", len(ARMS))


def word_items(pdf):
    import fitz
    page = fitz.open(pdf)[0]; items = []
    for b in page.get_text("rawdict")["blocks"]:
        for l in b.get("lines", []):
            for s in l["spans"]:
                for c in s["chars"]:
                    if c["c"].strip():
                        items.append((round(c["origin"][0], 2), round(c["origin"][1], 2), c["c"], round(s["size"], 1)))
    return items


def oxi_items(exe, docx):
    import subprocess
    t = tempfile.mkdtemp(); out = os.path.join(t, "l.json")
    subprocess.run([exe, str(docx), os.path.join(t, "p"), "--dump-layout=" + out], capture_output=True, timeout=600)
    pg = json.load(open(out, encoding="utf-8"))["pages"][0]; items = []
    for e in pg["elements"]:
        if e["type"] == "text" and (e.get("text") or "").strip():
            base = e["y"] + (e.get("baseline_offset") or 0.0)
            for i, ch in enumerate(e["text"]):
                if ch.strip():
                    items.append((round(e["x"], 2), round(base, 2), ch, e.get("font_size")))
    return items


def _read(items):
    # the ∑, the limits n / k=0 (small size), the base k (full size)
    sig = [it for it in items if it[2] == "∑"]
    if not sig:
        return None
    sx, sy = sig[0][0], sig[0][1]
    small = [it for it in items if it[3] and it[3] < 9.0]
    n = [it for it in small if it[2] == "n"]
    # the lower limit's k is the small k whose next small glyph is '='
    k0 = [small[i] for i in range(len(small) - 1) if small[i][2] == "k" and small[i + 1][2] == "="]
    base_k = [it for it in items if it[2] == "k" and it[3] and it[3] >= 9.0]
    by = base_k[0][1] if base_k else sy
    r = {"sum_x": sx, "n_dx": round(n[0][0] - sx, 2) if n else None, "n_dy": round(n[0][1] - by, 2) if n else None,
         "k0_dx": round(k0[0][0] - sx, 2) if k0 else None, "k0_dy": round(k0[0][1] - by, 2) if k0 else None,
         "basek_dx": round(base_k[0][0] - sx, 2) if base_k else None}
    return r


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
    res = {}
    for arm in ARMS:
        res[arm] = _read(word_items(OUT / f"{arm}.pdf"))
        print(f"{arm:22s} word {res[arm]}", flush=True)
    (OUT / "word.json").write_text(json.dumps(res), encoding="utf-8")


def oxi(exe):
    word = json.loads((OUT / "word.json").read_text(encoding="utf-8"))
    for arm in ARMS:
        o, wd = _read(oxi_items(os.path.abspath(exe), OUT / f"{arm}.docx")), word.get(arm)
        flag = "" if o and wd and all(o.get(k) is not None and wd.get(k) is not None and abs(o[k] - wd[k]) < 1.0 for k in ("n_dx", "n_dy", "k0_dx", "k0_dy")) else "  <-- differs"
        print(f"{arm:22s} oxi {o}  word {wd}{flag}", flush=True)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()
