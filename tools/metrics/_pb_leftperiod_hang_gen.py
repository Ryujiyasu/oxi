# -*- coding: utf-8 -*-
"""Does a trailing period hang past the right margin on a LEFT-aligned line?

en/technical__0186a34e: «…from your Organisation Administrator.» (Calibri 11,
compat 15, jc left by default) ends at 540.8 in Word's PDF against a 538.8 right
edge — the period hangs; Oxi wraps «Administrator.» (S809/S953 hang only on
justified lines).

Arms: <jc>_<tail>: jc = left / both / right / center; tail = '.' ',' ')' 'x' ';'.
Each paragraph = the same 11pt Calibri sentence ending in the tail, with a right
indent swept 0..120tw in 5tw steps so the last word straddles the edge. Read the
largest right indent at which the last word is still on line 1 (COM Info(6))
and compare the arms: hang = the '.' arm's flip is one period-width later.

    python tools/metrics/_pb_leftperiod_hang_gen.py gen
    python tools/metrics/_pb_leftperiod_hang_gen.py word
"""
import json, os, shutil, sys, tempfile, zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_txbxfit_gen as T  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_leftperiod_hang"
LEAD = "If you do not have access to it you will need to request access from your Organisation"
TAILS = {"dot": "Administrator.", "comma": "Administrator,", "paren": "Administrator)", "x": "Administratorx", "semi": "Administrator;"}
JCS = ("left", "both", "right", "center")
RIGHTS = list(range(300, 905, 5))  # ~454pt of text in a 482pt measure
RPR = '<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr>'
SETTINGS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + T.NS + '><w:compat>'
            '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>'
            '</w:compat></w:settings>')


def doc_xml(jc, tail):
    paras = "".join(
        f'<w:p><w:pPr><w:spacing w:before="0" w:after="0"/><w:ind w:right="{r}"/><w:jc w:val="{jc}"/>{RPR}</w:pPr>'
        f'<w:r>{RPR}<w:t xml:space="preserve">{LEAD} {TAILS[tail]}</w:t></w:r></w:p>' for r in RIGHTS)
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + T.NS + "><w:body>" + paras +
            '<w:sectPr><w:pgSz w:w="11910" w:h="16840"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134"'
            ' w:left="1134" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    ct = T.CT.replace("</Types>", '<Override PartName="/word/settings.xml" ContentType="application/vnd.'
                      'openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
    drels = T.DRELS.replace("</Relationships>", '<Relationship Id="rIdSet" Type="http://schemas.openxmlformats.org/'
                            'officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
    styles = T.STYLES.replace('w:ascii="MS Gothic" w:hAnsi="MS Gothic"', 'w:ascii="Calibri" w:hAnsi="Calibri"')
    for jc in JCS:
        for tail in TAILS:
            with zipfile.ZipFile(OUT / f"{jc}_{tail}.docx", "w", zipfile.ZIP_DEFLATED) as z:
                z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", T.RELS)
                z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/styles.xml", styles)
                z.writestr("word/settings.xml", SETTINGS); z.writestr("word/document.xml", doc_xml(jc, tail))
    print("ok", len(JCS) * len(TAILS))


def word():
    import win32com.client
    w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
    res = {}
    try:
        for jc in JCS:
            for tail in TAILS:
                arm = f"{jc}_{tail}"
                t = os.path.join(tempfile.mkdtemp(), arm + ".docx"); shutil.copy(OUT / f"{arm}.docx", t)
                d = w.Documents.Open(t, ReadOnly=True, AddToRecentFiles=False)
                fit = []
                for i, p in enumerate(d.Paragraphs):
                    r = p.Range
                    k = r.Text.find("Administrator")
                    s = d.Range(r.Start, r.Start); o = d.Range(r.Start + k, r.Start + k)
                    if abs(s.Information(6) - o.Information(6)) < 0.1:
                        fit.append(RIGHTS[i])
                d.Close(0)
                res[arm] = max(fit) if fit else None
                print(f"{arm:14s} last fitting right indent (tw): {res[arm]}", flush=True)
    finally:
        w.Quit()
    (OUT / "word.json").write_text(json.dumps(res), encoding="utf-8")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word}[sys.argv[1]]()
