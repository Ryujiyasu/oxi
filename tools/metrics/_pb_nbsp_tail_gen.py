# -*- coding: utf-8 -*-
"""Does a line-final NBSP (and the space after it) count toward the fit?

en/correspondence__00aee476 (TNR 15pt, compat 14): «…more than once.\\xa0 The only…»
Word wraps before «once.» although «once.» alone ends 6.2pt inside the line's
right edge; Oxi keeps it (and one fewer line shifts the page by 17.3pt).

One paragraph per arm: «L» + right indent R (swept) so that «once.» lands near the
margin; tails: sp = «once. The», nb = «once.\\xa0The», nbsp = «once.\\xa0 The».
Read whether «once.» is on line 1 (Word PDF) -> flip right-indent per tail.

    python tools/metrics/_pb_nbsp_tail_gen.py gen
    python tools/metrics/_pb_nbsp_tail_gen.py word
"""
import json, os, shutil, sys, tempfile, zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_txbxfit_gen as T  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_nbsp_tail"
LEAD = "own voters. There is no reason to split any county in Ohio more than"
TAILS = {"sp": " once. The only reason", "nb": " once. The only reason", "nbsp": " once.  The only reason"}
# right indent in twips: sweep 0..20pt in 0.25pt (5tw) steps
RIGHTS = list(range(-1000, 470, 5))


def doc_xml(tail):
    rpr = '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="30"/></w:rPr>'
    paras = "".join(
        f'<w:p><w:pPr><w:spacing w:before="0" w:after="0"/><w:ind w:right="{r}"/>{rpr}</w:pPr>'
        f'<w:r>{rpr}<w:t xml:space="preserve">R{r:+05d} {LEAD}{TAILS[tail]}</w:t></w:r></w:p>' for r in RIGHTS)
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + T.NS + "><w:body>" + paras +
            '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="720" w:right="1080" w:bottom="720"'
            ' w:left="1080" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')


SETTINGS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + T.NS + '><w:compat>'
            '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/>'
            '</w:compat></w:settings>')


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    ct = T.CT.replace("</Types>", '<Override PartName="/word/settings.xml" ContentType="application/vnd.'
                      'openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
    drels = T.DRELS.replace("</Relationships>", '<Relationship Id="rIdSet" Type="http://schemas.openxmlformats.org/'
                            'officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
    styles = T.STYLES.replace('w:ascii="MS Gothic" w:hAnsi="MS Gothic"', 'w:ascii="Times New Roman" w:hAnsi="Times New Roman"')
    for tail in TAILS:
        with zipfile.ZipFile(OUT / f"{tail}.docx", "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", T.RELS)
            z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/styles.xml", styles)
            z.writestr("word/settings.xml", SETTINGS); z.writestr("word/document.xml", doc_xml(tail))
    print("ok", list(TAILS))


def word():
    import win32com.client
    w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
    res = {}
    try:
        for tail in TAILS:
            t = os.path.join(tempfile.mkdtemp(), tail + ".docx"); shutil.copy(OUT / f"{tail}.docx", t)
            d = w.Documents.Open(t, ReadOnly=True, AddToRecentFiles=False)
            flips = []
            for i, p in enumerate(d.Paragraphs):
                r = p.Range
                k = r.Text.find("once.")
                s = d.Range(r.Start, r.Start); o = d.Range(r.Start + k, r.Start + k)
                same = abs(s.Information(6) - o.Information(6)) < 0.1
                flips.append((RIGHTS[i], same))
            d.Close(0)
            # the largest right indent at which «once.» still shares line 1
            fit = [r for r, same in flips if same]
            res[tail] = max(fit) if fit else None
            print(tail, "last fitting right indent (tw):", res[tail], flush=True)
    finally:
        w.Quit()
    (OUT / "word.json").write_text(json.dumps(res), encoding="utf-8")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word}[sys.argv[1]]()
