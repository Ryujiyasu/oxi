# -*- coding: utf-8 -*-
"""Inside a table cell, how far may a line-end 。 hang past the cell's inner edge?

policies__06c631e8bc061f40 (jablindC50): the checklist «□ 運転手は、車両の点検
（ライト、ランプの動作確認等）をしている。» (15pt ＭＳ ゴシック, line 440 exact,
hangingChars 100) sits in a 9639tw cell (inner width 471.1pt after 5.4pt
margins). Word: ONE line, the 。 starting at x=540 and ending 7.5pt past the
inner right edge 547.5 (2pt past the cell border). Oxi's cell wrapper wraps
«る。» -> half the checklist doubles, +1 page. The body-paragraph twin
(`_pb_hangpunct_gen.py`) already matches Word; the CELL path does not.

Arms: the same paragraph inside the same 2-column borderless table, '□ ' + N
kanji + '。' for N = 26..31 (K family) and with the four 約物 (Y family,
N = 22..27). Readout: Word line count + 。 x; Oxi line count.
"""
import os, sys, json, subprocess, tempfile, zipfile
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
GDI = Path(os.environ.get("OXI_GDI_EXE") or "tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe")
OUT = Path('tests/fixtures/cellhang'); OUT.mkdir(parents=True, exist_ok=True)
sys.path.insert(0, 'tools/metrics')
import _pb_hangpunct_gen as H  # noqa: E402  (sheet, styles, settings, families)


TBL_HEAD = ('<w:tbl><w:tblPr><w:tblW w:w="9923" w:type="dxa"/><w:tblBorders><w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
            '<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
            '<w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/></w:tblBorders><w:tblLook w:val="04A0"/></w:tblPr><w:tblGrid><w:gridCol w:w="284"/><w:gridCol w:w="9639"/></w:tblGrid>')


def table(text):
    return (TBL_HEAD + '<w:tr><w:tc><w:tcPr><w:tcW w:w="284" w:type="dxa"/></w:tcPr><w:p/></w:tc>'
            '<w:tc><w:tcPr><w:tcW w:w="9639" w:type="dxa"/></w:tcPr>' + H.para(text) + '</w:tc></w:tr></w:tbl>')


def document(x):
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {H.W}><w:body><w:p><w:r><w:t>A</w:t></w:r></w:p>{table(x)}<w:p><w:r><w:t>B</w:t></w:r></w:p>{H.SECT}</w:body></w:document>'


arms = {}
for n in range(26, 32):
    arms[f'K{n}'] = H.fam_k(n)
for n in range(22, 28):
    arms[f'Y{n}'] = H.fam_y(n)

if __name__ == '__main__':
    import win32com.client
    app = win32com.client.DispatchEx("Word.Application")
    try:
        app.Visible = False
    except Exception:
        pass
    try:
        for name, x in arms.items():
            at = OUT / f"{name}.docx"
            with zipfile.ZipFile(at, "w", zipfile.ZIP_DEFLATED) as z:
                z.writestr("[Content_Types].xml", H.CT); z.writestr("_rels/.rels", H.ROOT_RELS)
                z.writestr("word/_rels/document.xml.rels", H.DOC_RELS); z.writestr("word/styles.xml", H.STYLES)
                z.writestr("word/settings.xml", H.SETTINGS); z.writestr("word/document.xml", document(x))
            d = app.Documents.Open(str(at.resolve()), False, True)
            try:
                c = d.Tables(1).Cell(1, 2).Range
                s, e = c.Start, c.End
                ys = sorted({round(d.Range(k, k).Information(6), 2) for k in range(s, e - 1)})
                lx = round(d.Range(e - 2, e - 2).Information(5), 2)
                px = round(d.Range(e - 3, e - 3).Information(5), 2)
            finally:
                d.Close(False)
            with tempfile.TemporaryDirectory() as t:
                dump = Path(t) / 'l.json'
                subprocess.run([str(GDI), str(at), str(Path(t) / 'p'), "96", f"--dump-layout={dump}"], capture_output=True)
                dd = json.loads(dump.read_text(encoding='utf-8')) if dump.exists() else {'pages': [{'elements': []}]}
            oys = sorted({round(el['y'], 2) for el in dd['pages'][0]['elements'] if el.get('type') == 'text' and el.get('text', '').strip() and el['text'] not in ('A', 'B')})
            print(f"{name}: chars={len(x)} WORD lines {len(ys)} prev_x {px} last_x {lx} (cell inner right 547.5, border 552.9) | OXI lines {len(oys)}")
    finally:
        app.Quit()
