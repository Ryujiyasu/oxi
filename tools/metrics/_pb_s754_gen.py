"""S754 single-column row-split threshold derivation (report_technical__0061c884 §8).

The p5 secondary root: technical__0061c884 table 5 / row 3 = a no-TYPE-docGrid
(linePitch=360, no type) single-cell auto row, 5 single-line paragraphs, 41.1pt
remaining when the row starts. Word SPLITS it (keeps 2 lines on the prev page);
Oxi's S754 single-column tier requires 58.0pt (the unwrap_or fallback for
table_grid_pitch=None) so it WHOLE-PUSHES the row -> +1.

Question: is Word's split<->whole-push threshold a FIXED value or N x line-pitch?

Probe: Latin font, no-type docGrid linePitch=360, compat14, K filler single-line
paras + a 1-col 1-cell table with 5 single-line cell paras. Sweep the filler count
(moves the table start toward the bottom -> shrinks the remaining space R) at three
font sizes (10/12/14pt). Per variant measure: table-start y, page_bottom, and how
many of the 5 cell paras land on p1 (0 = whole-push, >=1 = split). R = page_bottom -
y_table_start. The R at the split<->whole-push transition = the threshold R*; compare
R* across the three pitches.

Usage:
  python _pb_s754_gen.py gen
  python _pb_s754_gen.py measure
"""
import os, sys, json

OUTDIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                      "pipeline_data", "_pb_s754")

FONT = "Times New Roman"
# short line (1 line each in body AND in the ~4000tw cell): ~30 chars
FILL_TXT = "The main objectives of the section"
CELL_TXT = "Establish close collaboration among"


def rpr(sz):
    return (f'<w:rFonts w:ascii="{FONT}" w:eastAsia="{FONT}" w:hAnsi="{FONT}" '
            f'w:cs="{FONT}"/><w:sz w:val="{sz}"/>')


def para(txt, sz, jc="left"):
    r = rpr(sz)
    return (f'<w:p><w:pPr><w:jc w:val="{jc}"/><w:rPr>{r}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">{txt}</w:t></w:r></w:p>')


def table5(sz):
    """1-col 1-cell auto row holding 5 single-line paragraphs, no trHeight/cantSplit."""
    cellp = "".join(para(f"{CELL_TXT} {i+1}", sz) for i in range(5))
    return (
        '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>'
        '<w:tblBorders>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '</w:tblBorders></w:tblPr>'
        '<w:tblGrid><w:gridCol w:w="8000"/></w:tblGrid>'
        f'<w:tr><w:tc><w:tcPr><w:tcW w:w="8000" w:type="dxa"/></w:tcPr>{cellp}</w:tc></w:tr>'
        '</w:tbl>')


def build(sz, fill):
    body = "".join(para(f"{FILL_TXT} {i+1}", sz) for i in range(fill))
    body += table5(sz)
    body += para("tail marker paragraph", sz)   # a body para AFTER the table
    # no-TYPE docGrid (linePitch present, no w:type) — matches the target
    grid = '<w:docGrid w:linePitch="360"/>'
    mar = ('<w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418" '
           'w:header="851" w:footer="992" w:gutter="0"/>')
    sect = f'<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>{mar}{grid}</w:sectPr>'
    return body + sect


def docxml(body):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'<w:body>{body}</w:body></w:document>')


def styles_xml(sz):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:docDefaults><w:rPrDefault><w:rPr>'
            f'<w:rFonts w:ascii="{FONT}" w:eastAsia="{FONT}" w:hAnsi="{FONT}" w:cs="{FONT}"/>'
            f'<w:sz w:val="{sz}"/></w:rPr></w:rPrDefault></w:docDefaults>'
            '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
            '</w:styles>')


def settings_xml():
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:compat><w:compatSetting w:name="compatibilityMode" '
            'w:uri="http://schemas.microsoft.com/office/word" w:val="14"/></w:compat>'
            '</w:settings>')


def write_docx(path, body, sz):
    import zipfile
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
          '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
          '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
          '</Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
            '</Relationships>')
    drels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
             '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
             '</Relationships>')
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/_rels/document.xml.rels", drels)
        z.writestr("word/document.xml", docxml(body))
        z.writestr("word/styles.xml", styles_xml(sz))
        z.writestr("word/settings.xml", settings_xml())


# 3 pitches; filler count sweep lands the table near the bottom then over it.
# For sz20/24/28 (10/12/14pt) the body single-line pitch ~ 12.5/14.4/16.4pt.
# Enough filler to put the table start in ~[y 600, 730], stepping R by ~1 line.
CASES = []
for sz, frange in ((20, range(40, 56)), (24, range(35, 49)), (28, range(31, 43))):
    for fill in frange:
        CASES.append((sz, fill))


def name(sz, fill):
    return f"s754_sz{sz}_f{fill}.docx"


def gen():
    os.makedirs(OUTDIR, exist_ok=True)
    for sz, fill in CASES:
        write_docx(os.path.join(OUTDIR, name(sz, fill)), build(sz, fill), sz)
    print(f"wrote {len(CASES)} docs to {OUTDIR}")


def measure():
    import win32com.client
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    PAGEH = 841.9
    results = []
    try:
        for sz, fill in CASES:
            path = os.path.abspath(os.path.join(OUTDIR, name(sz, fill)))
            d = word.Documents.Open(path, ReadOnly=True)
            try:
                d.Repaginate()
                # locate the 5 cell paragraphs: they are the ones inside the table.
                # Paragraphs 1..fill = filler (body). Next 5 = cell paras. Then tail.
                # Use Information(12) wdWithInTable to be robust.
                cell_pages = []
                y_tbl_start = None
                n = d.Paragraphs.Count
                for i in range(1, n + 1):
                    rng = d.Paragraphs(i).Range
                    st = d.Range(rng.Start, rng.Start)
                    in_tbl = bool(st.Information(12))  # wdWithInTable
                    if in_tbl:
                        pg = int(st.Information(3))
                        if y_tbl_start is None:
                            y_tbl_start = float(st.Information(6))
                        cell_pages.append(pg)
                        if len(cell_pages) >= 5:
                            break
                split_on_p1 = sum(1 for p in cell_pages if p == 1)
                page_bottom = PAGEH - (1418 / 20.0)  # bottom margin 1418tw
                R = (page_bottom - y_tbl_start) if y_tbl_start else None
                rec = {"sz": sz, "fill": fill, "cell_pages": cell_pages,
                       "split_on_p1": split_on_p1, "y_tbl_start": y_tbl_start,
                       "page_bottom": round(page_bottom, 2),
                       "R": round(R, 2) if R is not None else None}
                results.append(rec)
                print(f"sz{sz} f{fill}: cellpg={cell_pages} p1keep={split_on_p1} "
                      f"ytbl={y_tbl_start} R={rec['R']}")
            finally:
                d.Close(False)
    finally:
        word.Quit()
    out = os.path.join(OUTDIR, "_results.json")
    json.dump(results, open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(f"-> {out}")
    # summarize the transition per pitch
    print("\n=== split<->whole-push transition per pitch ===")
    for sz in (20, 24, 28):
        rows = [r for r in results if r["sz"] == sz and r["R"] is not None]
        rows.sort(key=lambda r: -r["R"])  # large R (split) first
        # find the R where p1keep drops to 0
        prev = None
        for r in rows:
            if r["split_on_p1"] == 0 and prev is not None and prev["split_on_p1"] > 0:
                print(f"  sz{sz} ({sz/2:.0f}pt): split down to R={prev['R']} "
                      f"(keep {prev['split_on_p1']}), whole-push at R={r['R']}")
                break
            prev = r
        else:
            span = f"R {rows[-1]['R']}..{rows[0]['R']}" if rows else "no data"
            kmax = max((r['split_on_p1'] for r in rows), default=0)
            print(f"  sz{sz}: no clean transition in {span} (max keep {kmax})")


if __name__ == "__main__":
    if sys.argv[1:] == ["gen"]:
        gen()
    elif sys.argv[1:] == ["measure"]:
        measure()
    else:
        print(__doc__)
