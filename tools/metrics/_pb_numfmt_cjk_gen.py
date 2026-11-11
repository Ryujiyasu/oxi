# -*- coding: utf-8 -*-
"""Which strings does Word draw for the CJK counting / ideograph number formats?

Oxi's format_number (crates/oxidocs-core/src/parser/numbering.rs) falls back to
decimal for every format below, so zh dev forms__00243d36's chineseCounting
headings read «1、» where Word prints «一、».  Corpus census (numbering.xml
definitions): ideographTraditional 24 docs, taiwaneseCountingThousand 15,
chineseCounting 6, ideographLegalTraditional 4, japaneseCounting 3,
decimalEnclosedParen 3, cardinalText / chineseCountingThousand /
decimalEnclosedCircleChinese 1.  The tables are read out of Word's PDF, item by
item, at the numbers where counting systems differ (10, 11, 20, 100, 101, 110,
1000, 10000 ...).

    python _pb_numfmt_cjk_gen.py gen
    python _pb_numfmt_cjk_gen.py word          # Word truth -> word.json
    python _pb_numfmt_cjk_gen.py oxi EXE       # Oxi, compared item by item
"""
import json
import os
import re
import subprocess
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_numfmt_cjk")
sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, DRELS, NS, RELS  # noqa: E402

FORMATS = [
    "ideographTraditional", "ideographZodiac", "taiwaneseCountingThousand", "taiwaneseCounting",
    "taiwaneseDigital", "chineseCounting", "chineseCountingThousand", "chineseLegalSimplified",
    "ideographLegalTraditional", "ideographDigital", "japaneseCounting", "japaneseDigitalTenThousand",
    "japaneseLegal", "decimalEnclosedParen", "decimalEnclosedFullstop", "decimalEnclosedCircleChinese",
    "ideographEnclosedCircle", "cardinalText", "ordinalText", "ordinal",
]
# (start, count): each run is its own list so the numbers are exact
RUNS = [(1, 25), (99, 4), (109, 3), (119, 3), (999, 3), (1010, 2), (9999, 3), (10010, 1), (20001, 1)]
FACE = "SimSun"


def numbers():
    return [s + k for s, n in RUNS for k in range(n)]


def gen():
    os.makedirs(OUT, exist_ok=True)
    body, num, absn, nid = [], [], 0, 0
    rpr = '<w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s"/><w:sz w:val="21"/></w:rPr>' % (FACE, FACE, FACE)
    for fi, fmt in enumerate(FORMATS):
        body.append('<w:p><w:pPr>%s%s</w:pPr><w:r>%s<w:t>F%02dSTART</w:t></w:r></w:p>'
                    % ("<w:pageBreakBefore/>" if fi else "", rpr, rpr, fi))
        for start, cnt in RUNS:
            num.append('<w:abstractNum w:abstractNumId="%d"><w:multiLevelType w:val="singleLevel"/>'
                       '<w:lvl w:ilvl="0"><w:start w:val="%d"/><w:numFmt w:val="%s"/>'
                       '<w:lvlText w:val="[%%1]"/><w:lvlJc w:val="left"/>'
                       '<w:pPr><w:ind w:left="0" w:firstLine="0"/></w:pPr></w:lvl></w:abstractNum>'
                       % (absn, start, fmt))
            nid += 1
            for _ in range(cnt):
                body.append('<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="%d"/></w:numPr>%s</w:pPr>'
                            '<w:r>%s<w:t>X</w:t></w:r></w:p>' % (nid, rpr, rpr))
            absn += 1
    nums = "".join('<w:num w:numId="%d"><w:abstractNumId w:val="%d"/></w:num>' % (i + 1, i) for i in range(absn))
    numbering = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering ' + NS + ">"
                 + "".join(num) + nums + "</w:numbering>")
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS + "><w:body>"
           + "".join(body) + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
           '<w:pgMar w:top="1000" w:right="1000" w:bottom="1000" w:left="1000" w:header="500" w:footer="500" w:gutter="0"/>'
           '<w:cols w:num="3" w:space="200"/></w:sectPr></w:body></w:document>')
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + ">"
              '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="%s" w:eastAsia="%s" w:hAnsi="%s"/>'
              '<w:sz w:val="21"/><w:lang w:val="en-US" w:eastAsia="zh-CN"/></w:rPr></w:rPrDefault>'
              '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
              '</w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/></w:style>'
              "</w:styles>" % (FACE, FACE, FACE))
    ct = CT.replace("</Types>", '<Override PartName="/word/numbering.xml" ContentType="application/'
                    'vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>')
    drels = DRELS.replace("</Relationships>", '<Relationship Id="rIdNum" Type="http://schemas.openxmlformats.org/'
                          'officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>')
    p = os.path.join(OUT, "numfmt_cjk.docx")
    with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", drels)
        z.writestr("word/styles.xml", styles)
        z.writestr("word/numbering.xml", numbering)
        z.writestr("word/document.xml", doc)
    print("wrote", p, len(FORMATS), "formats x", len(numbers()), "numbers")


def collect(text):
    """{format: [marker, ...]} from the document text in reading order."""
    per, cur = {}, None
    for line in text.splitlines():
        m = re.search(r"F(\d\d)START", line)
        if m:
            cur = FORMATS[int(m.group(1))]
            per[cur] = []
            continue
        m = re.match(r"\s*\[(.*?)\]\s*X\s*$", line)
        if cur and m:
            per[cur].append(m.group(1))
    return per


def word():
    """Each paragraph's list string straight from Word (ListFormat.ListString):
    the PDF text of 20 formats x 45 markers splits unevenly into lines."""
    import win32com.client as w
    src = os.path.join(OUT, "numfmt_cjk.docx")
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    lines = []
    try:
        d = app.Documents.Open(src, ReadOnly=True)
        for p in d.Paragraphs:
            t = p.Range.Text.strip()
            ls = p.Range.ListFormat.ListString
            lines.append(t if not ls else "[%s]X" % ls[1:-1] if ls.startswith("[") else "?" + ls)
        d.Close(False)
    finally:
        app.Quit()
    per = collect("\n".join(lines))
    json.dump(per, open(os.path.join(OUT, "word.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    for f in FORMATS:
        print(f, list(zip(numbers(), per.get(f, []))))


def oxi(exe):
    src = os.path.join(OUT, "numfmt_cjk.docx")
    dump = os.path.join(OUT, "oxi.json")
    subprocess.run([exe, src, os.path.join(OUT, "oxi"), "--dump-layout=" + dump], capture_output=True)
    lines = []
    for pg in json.load(open(dump, encoding="utf-8"))["pages"]:
        rows = {}
        for e in pg["elements"]:
            if e["type"] == "text":
                # three columns of ~190pt: the marker and its «X» share one
                col = int(e["x"] // (pg["width"] / 3))
                rows.setdefault((col, round(e["y"], 1)), []).append((e["x"], e.get("text") or ""))
        for _, v in sorted(rows.items()):
            lines.append("".join(t for _, t in sorted(v)))
    got = collect("\n".join(lines))
    want = json.load(open(os.path.join(OUT, "word.json"), encoding="utf-8"))
    bad = 0
    for f in FORMATS:
        for n, wv, ov in zip(numbers(), want.get(f, []), got.get(f, []) + [None] * 99):
            if wv != ov:
                bad += 1
                print(f"{f} {n}: word «{wv}» oxi «{ov}»")
    print("mismatches", bad)


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "gen"
    if cmd == "word":
        word()
    elif cmd == "oxi":
        oxi(sys.argv[2])
    else:
        gen()
