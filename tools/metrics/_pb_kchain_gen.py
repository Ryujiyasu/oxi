# -*- coding: utf-8 -*-
"""Word's keepNext CHAIN rule — how far does a keep-together group extend, and
where does Word stop?

policies__00148f8d's TOC gives TOC2-TOC7 a bare <w:keepNext/>, so a Part
heading, its Division heading and the entries under it form a chain. At the
page-8 bottom Word pushed the WHOLE prefix (ending the page ~40pt early) while
Oxi kept filling: both pairwise checks passed
  [KN635] "Part 10"    this_h=21.8 next_h=24.0 rem=62.6 do_push=false
  [KN635] "Division 1" this_h=27.0 next_h=12.0 rem=40.5 do_push=false
and the overflow only materialised during real layout, whose natural page push
never reaches S802B (that back-pull lives inside `if do_push`).

A document where EVERY paragraph is keepNext obviously cannot push forever, so
the missing piece is Word's TERMINATION rule. This probe measures it.

Layout per case: K filler paragraphs (no keepNext), then a chain H1..HN (all
keepNext), then BODY (keepNext explicitly off, B visual lines). Every paragraph
is one line of the same height, so page-bottom room is controlled in whole
slots. `cal` measures the page capacity C empirically first, so no font bbox is
ever hand-computed.

  python _pb_kchain_gen.py cal      # capacity calibration (filler only)
  python _pb_kchain_gen.py gen      # build the matrix from the measured C
  python _pb_kchain_gen.py measure  # Word COM -> PDF
  python _pb_kchain_gen.py read     # per-label page/baseline -> CSV + verdict
"""
from __future__ import annotations

import csv
import glob
import json
import os
import re
import sys
import zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_kchain"
FAITHFUL = (REPO / "pipeline_data" / "docx_corpus" / "en" / "policies"
            / "00148f8df1b6ad04.docx")
CAL_FILE = OUT / "_capacity.json"

# matrix (report v4 §7.2); N=60 only needs the overflow controls
CHAIN_LENS = [2, 3, 5, 8]
BODY_LINES = [1, 3]
WIDOWS = [True, False]
STARTS = ["mid", "top"]
ROOM_DELTAS = [-1, 0, 1]          # room_slots relative to (N + B)

CT = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
<Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>
<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
</Types>"""

RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DOCRELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>
<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>"""

# A probe-only paragraph style pins the geometry: Arial 12, no spacing, single.
STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr>
<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Arial" w:cs="Arial"/>
<w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="en-US" w:eastAsia="en-US"/>
</w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr>
<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>
</w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/>
<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/></w:rPr>
</w:style>
<w:style w:type="paragraph" w:styleId="Probe"><w:name w:val="Probe"/><w:basedOn w:val="Normal"/>
<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/></w:rPr>
</w:style>
</w:styles>"""

SETTINGS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:compat><w:compatSetting w:name="compatibilityMode"
 w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>
</w:settings>"""

SECTPR = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
          '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"'
          ' w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>')

SPACING = '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'


def faithful_parts() -> dict:
    """theme + fontTable transplanted from the real document (S931 lesson:
    a stubbed theme is schema-invalid and changes Word's font resolution)."""
    z = zipfile.ZipFile(FAITHFUL)
    out = {}
    for name in ("word/theme/theme1.xml", "word/fontTable.xml"):
        if name in z.namelist():
            out[name] = z.read(name)
    return out


def para(text: str, keep_next=False, widow=None, extra_lines=0) -> str:
    ppr = "<w:pPr><w:pStyle w:val=\"Probe\"/>"
    ppr += "<w:keepNext/>" if keep_next else '<w:keepNext w:val="0"/>'
    if widow is not None:
        ppr += f'<w:widowControl w:val="{1 if widow else 0}"/>'
    ppr += SPACING + "</w:pPr>"
    # extra_lines uses explicit <w:br/> BETWEEN lines only (never trailing)
    runs = f'<w:r><w:t xml:space="preserve">{text}-L1</w:t>'
    for i in range(extra_lines):
        runs += f'<w:br/><w:t xml:space="preserve">{text}-L{i + 2}</w:t>'
    runs += "</w:r>"
    return f"<w:p>{ppr}{runs}</w:p>"


def page_break() -> str:
    return ('<w:p><w:pPr><w:pStyle w:val="Probe"/>' + SPACING +
            '</w:pPr><w:r><w:br w:type="page"/></w:r></w:p>')


def write_docx(cid: str, body: str) -> None:
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
           f'<w:body>{body}{SECTPR}</w:body></w:document>')
    with zipfile.ZipFile(OUT / f"{cid}.docx", "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DOCRELS)
        z.writestr("word/document.xml", doc)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/settings.xml", SETTINGS)
        for name, blob in faithful_parts().items():
            z.writestr(name, blob)


def cal() -> None:
    """Capacity calibration: filler-only documents, K = 40..56."""
    OUT.mkdir(parents=True, exist_ok=True)
    for k in range(40, 57):
        body = "".join(para(f"CAL{k:02d}-F{i + 1:03d}") for i in range(k))
        write_docx(f"cal_k{k}", body)
    print("cal: generated K=40..56; run `measure` then `capacity`")


def capacity() -> None:
    """Read the calibration PDFs and record C = max paragraphs on page 1."""
    import fitz
    best = None
    for k in range(40, 57):
        path = OUT / f"cal_k{k}.pdf"
        if not path.is_file():
            continue
        doc = fitz.open(path)
        first = doc[0].get_text()
        on_p1 = len(re.findall(rf"CAL{k:02d}-F\d+", first))
        print(f"  K={k}: pages={doc.page_count} on_p1={on_p1}")
        if doc.page_count == 1:
            best = max(best or 0, k)
        elif best is None or on_p1 > (best or 0):
            best = max(best or 0, on_p1)
    CAL_FILE.write_text(json.dumps({"capacity": best}), encoding="utf-8")
    print(f"capacity C = {best} single-line paragraphs per page")


def cases() -> list[dict]:
    out = []
    for n in CHAIN_LENS:
        for b in BODY_LINES:
            for widow in WIDOWS:
                for start in STARTS:
                    for d in ROOM_DELTAS:
                        room = n + b + d
                        if room < 1:
                            continue
                        if start == "top" and d != 0:
                            continue          # top-start only needs one room
                        out.append({"N": n, "B": b, "widow": widow,
                                    "start": start, "room": room})
    for b in (1, 3):                          # N=60 overflow controls
        for widow in WIDOWS:
            out.append({"N": 60, "B": b, "widow": widow,
                        "start": "mid", "room": 6})
    return out


def cid_of(c: dict) -> str:
    return (f"k{c['N']}b{c['B']}w{1 if c['widow'] else 0}"
            f"{c['start']}r{c['room']}")


def gen() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    if not CAL_FILE.is_file():
        print("run `cal` + `measure` + `capacity` first"); return
    cap = json.loads(CAL_FILE.read_text(encoding="utf-8"))["capacity"]
    n_written = 0
    for c in cases():
        cid = cid_of(c)
        if c["start"] == "top":
            body = "".join(para(f"{cid}-F{i + 1:03d}") for i in range(3))
            body += page_break()
        else:
            k = cap - c["room"]
            if k < 1:
                continue
            body = "".join(para(f"{cid}-F{i + 1:03d}") for i in range(k))
        body += "".join(para(f"{cid}-H{i + 1}", keep_next=True)
                        for i in range(c["N"]))
        body += para(f"{cid}-BODY", keep_next=False, widow=c["widow"],
                     extra_lines=c["B"] - 1)
        body += "".join(para(f"{cid}-A{i + 1}") for i in range(3))
        write_docx(cid, body)
        n_written += 1
    print(f"gen: {n_written} cases (capacity C={cap})")


def measure() -> None:
    import win32com.client as win32
    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    try:
        for path in sorted(glob.glob(str(OUT / "*.docx"))):
            pdf = path[:-5] + ".pdf"
            if os.path.exists(pdf):
                continue
            d = word.Documents.Open(os.path.abspath(path), ReadOnly=True)
            try:
                d.ExportAsFixedFormat(OutputFileName=os.path.abspath(pdf),
                                      ExportFormat=17)
            finally:
                d.Close(False)
            print("measured", os.path.basename(path), flush=True)
    finally:
        word.Quit()


def labels_of(pdf) -> dict:
    """label -> (page, baseline); rows merged by y within 0.25pt."""
    found = {}
    for pi, page in enumerate(pdf, 1):
        for b in page.get_text("dict")["blocks"]:
            for line in b.get("lines", []):
                y = line["spans"][0]["origin"][1]
                text = "".join(s["text"] for s in line["spans"])
                for m in re.finditer(r"[A-Za-z0-9]+-(?:H\d+|BODY-L\d+|F\d+|A\d+)", text):
                    found.setdefault(m.group(0), (pi, round(y, 2)))
    return found


def read() -> None:
    import fitz
    rows = []
    for path in sorted(glob.glob(str(OUT / "k*.pdf"))):
        cid = Path(path).stem
        m = re.match(r"k(\d+)b(\d+)w(\d)(mid|top)r(\d+)", cid)
        if not m:
            continue
        n, b, widow, start, room = (int(m.group(1)), int(m.group(2)),
                                    int(m.group(3)), m.group(4), int(m.group(5)))
        pdf = fitz.open(path)
        lab = labels_of(pdf)
        hs = [lab.get(f"{cid}-H{i + 1}") for i in range(n)]
        body = [lab.get(f"{cid}-BODY-L{i + 1}") for i in range(b)]
        h_pages = [x[0] if x else None for x in hs]
        b_pages = [x[0] if x else None for x in body]
        # the page the chain STARTED on (H1's would-be page = filler's last page)
        fillers = [v for k2, v in lab.items() if "-F" in k2]
        start_page = max((p for p, _ in fillers), default=1)
        # did that page end up with no body content after the chain moved?
        blank = start_page not in [p for p, _ in lab.values()]
        split_after = None
        for i in range(n - 1):
            if h_pages[i] is not None and h_pages[i + 1] is not None \
                    and h_pages[i] != h_pages[i + 1]:
                split_after = f"H{i + 1}"
                break
        if split_after is None and h_pages and b_pages \
                and h_pages[-1] is not None and b_pages[0] is not None \
                and h_pages[-1] != b_pages[0]:
            split_after = f"H{n}"
        verdict = ("whole-chain" if split_after is None and len(set(
            p for p in h_pages + b_pages if p)) == 1 else f"split@{split_after}")
        rows.append({"case": cid, "N": n, "B": b, "widow": widow,
                     "start": start, "room": room, "n_pages": pdf.page_count,
                     "h_pages": ",".join(str(p) for p in h_pages),
                     "body_pages": ",".join(str(p) for p in b_pages),
                     "split_after": split_after or "", "blank": blank,
                     "verdict": verdict})
        pdf.close()
    csv_path = OUT / "_result.csv"
    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()) if rows else ["case"])
        w.writeheader()
        w.writerows(rows)
    for r in rows:
        print(f"{r['case']:<18} N={r['N']:<2} B={r['B']} w={r['widow']} "
              f"{r['start']:<3} room={r['room']:<2} pages={r['n_pages']} "
              f"H={r['h_pages']:<14} BODY={r['body_pages']:<8} {r['verdict']}")
    print(f"\n{len(rows)} cases -> {csv_path}")


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "cal"
    {"cal": cal, "capacity": capacity, "gen": gen,
     "measure": measure, "read": read}[cmd]()
