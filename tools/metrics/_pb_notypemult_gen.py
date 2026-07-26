"""§7 probe for REPORT_administrative__0010e437: pin the no-type docGrid Latin
auto-multiple-spacing page-bottom keep/push model.

administrative__0010e437 target (Calibri 12, line=257=1.0708x, no-type docGrid
linePitch=360) keeps a 3-line paragraph whose INK bottom (768.216) is within
content_bottom (769.90) but whose win/spacing box (line2_top + 15.69 = 771.9)
would push. Report §7: is this INK model general across L, and font-dependent?

Construction per (font, L, grid, widow): exact-height filler places a 3-wrap
target paragraph so its line-2 lands near the page bottom, then sweep the bottom
margin (=content_bottom) in ~1.2pt steps. Read from Word PDF which page each of
the target's 3 lines is on -> the flip content_bottom = the effective box Word
uses. Compare to ink (line2 bbox bottom) vs win/spacing (line2_top + pitch).

Discriminators resolved:
  - `< 1.15` boundary: which L flip to ink vs stay win?
  - font: Calibri vs Times New Roman?
  - widow 0 vs 1: does the split (widow=0) also use ink?
  - no-type docGrid vs no docGrid?

Usage: python _pb_notypemult_gen.py gen | measure | read
"""
import os, sys, json, zipfile

OUTDIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                      "pipeline_data", "_pb_notypemult")

# natural (single, LM0) line heights, pt, at 12pt
NAT = {"Calibri": 14.648, "Times New Roman": 13.799}
PAGE_H = 841.89  # A4 pt
TOP_TW = 1440    # 72pt top margin

# the 3-wrap target text (long enough to wrap to exactly 3 lines at ~450pt width)
TARGET_TEXT = ("Use your carbon footprint calculator results and the guidance in this "
               "section to set a realistic annual reduction target that your whole "
               "organisation can actually commit to and measure against each quarter")

FONTS = ["Calibri", "Times New Roman"]
LINES = [240, 254, 257, 259, 264, 276, 360, 480]  # 1.0 .. 2.0x auto


def rpr(font, sz=24):
    return (f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}"/>'
            f'<w:sz w:val="{sz}"/>')


def exact_filler(font, twips):
    """One empty paragraph of exact height `twips`."""
    r = rpr(font)
    return (f'<w:p><w:pPr><w:spacing w:after="0" w:line="{twips}" w:lineRule="exact"/>'
            f'<w:rPr>{r}</w:rPr></w:pPr></w:p>')


def target_para(font, line, widow):
    r = rpr(font)
    wc = '' if widow else '<w:widowControl w:val="0"/>'
    return (f'<w:p><w:pPr>{wc}<w:spacing w:after="0" w:line="{line}" w:lineRule="auto"/>'
            f'<w:jc w:val="left"/><w:rPr>{r}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{r}</w:rPr><w:t>{TARGET_TEXT}</w:t></w:r></w:p>')


def build(font, line, widow, grid, bt_tw, line0_top):
    """Filler places the target line-0 top at `line0_top` pt, bottom margin bt_tw."""
    # filler height needed above the target
    fill_pt = line0_top - 72.0
    fill_tw = int(round(fill_pt * 20))
    # split into 400tw blocks + remainder (exact paras can't exceed practical; use chunks)
    body = ""
    remaining = fill_tw
    while remaining > 400:
        body += exact_filler(font, 400)
        remaining -= 400
    if remaining > 20:
        body += exact_filler(font, remaining)
    body += target_para(font, line, widow)
    # a trailing marker paragraph so the target is never the last (repagination)
    body += exact_filler(font, 200)
    grid_xml = '<w:docGrid w:linePitch="360"/>' if grid else ''
    mar = (f'<w:pgMar w:top="{TOP_TW}" w:right="1418" w:bottom="{bt_tw}" w:left="1418" '
           'w:header="851" w:footer="600" w:gutter="0"/>')
    sect = f'<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>{mar}{grid_xml}</w:sectPr>'
    return body + sect


def docxml(body):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'<w:body>{body}</w:body></w:document>')


def write_docx(path, body, font):
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
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
              f'<w:docDefaults><w:rPrDefault><w:rPr>{rpr(font)}</w:rPr></w:rPrDefault></w:docDefaults>'
              '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
              '</w:styles>')
    settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:settings>')
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/_rels/document.xml.rels", drels)
        z.writestr("word/document.xml", docxml(body))
        z.writestr("word/styles.xml", styles)
        z.writestr("word/settings.xml", settings)


# bottom-margin sweep: content_bottom = PAGE_H - bt_tw/20. Sweep cb over a window
# around the target line-2 ink bottom to bracket the keep->push flip.
def cases():
    out = []
    for font in FONTS:
        for line in LINES:
            for widow in (1,):          # widow=1 is the target's mode; widow=0 spot-checked below
                for grid in (True,):     # no-type docGrid is the target's mode; no-grid spot-checked
                    pitch = NAT[font] * line / 240.0
                    ink_bottom_rel = 2 * pitch + 12.0  # line0_top -> line2 ink bottom
                    # place line-2 ink bottom near 765 with a nominal 72pt bottom margin
                    line0_top = 765.0 - ink_bottom_rel
                    # sweep content_bottom from 758 to 776 (cb high = smaller bottom margin)
                    for cb in [758 + 1.2 * k for k in range(16)]:
                        bt_tw = int(round((PAGE_H - cb) * 20))
                        out.append((font, line, widow, grid, bt_tw, round(line0_top, 2), round(cb, 2)))
    # spot-checks: widow=0 and no-grid, only for Calibri L=257
    for widow in (0,):
        for grid in (True, False):
            pitch = NAT["Calibri"] * 257 / 240.0
            line0_top = 765.0 - (2 * pitch + 12.0)
            for cb in [758 + 1.2 * k for k in range(16)]:
                bt_tw = int(round((PAGE_H - cb) * 20))
                out.append(("Calibri", 257, widow, grid, bt_tw, round(line0_top, 2), round(cb, 2)))
    return out


def nm(font, line, widow, grid, cb):
    f = "C" if font == "Calibri" else "T"
    g = "g" if grid else "n"
    return f"nt_{f}_L{line}_w{widow}_{g}_cb{cb:.1f}.docx"


def gen():
    os.makedirs(OUTDIR, exist_ok=True)
    cs = cases()
    for font, line, widow, grid, bt_tw, line0_top, cb in cs:
        write_docx(os.path.join(OUTDIR, nm(font, line, widow, grid, cb)),
                   build(font, line, widow, grid, bt_tw, line0_top), font)
    print(f"wrote {len(cs)} docs to {OUTDIR}")
    json.dump([{"font": f, "line": l, "widow": w, "grid": g, "bt_tw": bt, "line0_top": lt, "cb": cb}
               for (f, l, w, g, bt, lt, cb) in cs],
              open(os.path.join(OUTDIR, "_cases.json"), "w"), indent=1)


def measure():
    import win32com.client
    cs = json.load(open(os.path.join(OUTDIR, "_cases.json")))
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False; word.DisplayAlerts = 0
    try:
        for c in cs:
            src = os.path.abspath(os.path.join(OUTDIR, nm(c["font"], c["line"], c["widow"],
                                                          c["grid"], c["cb"])))
            pdf = src[:-5] + ".pdf"
            d = word.Documents.Open(src, ReadOnly=True)
            d.ExportAsFixedFormat(pdf, 17)  # wdExportFormatPDF
            d.Close(False)
        print(f"exported {len(cs)} PDFs")
    finally:
        word.Quit()


def read():
    import fitz
    cs = json.load(open(os.path.join(OUTDIR, "_cases.json")))
    # marker: the target's 3 lines start with "Use your carbon", middle "commit"-ish,
    # last "quarter". Identify by finding the 3 consecutive lines of TARGET_TEXT.
    rows = []
    for c in cs:
        pdf = os.path.join(OUTDIR, nm(c["font"], c["line"], c["widow"], c["grid"], c["cb"])[:-5] + ".pdf")
        if not os.path.exists(pdf):
            continue
        doc = fitz.open(pdf)
        # collect all spans of the target text with their page + y
        hits = []  # (page, y0, y1, text)
        for pno in range(doc.page_count):
            pg = doc[pno]
            for blk in pg.get_text("dict")["blocks"]:
                for ln in blk.get("lines", []):
                    t = "".join(s["text"] for s in ln["spans"])
                    if "carbon footprint" in t or "organisation can" in t or "quarter" in t:
                        y0 = min(s["bbox"][1] for s in ln["spans"])
                        y1 = max(s["bbox"][3] for s in ln["spans"])
                        hits.append((pno, round(y0, 2), round(y1, 2), t[:30]))
        doc.close()
        pages = sorted(set(h[0] for h in hits))
        split = len(pages) > 1  # target's lines span >1 page = pushed/split
        last_bottom = max((h[2] for h in hits), default=None)
        rows.append({**c, "n_lines_found": len(hits), "pages": pages,
                     "split": split, "last_bottom": last_bottom})
    # group by (font, line, widow, grid): find flip cb (last kept -> first split)
    from collections import defaultdict
    groups = defaultdict(list)
    for r in rows:
        groups[(r["font"], r["line"], r["widow"], r["grid"])].append(r)
    print(f"{'font':16} {'L':>4} {'w':>1} {'g':>1} {'flip_cb':>8} {'ink_bot':>8} {'win_bot':>8}")
    summary = []
    for key, rs in sorted(groups.items()):
        rs.sort(key=lambda r: r["cb"])
        flip = None
        for i in range(len(rs) - 1):
            if not rs[i]["split"] and rs[i + 1]["split"]:
                flip = (rs[i]["cb"] + rs[i + 1]["cb"]) / 2
                break
        # ink bottom of line-2 (last kept sample) and its top
        kept = [r for r in rs if not r["split"] and r["n_lines_found"] == 3]
        ink_bot = kept[-1]["last_bottom"] if kept else None
        font, line, widow, grid = key
        pitch = NAT[font] * line / 240.0
        # line-2 top ≈ ink_bot - 12; win box bottom = top + pitch
        win_bot = (ink_bot - 12.0 + pitch) if ink_bot else None
        print(f"{font:16} {line:>4} {widow:>1} {int(grid):>1} "
              f"{flip if flip else 0:>8.2f} {ink_bot if ink_bot else 0:>8.2f} "
              f"{win_bot if win_bot else 0:>8.2f}")
        summary.append({"font": font, "line": line, "factor": round(line / 240, 4),
                        "widow": widow, "grid": grid, "flip_cb": flip,
                        "ink_bot": ink_bot, "win_bot": win_bot})
    json.dump({"rows": rows, "summary": summary},
              open(os.path.join(OUTDIR, "_result.json"), "w"), indent=1)
    print(f"\nwrote {os.path.join(OUTDIR, '_result.json')}")


if __name__ == "__main__":
    {"gen": gen, "measure": measure, "read": read}[sys.argv[1]]()
