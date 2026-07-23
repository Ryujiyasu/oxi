#!/usr/bin/env python3
"""usePrinterMetrics single-printer isolation probe (REPORT_usePrinterMetrics_isolation_probe.md).

Clones the faithful host legal__0011dcc, rewrites only word/settings.xml (the
W/P matrix) and word/document.xml (7 specimens, each on its own page + bookmark),
then measures line counts / page positions with one DispatchEx Word instance on
the UNCHANGED active printer. Deliverable = the isolated W/P/NBSP/firstLine
behavior (NOT a fixed 6.72pt constant — implementation stays blocked per v30).

    python _pb_upm_isolation_gen.py gen       # build the 5 docx
    python _pb_upm_isolation_gen.py measure    # Word COM + PDF export
    python _pb_upm_isolation_gen.py oxi         # Oxi GDI control (line counts)
"""
import os, sys, copy, zipfile
from lxml import etree

HERE = os.path.dirname(os.path.abspath(__file__))
HOST = os.path.join(HERE, "..", "..", "pipeline_data", "docx_corpus", "en",
                    "legal", "0011dcc222423b30.docx")
OUTDIR = os.path.join(HERE, "..", "..", "pipeline_data", "_pb_upm_isolation")

WNS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML = "http://www.w3.org/XML/1998/namespace"
NS = {"w": WNS}
def qn(t): return "{%s}%s" % (WNS, t)

VARIANTS = {
    # name: (wpJustification, usePrinterMetrics, forced_val)
    "W0P0": (False, False, None),
    "W1P0": (True,  False, None),
    "W0P1": (False, True,  None),
    "W1P1": (True,  True,  None),
    "C00":  (True,  True,  "0"),
}

NBSP = " "
I75_NBSP  = "(2)" + NBSP + NBSP + "expression of an intent to be partners in the business;"
I75_ASCII = "(2)  expression of an intent to be partners in the business;"
I25_TEXT  = ("The following section was amended by the 89th Legislature. Pending "
             "publication of the current statutes, see S.B. 29, 89th Legislature, "
             "Regular Session, for amendments affecting the following section.")
M60 = "M" * 60

# (name, jc, firstLine|None, text)
CASES = [
    ("C25_CENTER",      "center", None, I25_TEXT),
    ("J75_NBSP_F0000",  "both",   None, I75_NBSP),
    ("J75_NBSP_F0720",  "both",   720,  I75_NBSP),
    ("J75_NBSP_F1440",  "both",   1440, I75_NBSP),
    ("J75_NBSP_F2160",  "both",   2160, I75_NBSP),
    ("J75_ASCII_F1440", "both",   1440, I75_ASCII),
    ("CAL_M60",         "left",   None, M60),
]
CASE_NAMES = [c[0] for c in CASES]


def rewrite_settings(raw, wpj, upm, forced_val):
    root = etree.fromstring(raw)
    compat = root.find(".//w:compat", NS)
    assert compat is not None, "faithful host must have <w:compat>"
    for local, enabled in (("wpJustification", wpj), ("usePrinterMetrics", upm)):
        node = compat.find("w:" + local, NS)
        assert node is not None, "faithful host must contain <w:%s>" % local
        if not enabled:
            compat.remove(node)          # OFF variant: remove only that element
        elif forced_val is None:
            node.attrib.pop(qn("val"), None)  # ON: bare, no val
        else:
            node.set(qn("val"), forced_val)   # C00: val="0"
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def add_case(body, name, jc, first_line, text, bid):
    p = etree.SubElement(body, qn("p"))
    ppr = etree.SubElement(p, qn("pPr"))
    etree.SubElement(ppr, qn("pageBreakBefore"))          # order: pbb -> spacing -> ind -> jc
    sp = etree.SubElement(ppr, qn("spacing")); sp.set(qn("line"), "480"); sp.set(qn("lineRule"), "auto")
    if first_line is not None:
        ind = etree.SubElement(ppr, qn("ind")); ind.set(qn("firstLine"), str(first_line))
    jce = etree.SubElement(ppr, qn("jc")); jce.set(qn("val"), jc)
    bs = etree.SubElement(p, qn("bookmarkStart")); bs.set(qn("id"), str(bid)); bs.set(qn("name"), name)
    r = etree.SubElement(p, qn("r"))
    t = etree.SubElement(r, qn("t")); t.set("{%s}space" % XML, "preserve"); t.text = text
    etree.SubElement(p, qn("bookmarkEnd")).set(qn("id"), str(bid))


def rewrite_document(raw):
    root = etree.fromstring(raw)
    body = root.find("w:body", NS)
    assert body is not None
    sect = body.find("w:sectPr", NS)
    assert sect is not None, "faithful host body must end with sectPr"
    saved_sect = copy.deepcopy(sect)
    for child in list(body):
        body.remove(child)
    for bid, (name, jc, fl, text) in enumerate(CASES, 1):
        add_case(body, name, jc, fl, text, bid)
    body.append(saved_sect)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def gen():
    os.makedirs(OUTDIR, exist_ok=True)
    with zipfile.ZipFile(HOST) as src:
        original = {n: src.read(n) for n in src.namelist()}
    new_doc = rewrite_document(original["word/document.xml"])
    for name, (wpj, upm, forced_val) in VARIANTS.items():
        parts = dict(original)
        parts["word/document.xml"] = new_doc
        parts["word/settings.xml"] = rewrite_settings(original["word/settings.xml"], wpj, upm, forced_val)
        out = os.path.join(OUTDIR, name + ".docx")
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as dst:
            for pn, data in parts.items():
                dst.writestr(pn, data)
        print("wrote", out)
    # self-check the generated files
    for name in VARIANTS:
        with zipfile.ZipFile(os.path.join(OUTDIR, name + ".docx")) as z:
            doc = z.read("word/document.xml").decode("utf-8")
            st = z.read("word/settings.xml").decode("utf-8")
        assert "w:w:val" not in st, "%s: doubled prefix" % name
        # NBSP count: 4 NBSP cases * 2 = 8 across the doc
        assert doc.count(NBSP) == 8, "%s: NBSP=%d (want 8)" % (name, doc.count(NBSP))
        wpj_present = "<w:wpJustification" in st
        upm_present = "<w:usePrinterMetrics" in st
        exp_wpj, exp_upm, fv = VARIANTS[name]
        assert wpj_present == exp_wpj, "%s wpJustification present=%s" % (name, wpj_present)
        assert upm_present == exp_upm, "%s usePrinterMetrics present=%s" % (name, upm_present)
        if fv == "0":
            assert 'w:wpJustification w:val="0"' in st and 'w:usePrinterMetrics w:val="0"' in st, "%s val=0 missing" % name
        # compat order preserved (wpJustification before usePrinterMetrics when both present)
        if wpj_present and upm_present:
            assert st.index("wpJustification") < st.index("usePrinterMetrics"), "%s compat order" % name
    print("self-check OK (settings/NBSP/compat-order/CT_OnOff)")


def measure():
    import json, win32com.client, win32print, ctypes
    gdi32 = ctypes.windll.gdi32
    os.makedirs(OUTDIR, exist_ok=True)
    # printer / DPI from the DEFAULT printer DC (not screen). LOGPIXELSX=88, LOGPIXELSY=90.
    default_printer = win32print.GetDefaultPrinter()
    hdc = gdi32.CreateDCW("WINSPOOL", default_printer, None, None)
    dpi_x = gdi32.GetDeviceCaps(hdc, 88) if hdc else None
    dpi_y = gdi32.GetDeviceCaps(hdc, 90) if hdc else None
    if hdc:
        gdi32.DeleteDC(hdc)

    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    active_printer = str(word.ActivePrinter)
    word_version = str(word.Version)
    out = {"word_active_printer": active_printer, "windows_default_printer": default_printer,
           "dpi_x": dpi_x, "dpi_y": dpi_y, "word_version": word_version, "variants": {}}
    try:
        for name in VARIANTS:
            path = os.path.abspath(os.path.join(OUTDIR, name + ".docx"))
            doc = word.Documents.Open(path, ReadOnly=True)
            doc.Repaginate()
            total_pages = int(doc.ComputeStatistics(2))  # wdStatisticPages
            v = {"total_pages": total_pages, "cases": {}}
            for bn in CASE_NAMES:
                try:
                    r = doc.Bookmarks(bn).Range
                    start = doc.Range(r.Start, r.Start)
                    e_pos = max(r.Start, r.End - 1)
                    end = doc.Range(e_pos, e_pos)
                    v["cases"][bn] = {
                        "lines": int(r.ComputeStatistics(1)),  # wdStatisticLines
                        "page_start": int(start.Information(3)),
                        "page_end": int(end.Information(3)),
                        "x_start": round(float(start.Information(5)), 2),
                        "x_end": round(float(end.Information(5)), 2),
                        "y_start": round(float(start.Information(6)), 2),
                        "y_end": round(float(end.Information(6)), 2),
                    }
                except Exception as ex:
                    v["cases"][bn] = {"error": str(ex)}
            pdf = os.path.abspath(os.path.join(OUTDIR, name + ".pdf"))
            doc.ExportAsFixedFormat(pdf, 17)  # wdExportFormatPDF
            doc.Close(False)
            out["variants"][name] = v
            print(name, "pages=%d" % total_pages)
    finally:
        word.Quit()
    with open(os.path.join(OUTDIR, "_result.json"), "w", encoding="utf-8") as f:
        json.dump(out, f, indent=2, ensure_ascii=False)
    print("wrote _result.json;", active_printer, "dpi", dpi_x, "x", dpi_y, "Word", word_version)


def oxi():
    """Oxi S994 renderer control: line counts per specimen (each on its own page)."""
    import json, subprocess, tempfile
    r = os.path.join(HERE, "..", "oxi-gdi-renderer", "target", "release", "oxi-gdi-renderer.exe")
    res = {}
    for name in VARIANTS:
        docx = os.path.abspath(os.path.join(OUTDIR, name + ".docx"))
        pref = os.path.join(tempfile.gettempdir(), "upm_iso_" + name)
        dump = pref + ".json"
        subprocess.run([r, docx, pref + ".png", "96", "--dump-layout=" + dump],
                       capture_output=True)
        d = json.load(open(dump, encoding="utf-8"))
        pages = d.get("pages", d) if isinstance(d, dict) else d
        if isinstance(pages, dict) and "pages" in pages:
            pages = pages["pages"]
        # each specimen is alone on its page (pageBreakBefore) — count unique y text lines per page
        page_lines = []
        for pg in pages:
            els = pg.get("elements") if isinstance(pg, dict) else pg
            ys = set()
            for e in (els or []):
                if (e.get("text") or "").strip():
                    ys.add(round(float(e.get("y", 0)), 1))
            page_lines.append(len(ys))
        res[name] = page_lines
    # map pages to CASES in order (page1 may be empty/first; specimens follow pageBreakBefore)
    print("=== Oxi line counts per page (each specimen on its own page) ===")
    for name in VARIANTS:
        print("%-6s %s" % (name, res[name]))
    # control assertions
    print("--- controls (P unimplemented in Oxi) ---")
    print("W0P0==W0P1:", res["W0P0"] == res["W0P1"])
    print("W1P0==W1P1:", res["W1P0"] == res["W1P1"])
    print("C00 ==W0P0:", res["C00"]  == res["W0P0"])
    json.dump(res, open(os.path.join(OUTDIR, "_oxi_control.json"), "w"), indent=2)


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "gen"
    {"gen": gen, "measure": measure, "oxi": oxi}[cmd]()
