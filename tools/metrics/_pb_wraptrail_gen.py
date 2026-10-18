# -*- coding: utf-8 -*-
"""<w:wrapTrailSpaces/>: where do the spaces of a multi-space run at a line break go?

legal__0011dcc (compat 14, wpJustification, wrapTrailSpaces, Courier New 12):
«…TO BIND PARTNERSHIP.  (a)  The action…» -- when line 1 is already over the
column, the SECOND of the two spaces starts line 2 (Word draws line 2 from
72 + 7.2), which costs line 2 a character and wraps «the»; with room on line 1
both spaces stay at its end.  Oxi does not read wrapTrailSpaces.

Arms: settings with / without wrapTrailSpaces (host 0011dcc's settings, the flag
removed for the second) x k = 2 / 3 spaces after «end.»; line 1 = 12 «mmmm»
words + «end.»; right indent swept.  Per right indent: how many leading spaces
line 2 carries (x0 - 72) / 7.2.

    python tools/metrics/_pb_wraptrail_gen.py
"""
import os, re, sys, zipfile
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.stdout.reconfigure(encoding="utf-8")
import _pb_c14floor_gen as P

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "pipeline_data", "_pb_wraptrail")
ADV = P.ADV
LINE1 = " ".join(["mmmm"] * 12) + " end."            # 64 chars
RIGHTS = list(range(-800, 400, 20))


def build(name, wrap_flag, k):
    specs = []
    for r in RIGHTS:
        text = LINE1 + " " * k + "next words follow here on the second line of the paragraph."
        specs.append(("%s_r%d" % (name, r), text, {"ri": r, "cand": "x"}))
    P.specimens = lambda: specs
    with zipfile.ZipFile(P.HOST) as z:
        parts = {n: z.read(n) for n in z.namelist()}
    st = parts["word/settings.xml"].decode("utf-8")
    if not wrap_flag:
        st = st.replace("<w:wrapTrailSpaces/>", "")
    parts["word/settings.xml"] = st.encode("utf-8")
    doc = P.rewrite_document(parts["word/document.xml"], specs).decode("utf-8")
    doc = doc.replace('<w:jc w:val="both"/>', "")       # left-aligned: no justification noise
    parts["word/document.xml"] = doc.encode("utf-8")
    os.makedirs(OUT, exist_ok=True)
    p = os.path.join(OUT, name + ".docx")
    with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as z:
        for n, b in parts.items():
            z.writestr(n, b)
    return p


def measure(p):
    import fitz, win32com.client
    pdf = p[:-5] + ".pdf"
    w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
    try:
        d = w.Documents.Open(os.path.abspath(p), ReadOnly=True); d.ExportAsFixedFormat(os.path.abspath(pdf), 17); d.Close(False)
    finally:
        w.Quit()
    out = []
    for r, pg in zip(RIGHTS, fitz.open(pdf)):
        rows = {}
        for b in pg.get_text("rawdict")["blocks"]:
            for l in b.get("lines", []):
                for s in l["spans"]:
                    for c in s["chars"]:
                        if c["c"].strip():
                            rows.setdefault(round(c["origin"][1]), []).append(c)
        lines = [sorted(v, key=lambda c: c["origin"][0]) for _, v in sorted(rows.items())]
        body = [l for l in lines if l and l[0]["origin"][1] < 700]
        if len(body) < 2:
            out.append((r, None, None)); continue
        l1 = "".join(c["c"] for c in body[0]); l2x = body[1][0]["bbox"][0]
        out.append((r, l1[-6:], round((l2x - 72.0) / ADV, 2)))
    return out


if __name__ == "__main__":
    for wrap_flag in (True, False):
        for k in (2, 3):
            name = "wrap%s_k%d" % ("ON" if wrap_flag else "off", k)
            res = measure(build(name, wrap_flag, k))
            prev = None
            print("==", name)
            for r, end1, lead in res:
                if (end1, lead) != prev:
                    avail = P.COL_PT - r / 20
                    print("   right %5d avail %6.1f  line1 ends %-8r line2 leading spaces %s" % (r, avail, end1, lead))
                prev = (end1, lead)
