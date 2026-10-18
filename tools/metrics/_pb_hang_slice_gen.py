# -*- coding: utf-8 -*-
"""Faithful slice for the line-end hang: keep ONE paragraph of a real document
(its styles, settings, theme, fonts untouched), repeat it once per page with its
right indent swept, and ask Word which line-end marks hang.

administrative__003f0741 «…commercial electric vehicle refueling.» wraps in the
2026-10 Word although its period would hang in every `_pb_hangmark_prop` arm
(face, alignment, line 1 / line 2, mid / end of paragraph).  This isolates
whether the document itself switches the hang off.

    python tools/metrics/_pb_hang_slice_gen.py <doc_id> "<text in the paragraph>" [start,stop,step]
    python tools/metrics/_pb_hang_slice_gen.py en/administrative__003f0741a36f3fad "commercial electric vehicle refueling"

Env SLICE_SETTINGS=bare replaces settings.xml with a compat-14-only one
(additive bisection: does a settings flag do it?).
"""
import os, re, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import feature_census as FC  # noqa: E402
import _pb_hangmark_prop_gen as H  # noqa: E402

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))


def build(did, needle, rights, out):
    z = zipfile.ZipFile(FC.load()[did]["path"])
    parts = {n: z.read(n) for n in z.namelist()}
    d = parts["word/document.xml"].decode("utf-8")
    i = d.find(needle)
    s = max(d.rfind("<w:p ", 0, i), d.rfind("<w:p>", 0, i)); e = d.find("</w:p>", i) + 6
    p = d[s:e]
    # SLICE_SUB="regex=>replacement;;regex=>replacement": edit the paragraph (subtractive / additive bisection)
    for rule in filter(None, os.environ.get("SLICE_SUB", "").split(";;")):
        pat, rep = rule.split("=>")
        p = re.sub(pat, rep, p)
    sect = re.findall(r"<w:sectPr[ >].*?</w:sectPr>", d, re.S)[-1]
    if os.environ.get("SLICE_SECT") == "bare":   # page size + margins only
        sect = "<w:sectPr>%s%s</w:sectPr>" % (re.search(r"<w:pgSz [^>]*/>", sect).group(0), re.search(r"<w:pgMar [^>]*/>", sect).group(0))
    paras = []
    for k, r in enumerate(rights):
        q = re.sub(r"<w:ind ([^>]*?)w:right=\"-?\d+\"", r"<w:ind \1", p)
        if "<w:ind " in q:
            q = q.replace("<w:ind ", '<w:ind w:right="%d" ' % r, 1)
        else:
            q = q.replace("<w:pPr>", '<w:pPr><w:ind w:right="%d"/>' % r, 1)
        if k:
            q = q.replace("<w:pPr>", "<w:pPr><w:pageBreakBefore/>", 1)
        paras.append(q)
    b0 = re.search(r"<w:body[^>]*>", d).end()   # some packages write <w:body ...> with attributes
    parts["word/document.xml"] = (d[:b0] + "".join(paras) + sect + "</w:body></w:document>").encode("utf-8")
    if os.environ.get("SLICE_SETTINGS") == "bare":
        parts["word/settings.xml"] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/></w:compat></w:settings>').encode()
    # SLICE_STYLES_SUB="regex=>replacement;;...": edit styles.xml (e.g. drop the docDefaults lang)
    st = parts["word/styles.xml"].decode("utf-8")
    for rule in filter(None, os.environ.get("SLICE_STYLES_SUB", "").split(";;")):
        pat, rep = rule.split("=>")
        st = re.sub(pat, rep, st, flags=re.S)
    parts["word/styles.xml"] = st.encode("utf-8")
    os.makedirs(os.path.dirname(out), exist_ok=True)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as w:
        for n, b in parts.items():
            w.writestr(n, b)
    # page geometry for the hang test: right edge = page width - right margin - right indent
    pw = int(re.search(r'<w:pgSz [^>]*w:w="(\d+)"', sect).group(1)) / 20
    rm = int(re.search(r'<w:pgMar [^>]*w:right="(\d+)"', sect).group(1)) / 20
    return pw - rm


def main():
    did, needle = sys.argv[1], sys.argv[2]
    rights = list(range(*map(int, (sys.argv[3] if len(sys.argv) > 3 else "0,1200,5").split(","))))
    tag = os.environ.get("SLICE_TAG") or os.environ.get("SLICE_SETTINGS", "faithful")
    out = os.path.join(REPO, "pipeline_data", "_pb_hang_slice", "%s_%s.docx" % (did.replace("/", "__"), tag))
    edge = build(did, needle, rights, out)
    H.RIGHTS = rights
    H.limit = lambda r: edge - r / 20
    pdf = out[:-5] + ".pdf"
    if os.path.exists(pdf):
        os.remove(pdf)
    print(tag, "max hang (pt) over the sweep:", H.word_hang(out))
    import fitz
    hung = []
    for r, pg in zip(rights, fitz.open(pdf)):
        for b in pg.get_text("rawdict")["blocks"]:
            for l in b.get("lines", []):
                ch = [c for s in l["spans"] for c in s["chars"] if c["c"].strip()]
                if ch and ch[-1]["bbox"][2] > H.limit(r) + 0.05:
                    hung.append((r, ch[-1]["c"], round(ch[-1]["bbox"][0] - H.limit(r), 2), round(ch[-1]["bbox"][2] - H.limit(r), 2)))
    print("line ends past the edge (right, char, start-edge, end-edge):", hung[:12])


if __name__ == "__main__":
    main()
