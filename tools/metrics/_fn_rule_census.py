# -*- coding: utf-8 -*-
"""Footnote rule census: Word's separator rule against Oxi's, page by page.

Every corpus document with real footnotes and a one-paragraph separator.
Word: one fresh process per document exports a PDF (cached beside the run dir);
the rule is the 144pt-wide horizontal stroke under 1.6pt tall. Oxi: --dump-layout,
the element of the same width and height. Prints, per document, each page's
Word rule top / thickness and each exe's (top - Word).

    python tools/metrics/_fn_rule_census.py <exe>[;<exe>] [out_dir]
"""
import glob, json, os, re, subprocess, sys, tempfile, zipfile
from pathlib import Path
sys.stdout.reconfigure(encoding="utf-8")
REPO = Path(__file__).resolve().parents[2]
CORPUS = REPO / "pipeline_data" / "docx_corpus"


def docs():
    for f in sorted(glob.glob(str(CORPUS / "*" / "*" / "*.docx"))):
        try:
            z = zipfile.ZipFile(f)
            x = z.read("word/footnotes.xml").decode("utf-8", "replace")
        except Exception:
            continue
        m = re.search(r'<w:footnote [^>]*w:type="separator"[^>]*>(.*?)</w:footnote>', x, re.S)
        if m and re.search(r'<w:footnote (?![^>]*w:type=)', x) and len(re.findall(r"<w:p[ >]", m.group(1))) == 1:
            yield Path(f)


def word_rules(docx, pdf):
    if not pdf.exists():
        import win32com.client
        w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
        try:
            d = w.Documents.Open(str(docx), ReadOnly=True)
            d.ExportAsFixedFormat(str(pdf), 17); d.Close(False)
        finally:
            w.Quit()
    import pymupdf
    out = []
    for i, pg in enumerate(pymupdf.open(pdf)):
        for d in pg.get_drawings():
            r = d["rect"]
            if abs(r.width - 144) < 0.6 and r.height < 1.6:
                out.append((i, round(r.y0, 2), round(r.height, 2)))
    return sorted(set(out))


def oxi_rules(exe, docx):
    t = tempfile.mkdtemp(); lay = os.path.join(t, "l.json")
    subprocess.run([exe, str(docx), os.path.join(t, "q"), "--dump-layout=" + lay], capture_output=True, timeout=1800)
    out = []
    for i, pg in enumerate(json.load(open(lay, encoding="utf-8"))["pages"]):
        for e in pg["elements"]:
            if abs(e.get("w", 0) - 144) < 0.6 and e.get("h", 9) < 1.6:
                out.append((i, round(e["y"], 2), round(e["h"], 2)))
    return sorted(set(out))


if __name__ == "__main__":
    exes = [e for e in sys.argv[1].split(";") if e]
    out = Path(sys.argv[2]) if len(sys.argv) > 2 else REPO / "pipeline_data" / "_fn_rule_census"
    out.mkdir(parents=True, exist_ok=True)
    for docx in docs():
        name = f"{docx.parent.parent.name}_{docx.parent.name}_{docx.stem[:8]}"
        w = word_rules(docx, out / f"{name}.pdf")
        os_ = [dict((p, (y, h)) for p, y, h in oxi_rules(e, docx)) for e in exes]
        cells = []
        for p, y, h in w:
            parts = [f"p{p + 1} W {y:7.2f}/{h:.2f}"]
            for o in os_:
                parts.append(f"{o[p][0] - y:+7.2f}/{o[p][1]:.2f}" if p in o else "   -   ")
            cells.append(" ".join(parts))
        for k, o in enumerate(os_):
            dys = sorted(abs(o[p][0] - y) for p, y, _ in w if p in o)
            if dys:
                cells.append(f"exe{k}: {len(dys)} pages, max|dy| {dys[-1]:.2f} median {dys[len(dys) // 2]:.2f}")
        extra = [sorted(set(o) - {p for p, _, _ in w}) for o in os_]
        print(f"{name:28} " + (" | ".join(cells[:3] + [c for c in cells if c.startswith("exe")]) or "no Word rule")
              + (f"  oxi-only pages {extra}" if any(extra) else ""), flush=True)
