# -*- coding: utf-8 -*-
"""Re-measure a line-text fixture (tests/fixtures/<name>/word.json) on the truth machine.

Each .docx named in word.json is exported to PDF by its own Word process (DispatchEx,
one document per process -- Word caches face substitutes and language state per
process), the PDF's lines are read back (characters grouped by page and baseline,
joined in x order, BEFORE / AFTER marker paragraphs dropped, as the Rust test reads
Oxi's elements), and the result is written next to the old file as word.new.json.
Prints every case whose lines changed.

    python tools/metrics/remeasure_fixture_lines.py tests/fixtures/body_punctuation_language
"""
import json, os, shutil, sys
sys.stdout.reconfigure(encoding="utf-8")
import pymupdf, win32com.client  # noqa: E402

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))


def word_lines(pdf):
    out = []
    for pg in pymupdf.open(pdf):
        rows = {}
        for b in pg.get_text("rawdict")["blocks"]:
            for l in b.get("lines", []):
                for s in l["spans"]:
                    for c in s["chars"]:
                        rows.setdefault(round(c["origin"][1] * 10), []).append((c["origin"][0], c["c"]))
        for k in sorted(rows):
            t = "".join(c for _, c in sorted(rows[k])).replace(" ", " ").strip()
            if t and t not in ("BEFORE", "AFTER"):
                out.append(t)
    return out


def export(docx, pdf):
    w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
    try:
        d = w.Documents.Open(docx, ReadOnly=True); d.ExportAsFixedFormat(pdf, 17); d.Close(False)
    finally:
        w.Quit()


if __name__ == "__main__":
    root = os.path.join(REPO, sys.argv[1]) if not os.path.isabs(sys.argv[1]) else sys.argv[1]
    cases = json.load(open(os.path.join(root, "word.json"), encoding="utf-8"))
    out = os.path.join(REPO, "pipeline_data", "_fixture_truth", os.path.basename(root.rstrip("/\\")))
    os.makedirs(out, exist_ok=True)
    changed = 0
    for case in cases:
        name = case["name"]
        docx = os.path.join(out, name + ".docx"); pdf = docx[:-5] + ".pdf"
        if not os.path.exists(pdf):
            shutil.copy(os.path.join(root, name + ".docx"), docx)
            export(docx, pdf)
        new = word_lines(pdf)
        if new != case["lines"]:
            changed += 1
            print("%s\n  old %s\n  new %s" % (name, case["lines"], new), flush=True)
        case["lines"] = new
    json.dump(cases, open(os.path.join(root, "word.new.json"), "w", encoding="utf-8"), indent=2)
    print("changed %d / %d -> %s" % (changed, len(cases), os.path.join(root, "word.new.json")))
