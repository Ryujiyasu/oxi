# -*- coding: utf-8 -*-
"""Read the deleted-paragraph-mark probe. Usage: _pb_delmark_read.py word|oxi

Word is measured the way the truth generators now do it: revisions ACCEPTED on a
copy (see measure_pagination_word.open_clean), which is the state
ShowRevisions::Final renders.
"""
import os, sys, json, subprocess, shutil, tempfile, time
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _pb_delmark_gen import TAGS, OUT

REND = os.path.abspath("tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe")


def word_lines(docx):
    import fitz, win32com.client
    pdf = docx[:-5] + ".pdf"
    if not os.path.exists(pdf):
        fd, tmp = tempfile.mkstemp(suffix=".docx", prefix="dm_")
        os.close(fd)
        shutil.copy(docx, tmp)
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        try:
            d = word.Documents.Open(os.path.abspath(tmp), ReadOnly=False)
            try:
                if d.Revisions.Count:
                    d.Revisions.AcceptAll()
                    d.Repaginate()
                    time.sleep(0.5)
            except Exception:
                pass
            d.SaveAs2(os.path.abspath(pdf), FileFormat=17)
            d.Close(False)
        finally:
            word.Quit()
            try:
                os.remove(tmp)
            except Exception:
                pass
    doc = fitz.open(pdf)
    out = []
    for b in doc[0].get_text("dict")["blocks"]:
        for l in b.get("lines", []):
            for s in l["spans"]:
                t = s["text"].strip()
                if t:
                    out.append((round(s["origin"][1], 1), round(s["origin"][0], 2), t))
    return sorted(out)


def oxi_lines(docx):
    dump = docx[:-5] + ".layout.json"
    subprocess.run([REND, docx, docx[:-5] + "_r", "96", "--dump-layout=" + dump],
                   capture_output=True)
    d = json.load(open(dump, encoding="utf-8"))
    rows = {}
    for e in d["pages"][0]["elements"]:
        t = (e.get("text") or "")
        if not t.strip():
            continue
        rows.setdefault(round(e["y"], 1), []).append((round(e["x"], 2), t))
    return sorted((y, min(x for x, _ in v), " ".join(t for _, t in sorted(v)))
                  for y, v in rows.items())


mode = sys.argv[1] if len(sys.argv) > 1 else "word"
for tag in TAGS:
    docx = os.path.join(OUT, tag + ".docx")
    rows = word_lines(docx) if mode == "word" else oxi_lines(docx)
    print("\n--- %s  %s ---" % (mode.upper(), tag))
    for y, x, t in rows:
        if t.startswith("#"):
            continue
        print("   y=%7.1f  x=%7.2f  %r" % (y, x, t[:44]))
