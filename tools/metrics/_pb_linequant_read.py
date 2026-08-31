# -*- coding: utf-8 -*-
"""Read the linequant probe. Usage: _pb_linequant_read.py word|oxi [--sweep ...]

Two tables:

  ORIGIN  the first baseline against the top margin. A quantised origin moves in
          jumps of one grid unit; an exact one tracks the 0.05pt input steps.
  PITCH   for one arm, every line's deviation from n * P_exact. Cumulative
          rounding keeps the deviation inside half a unit and does not drift;
          per-pitch rounding random-walks away from zero.

P_exact is taken from the arm itself (the mean pitch), so no font table is
needed and the reading does not depend on any metric Oxi also computes.
"""
import os, sys, json, subprocess, collections
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _pb_linequant_gen import FACES, OUT, parse_sweep

REND = os.path.abspath("tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe")


def word_baselines(docx):
    import fitz, win32com.client
    pdf = docx[:-5] + ".pdf"
    if not os.path.exists(pdf):
        app = win32com.client.DispatchEx("Word.Application")
        app.Visible = False
        app.DisplayAlerts = 0
        try:
            d = app.Documents.Open(os.path.abspath(docx), ReadOnly=True)
            d.SaveAs2(os.path.abspath(pdf), FileFormat=17)
            d.Close(False)
        finally:
            app.Quit()
    doc = fitz.open(pdf)
    ys = []
    for blk in doc[0].get_text("dict")["blocks"]:
        for l in blk.get("lines", []):
            if [s for s in l["spans"] if s["text"].strip()]:
                ys.append(round(l["spans"][0]["origin"][1], 3))
    doc.close()
    return sorted(ys)


def oxi_baselines(docx):
    dump = docx[:-5] + ".layout.json"
    subprocess.run([REND, docx, docx[:-5] + "_r", "96", "--dump-layout=" + dump],
                   capture_output=True)
    d = json.load(open(dump, encoding="utf-8"))
    ys = set()
    for e in d["pages"][0]["elements"]:
        if (e.get("text") or "").strip():
            ys.add(round(e["y"], 3))          # raw box top: no convention mixing
    return sorted(ys)


mode = sys.argv[1] if len(sys.argv) > 1 else "word"
reader = word_baselines if mode == "word" else oxi_baselines
sw = parse_sweep(sys.argv)

print("%s   ORIGIN: first line vs top margin (input step 0.05pt)\n" % mode.upper())
print("  top_tw  top_pt |" + "".join("  %-11s" % lab for lab, _, _, _ in FACES))
prev = {}
for t in sw:
    cells = []
    for lab, _, _, _ in FACES:
        p = os.path.join(OUT, "%s_t%d.docx" % (lab, t))
        if not os.path.exists(p):
            cells.append("   MISSING  ")
            continue
        ys = reader(p)
        y0 = ys[0] if ys else float("nan")
        d = y0 - prev.get(lab, y0)
        prev[lab] = y0
        cells.append("  %8.3f%+.2f" % (y0, d))
    print("  %6d %7.2f |%s" % (t, t / 20.0, "".join(cells)))

print("\n%s   PITCH: deviation from n * mean_pitch (one arm per face, top=1440)\n"
      % mode.upper())
for lab, _, _, _ in FACES:
    p = os.path.join(OUT, "%s_t%d.docx" % (lab, sw[0]))
    if not os.path.exists(p):
        continue
    ys = reader(p)
    if len(ys) < 6:
        print("  %-11s only %d lines" % (lab, len(ys)))
        continue
    pitches = [round(b - a, 3) for a, b in zip(ys, ys[1:])]
    mean = sum(pitches) / len(pitches)
    devs = [ys[n] - ys[0] - n * mean for n in range(len(ys))]
    c = collections.Counter(pitches)
    vals = sorted(c)
    print("  %-11s n=%d mean_pitch=%.4f  distinct=%s"
          % (lab, len(ys), mean, [(v, c[v]) for v in vals[:4]]))
    print("               unit(max-min)=%.3f  dev range=[%+.3f, %+.3f]  last dev=%+.3f"
          % (vals[-1] - vals[0], min(devs), max(devs), devs[-1]))
