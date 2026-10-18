# -*- coding: utf-8 -*-
"""wpJustification allowance in compat 14: proportional to the line WIDTH or to the
number of SPACES?

`_pb_c14wpj_gen.py` varied the gap count G together with the line width (and its
left-indent arms could not separate them: with fixed content the flip happens at
the same available width).  Here G is fixed (6 words + candidate = 6 spaces) and
the WORD LENGTH changes (4 / 6 / 8 'm'), so the available width at the flip
changes while the space count does not.  Width model: allowance = W x 281/7200;
space model: allowance = k x 6.  Host = legal__0011dcc (wpJustification, Courier
New 12, compat 14), right indent swept in 0.6pt steps.

    python tools/metrics/_pb_c14wpj2_gen.py
"""
import json, os, sys, zipfile
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.stdout.reconfigure(encoding="utf-8")
import _pb_c14floor_gen as P

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "pipeline_data", "_pb_c14wpj2")
ADV = P.ADV
G = 6
SPECS = []
for L in (4, 6, 8):
    word = "m" * L
    before = (G * L + (G - 1)) * ADV            # G words, G-1 spaces
    need = ADV + 6 * ADV                        # space + «final.»
    for k in range(0, 50):
        R = need - 22.0 + 0.6 * k               # room left for space+candidate
        ri = int(round((P.COL_PT - before - R) * 20))
        if ri < 0:
            continue
        text = " ".join([word] * G) + " final."
        SPECS.append(("L%d_R%05d" % (L, int(round(R * 100))), text, {"L": L, "R": round(R, 2), "ri": ri,
                                                                      "before": round(before, 2), "cand": "final."}))
P.specimens = lambda: SPECS


def main():
    os.makedirs(OUT, exist_ok=True)
    with zipfile.ZipFile(P.HOST) as z:
        parts = {n: z.read(n) for n in z.namelist()}
    parts["word/document.xml"] = P.rewrite_document(parts["word/document.xml"], SPECS)
    with zipfile.ZipFile(os.path.join(OUT, "floor.docx"), "w", zipfile.ZIP_DEFLATED) as z:
        for n, b in parts.items():
            z.writestr(n, b)
    P.OUTDIR = OUT
    try:
        P.measure()
    except Exception:
        pass
    res = json.load(open(os.path.join(OUT, "_result.json")))
    for L in (4, 6, 8):
        rows = sorted((r for n, r in res.items() if n.startswith("L%d_" % L)), key=lambda r: r["R"])
        keep = [r["R"] for r in rows if r["com_lines"] == 1]
        wrap = [r["R"] for r in rows if r["com_lines"] == 2]
        hi = min(keep) if keep else None
        lo = max((w for w in wrap if hi is None or w < hi), default=None)
        before = rows[0]["before"]
        need = 7 * ADV
        allow = (need - hi, need - lo) if hi is not None and lo is not None else None
        width = before + (hi or 0)
        print("word len %d: flip room R in (%s, %s] -> allowance %s ; width at flip %.1f -> width model %.2f, space model (1.5/space) %.2f"
              % (L, lo, hi, tuple(round(a, 2) for a in allow) if allow else None, width, width * 281 / 7200, 1.5 * 6))


if __name__ == "__main__":
    main()
