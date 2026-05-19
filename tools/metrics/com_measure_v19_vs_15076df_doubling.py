"""COM-measure V19 vs 15076df to test the compressPunctuation gate hypothesis.

S111 found: 15076df L12 has balanceSBDB + compressPunctuation + cs=-9tw,
            and Word's effective cw per CJK = 9.75pt (≈ single cs applied).
S56 Finding 3 says: V19/V25/V26/V27 confirmed balanceSBDB DOUBLES cs.
            But V19/V27 do NOT have compressPunctuation.

Test: COM-measure V19 paragraph (same script structure as 15076df L12).
  If Word doubles → per-CJK cw ≈ 9.6pt = 10.5 - 0.9 (= cs × 2)
  If Word single → per-CJK cw ≈ 10.05pt = 10.5 - 0.45 (= cs × 1)
  At pixel rounding @96dpi:
    single → 10.05pt = 13.4 px → rounds to 13 px = 9.75pt
    double → 9.60pt  = 12.8 px → rounds to 13 px = 9.75pt
  ← BOTH ROUND TO 9.75pt — can't distinguish from CJK alone!

The trick: use a HALFWIDTH char where the difference matters:
  - At natural=5.25pt half + cs=-0.45 single: 4.8 → 6.4 px → 6 px = 4.5pt
  - At natural=5.25pt half + cs=-0.9 double:  4.35 → 5.8 px → 6 px = 4.5pt
  Still same! Pixel rounding masks both.

Actually the way to detect: measure across MANY consecutive chars and
look at cumulative drift. Word does NOT round per-char to pixels — it
accumulates exact and renders at sub-pixel positions (DirectWrite).

Approach: take total content width = end_x - start_x for many chars,
divide by char count, compare to single vs double prediction.
"""
import os
import sys
import io
import win32com.client

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))

cases = [
    ("V19_cs=-9_no_compressPunc",
        os.path.normpath(os.path.join(REPO, "tools/metrics/_repros/repro_pcw_V19.docx"))),
    ("V25_cs=-20_no_compressPunc",
        os.path.normpath(os.path.join(REPO, "tools/metrics/_repros/repro_pcw_V25.docx"))),
    ("V27_cs=-5_no_compressPunc",
        os.path.normpath(os.path.join(REPO, "tools/metrics/_repros/repro_pcw_V27.docx"))),
    ("15076df_cs=-9_WITH_compressPunc",
        os.path.normpath(os.path.join(REPO, "tools/golden-test/documents/docx/15076df085f5_tokumei_08_09.docx"))),
]

word = win32com.client.gencache.EnsureDispatch("Word.Application")
word.Visible = False

for label, docx_path in cases:
    print(f"\n========== {label} ==========")
    print(f"DOCX: {docx_path}")
    doc = word.Documents.Open(docx_path, ReadOnly=True)
    try:
        wdHoriz, wdVert = 5, 6
        # Pick a paragraph with CJK content
        target_para = None
        target_idx = None
        for i in range(1, min(doc.Paragraphs.Count, 100) + 1):
            p = doc.Paragraphs(i)
            txt = p.Range.Text
            if "匿名" in txt and len(txt.replace("\r","").replace("\x07","")) >= 5:
                target_para = p
                target_idx = i
                break
        if target_para is None:
            print("  no 匿名 paragraph found")
            doc.Close(SaveChanges=False)
            continue
        p = target_para
        txt = p.Range.Text.replace("\r","").replace("\x07","")
        print(f"  para idx={target_idx} text={txt[:50]!r}")

        rng_start = p.Range.Start
        rng_end = p.Range.End
        chars = []
        prev_y = None
        for j in range(rng_start, min(rng_end, rng_start + 30)):
            r = doc.Range(j, j)
            x = r.Information(wdHoriz)
            y = r.Information(wdVert)
            nr = doc.Range(j, j + 1)
            ch = nr.Text
            chars.append({"x": x, "y": y, "ch": ch})
            prev_y = y
        # Collect first line: same y as char 0
        line1 = [c for c in chars if c["y"] == chars[0]["y"]]
        if len(line1) < 4:
            print("  too few L1 chars")
            doc.Close(SaveChanges=False)
            continue
        # Compute advances for CJK chars only (exclude halfwidth/yakumono)
        cjk_advs = []
        for k in range(1, len(line1)):
            adv = line1[k]["x"] - line1[k-1]["x"]
            prev_ch = line1[k-1]["ch"]
            cjk = ord(prev_ch) >= 0x3040 and ord(prev_ch) <= 0x9FFF
            if cjk:
                cjk_advs.append(adv)
        print(f"  L1 chars: {len(line1)}")
        # Print first 10 chars with advances
        for k in range(min(len(line1), 10)):
            if k == 0:
                print(f"    [{k:2}] {line1[k]['ch']!r} x={line1[k]['x']:.3f}")
            else:
                adv = line1[k]["x"] - line1[k-1]["x"]
                cjk = ord(line1[k-1]["ch"]) >= 0x3040 and ord(line1[k-1]["ch"]) <= 0x9FFF
                tag = "CJK" if cjk else "   "
                print(f"    [{k:2}] {line1[k]['ch']!r} x={line1[k]['x']:.3f} adv-of-prev[{line1[k-1]['ch']!r}]={adv:.3f}pt {tag}")
        if cjk_advs:
            avg_cjk = sum(cjk_advs) / len(cjk_advs)
            print(f"  CJK advance stats: n={len(cjk_advs)} min={min(cjk_advs):.3f} max={max(cjk_advs):.3f} avg={avg_cjk:.3f}pt")
            # At fs=10.5pt:
            #   single cs=-0.45 → 10.05 → 13.4 px @96dpi → 9.75pt rendered
            #   double cs=-0.9  → 9.60  → 12.8 px @96dpi → 9.75pt rendered (same!)
            # But cumulative diverges:
            #   single sum over 10 chars = 100.5pt
            #   double sum over 10 chars = 96.0pt
            #   actual Word avg × 10 = avg_cjk × 10
            print(f"  Predict if SINGLE cs (no doubling): avg≈10.05pt (cs=-9tw)")
            print(f"  Predict if DOUBLE cs (S56 F3):      avg≈9.60pt  (cs=-9tw)")
            print(f"  WORD MEASURED avg: {avg_cjk:.3f}pt")
    finally:
        doc.Close(SaveChanges=False)

word.Quit()
