"""COM-measure '．' advance in 2 minimal repros differing only in
characterSpacingControl (compressPunctuation vs doNotCompress).
"""
import os
import sys
import io
import win32com.client

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))

cases = [
    ("compressPunc",
        os.path.normpath(os.path.join(REPO, "tools/metrics/_repros/repro_dot_15076_compressPunc.docx"))),
    ("doNotCompress",
        os.path.normpath(os.path.join(REPO, "tools/metrics/_repros/repro_dot_15076_doNotCompress.docx"))),
]

word = win32com.client.gencache.EnsureDispatch("Word.Application")
word.Visible = False

results = {}

for label, docx_path in cases:
    print(f"\n========== {label} ==========")
    print(f"DOCX: {docx_path}")
    doc = word.Documents.Open(docx_path, ReadOnly=True)
    try:
        wdHoriz, wdVert = 5, 6
        p = doc.Paragraphs(1)
        txt = p.Range.Text.replace("\r","").replace("\x07","")
        print(f"  text={txt!r}")
        rng_start = p.Range.Start
        rng_end = p.Range.End
        prev_y = None
        line_no = 1
        chars = []
        for j in range(rng_start, rng_end):
            r = doc.Range(j, j)
            x = r.Information(wdHoriz)
            y = r.Information(wdVert)
            nr = doc.Range(j, j + 1)
            ch = nr.Text
            if prev_y is not None and y != prev_y:
                line_no += 1
            chars.append({"x": x, "y": y, "ch": ch, "line": line_no})
            prev_y = y
        # Print all
        for k, c in enumerate(chars):
            adv = "    -   "
            if k > 0 and chars[k]["y"] == chars[k-1]["y"]:
                adv = f"{chars[k]['x']-chars[k-1]['x']:+8.3f}"
            prev_ch = repr(chars[k-1]['ch']) if k > 0 else "'-'"
            cur_ch = repr(c['ch'])
            print(f"    [{k:2}] L{c['line']} x={c['x']:.3f} y={c['y']:.3f} adv-of-prev[{prev_ch}]={adv}pt  cur={cur_ch}")
        # Specifically: advance of '．' (= position of '提' - position of '．')
        dot_idx = None
        for k, c in enumerate(chars):
            if c["ch"] == "．":
                dot_idx = k
                break
        if dot_idx is not None and dot_idx + 1 < len(chars) and chars[dot_idx]["y"] == chars[dot_idx+1]["y"]:
            dot_adv = chars[dot_idx+1]["x"] - chars[dot_idx]["x"]
            print(f"  ★ '．' advance = {dot_adv:.3f}pt (= {dot_adv*4/3:.1f} px @96dpi)")
            results[label] = dot_adv
        # Also '１' advance
        one_idx = None
        for k, c in enumerate(chars):
            if c["ch"] == "１":
                one_idx = k
                break
        if one_idx is not None and one_idx + 1 < len(chars) and chars[one_idx]["y"] == chars[one_idx+1]["y"]:
            one_adv = chars[one_idx+1]["x"] - chars[one_idx]["x"]
            print(f"  ★ '１' advance = {one_adv:.3f}pt")
        # L1 line break point
        l1 = [c for c in chars if c["line"] == 1]
        l1_text = "".join(c["ch"] for c in l1).rstrip("\r\n\x07")
        print(f"  L1 fits {len(l1)} chars: {l1_text!r}")
    finally:
        doc.Close(SaveChanges=False)

word.Quit()

print(f"\n=== Summary ===")
for label, adv in results.items():
    print(f"  '．' advance with {label}: {adv:.3f}pt")
if 'compressPunc' in results and 'doNotCompress' in results:
    diff = results['doNotCompress'] - results['compressPunc']
    print(f"  Δ (doNotCompress - compressPunc) = {diff:+.3f}pt")
    if abs(diff) < 0.5:
        print(f"  → compressPunctuation does NOT gate '．' advance — intrinsic to MS Mincho")
    else:
        print(f"  → compressPunctuation DOES gate '．' advance — fix in layout/font path")
