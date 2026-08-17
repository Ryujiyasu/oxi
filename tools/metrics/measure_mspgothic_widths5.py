# -*- coding: utf-8 -*-
"""Measure the MS PGothic advances the shipped table is still rounding.

S1169 replaced 543 entries of `com_tw_overrides.json` with the finer numbers
that already existed in `pipeline_data/MS_PGothic_com_widths.json`, and gained
+0.0063 SSIM on c7b923e5 for it. But that file covers 141 characters, of which
only 45 appear in the corpus at all: 1224 of the characters the MS PGothic
documents actually use are still carried at the 10tw (0.5pt) rounding, worth
about 0.06pt each.

Same setup as `measure_mspgothic_widths4.py`, which produced the file S1169
shipped: each character repeated on its own left-aligned paragraph, so no
autoSpaceDE gap and no justification touches the advances.

★With one difference that decides whether the run is worth anything:
`Information(5)` is quantised to 0.25pt (5tw), so reading ONE advance produces
values that are all multiples of 5tw -- coarser than the 2tw table already
shipped, and measurably worse against Word's own PDF (0.177pt mean error on
kana against the shipped 0.115pt). So the advance is read across a RUN of
`REPEAT` characters and divided, which averages the quantisation down to
0.25/(REPEAT-1) pt. The run has to fit on one line or the x positions wrap:
40 fullwidth characters at 10.5pt is 420pt inside a 468pt column.

    python measure_mspgothic_widths5.py                      # the corpus's missing chars
    python measure_mspgothic_widths5.py --sizes 10.5,12      # only these sizes
"""
import json
import os
import sys
import time

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
sys.path.insert(0, HERE)
os.environ.setdefault("PYTHONIOENCODING", "utf-8")
sys.stdout.reconfigure(encoding="utf-8")

FONT = "ＭＳ Ｐゴシック"      # ＭＳ Ｐゴシック
MISSING = os.path.join(REPO, "pipeline_data", "_pgw_snap", "missing_chars.json")
OUT = os.path.join(REPO, "pipeline_data", "_pgw_snap", "mspgothic_widths_v5.json")
# A paragraph per character costs one COM round trip per measurement, so keep
# the batch to what Word stays responsive on.
BATCH = 400


def measure(word, chars, size, repeat):
    """{codepoint: twips} for one size, measured a batch at a time."""
    out = {}
    for start in range(0, len(chars), BATCH):
        batch = chars[start:start + BATCH]
        doc = word.Documents.Add()
        time.sleep(1.0)
        ps = doc.Sections(1).PageSetup
        ps.LeftMargin = 72
        ps.RightMargin = 72
        rng = doc.Range(0, 0)
        rng.Text = "\r".join(ch * repeat for ch in batch)
        rng.Font.Name = FONT
        rng.Font.Size = size
        rng.ParagraphFormat.Alignment = 0
        rng.ParagraphFormat.SpaceAfter = 0
        rng.ParagraphFormat.SpaceBefore = 0
        time.sleep(1.5)
        pos = 0
        for ch in batch:
            try:
                x0 = doc.Range(pos, pos + 1).Information(5)
                xn = doc.Range(pos + repeat - 1, pos + repeat).Information(5)
                a = (xn - x0) / (repeat - 1)
                if 0 < a < 40:
                    out[ord(ch)] = round(a * 20, 2)
            except Exception:
                pass
            pos += repeat + 1                          # the run plus its \r
        doc.Close(False)
        print("    %s: %d/%d measured" % (size, len(out), len(chars)))
    return out


def main():
    sizes = ["10.5", "11", "12", "14"]
    # 40 fullwidth glyphs at 10.5pt span 420pt, inside the 468pt column; at 14pt
    # they would not fit, so the run shrinks with the size.
    repeat_for = {"10.5": 40, "11": 38, "12": 34, "14": 30}
    for a in sys.argv[1:]:
        if a.startswith("--sizes"):
            sizes = (a.split("=", 1)[1] if "=" in a else sys.argv[sys.argv.index(a) + 1]).split(",")
    chars = json.load(open(MISSING, encoding="utf-8"))
    # Word measures through cp932; anything outside it is not what these
    # documents contain anyway.
    keep = []
    for ch in chars:
        try:
            ch.encode("cp932")
            keep.append(ch)
        except Exception:
            pass
    print("measuring %d characters (%d skipped as non-cp932) at %s"
          % (len(keep), len(chars) - len(keep), ",".join(sizes)))

    import win32com.client
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    result = {}
    try:
        for size in sizes:
            print("  size %s ..." % size)
            rep = repeat_for.get(size, 30)
            result[size] = {str(k): v
                            for k, v in measure(word, keep, float(size), rep).items()}
    finally:
        try:
            word.Quit()
        except Exception:
            pass
    json.dump({"MS PGothic": result}, open(OUT, "w", encoding="utf-8"),
              ensure_ascii=False, indent=1)
    print("wrote %s  (%s)" % (OUT, {s: len(v) for s, v in result.items()}))


if __name__ == "__main__":
    main()
