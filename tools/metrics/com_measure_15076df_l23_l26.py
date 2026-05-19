"""Extend COM measurement to L23/L26/L50 paragraphs of 15076df.

Same diff pattern (Oxi 1-char-short). Question: is the cause the same as L12
(narrow yakumono '．','、'  not detected by Oxi)?
"""
import os
import sys
import io
import win32com.client

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
DOCX = os.path.normpath(os.path.join(
    REPO, "tools/golden-test/documents/docx/15076df085f5_tokumei_08_09.docx"))

# Substrings to match the target paragraphs
targets = [
    ("L23-(3)", "（３）統計又は統計的研究の成果の概要"),
    ("L26-(4)", "（４）匿名データを利用して行った研究の成果"),
    ("L50-3.", "３．匿名データ利用後の措置状況"),
]

word = win32com.client.gencache.EnsureDispatch("Word.Application")
word.Visible = False
doc = word.Documents.Open(DOCX, ReadOnly=True)
try:
    for label, needle in targets:
        target_para = None
        for i in range(1, doc.Paragraphs.Count + 1):
            p = doc.Paragraphs(i)
            if needle in p.Range.Text:
                target_para = p
                break
        if target_para is None:
            print(f"\n=== {label}: NOT FOUND ===")
            continue
        p = target_para
        txt = p.Range.Text
        print(f"\n=== {label}: idx={i}, len={p.Range.End - p.Range.Start} ===")
        print(f"text: {txt[:80]!r}")

        wdHoriz, wdVert = 5, 6
        rng_start = p.Range.Start
        rng_end = p.Range.End
        prev_x = prev_y = None
        line_num = 1
        chars = []
        for j in range(rng_start, rng_end):
            r = doc.Range(j, j)
            x = r.Information(wdHoriz)
            y = r.Information(wdVert)
            nr = doc.Range(j, j + 1)
            ch = nr.Text
            if prev_x is not None and y != prev_y:
                line_num += 1
            chars.append({"i": j - rng_start, "x": x, "y": y, "ch": ch, "line": line_num})
            prev_x, prev_y = x, y

        # Group by line, find last char per line
        lines = {}
        for c in chars:
            lines.setdefault(c["line"], []).append(c)
        for lno, lchars in sorted(lines.items()):
            text = "".join(c["ch"] for c in lchars).rstrip("\r\n\x07")
            x0 = lchars[0]["x"]
            x1 = lchars[-1]["x"]
            # find advance widths of last 5 chars to spot narrow yakumono
            advs = []
            for k in range(max(1, len(lchars) - 6), len(lchars)):
                if lchars[k]["line"] == lchars[k - 1]["line"]:
                    adv = lchars[k]["x"] - lchars[k - 1]["x"]
                    advs.append(f"{lchars[k - 1]['ch']}={adv:.2f}")
            advs_str = " ".join(advs)
            print(f"  L{lno}: x={x0:.2f}..{x1:.2f}, {len(lchars)} chars, last-adv: {advs_str}")
            print(f"        text: {text!r}")
finally:
    doc.Close(SaveChanges=False)
    word.Quit()
