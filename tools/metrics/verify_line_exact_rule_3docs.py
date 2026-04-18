"""Verify "B→C advance uses B's line" rule on 3+ real docs.

For each doc, pick the first boundary (para A with lineA exact → para B
with lineB exact where lineA != lineB), measure Word's Y positions,
check whether A→B advance equals lineA/20 pt (B's rule) vs lineB/20 pt
(C's rule — the bug).

In the repro, "B" is the empty/middle para. In real docs we measure
para N (A) and para N+1 (B) directly — the "B→C" in the bug terminology
is the N → N+1 transition.

Rule check: delta_y(N, N+1) should equal lineA / 20 (using A's line value).
If matches → rule confirmed. If equals lineB / 20 → different rule.
"""
import os, sys, re, subprocess, time, zipfile
from pathlib import Path
import win32com.client
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOC_DIR = r"tools/golden-test/documents/docx"

# Target 3 docs with clear boundaries
TARGETS = [
    "2ea81a8441cc_0025006-192.docx",
    "29dc6e8943fe_order_01.docx",
    "6514f214e482_tokumei_08_01-2.docx",
]


def find_first_boundary(path):
    with zipfile.ZipFile(path) as zf:
        xml = zf.read("word/document.xml").decode("utf-8")
    p_blocks = re.findall(r'<w:p[ >](?:(?!<w:p[ >]).)*?</w:p>', xml, re.DOTALL)
    lines = []
    for b in p_blocks:
        m = re.search(r'<w:spacing[^/]*w:line="(\d+)"[^/]*w:lineRule="exact"', b)
        if not m:
            m = re.search(r'<w:spacing[^/]*w:lineRule="exact"[^/]*w:line="(\d+)"', b)
        lines.append(int(m.group(1)) if m else None)
    for i in range(len(lines) - 1):
        a, b = lines[i], lines[i+1]
        if a is not None and b is not None and a != b:
            return (i+1, a, b)  # para_idx (1-indexed), lineA, lineB
    return None


def measure_word(path, idx_a, idx_b):
    subprocess.run(["taskkill", "/F", "/IM", "WINWORD.EXE"], capture_output=True, timeout=5)
    time.sleep(0.5)
    word = win32com.client.DispatchEx("Word.Application")
    try:
        try: word.Visible = False
        except: pass
        try: word.DisplayAlerts = False
        except: pass
        doc = word.Documents.Open(str(Path(path).resolve()), ReadOnly=True)
        time.sleep(0.3)
        doc.Repaginate()
        ya = doc.Paragraphs(idx_a).Range.Information(6)
        yb = doc.Paragraphs(idx_b).Range.Information(6)
        pg_a = doc.Paragraphs(idx_a).Range.Information(3)
        pg_b = doc.Paragraphs(idx_b).Range.Information(3)
        doc.Close(False)
        return (round(ya, 2), round(yb, 2), pg_a, pg_b)
    finally:
        try: word.Quit()
        except: pass


def main():
    print(f"{'doc':45} {'para':>6} {'lineA':>6} {'lineB':>6} {'expectA':>8} {'expectB':>8} {'measured':>9} {'verdict':>15}")
    for name in TARGETS:
        path = os.path.join(DOC_DIR, name)
        if not os.path.exists(path):
            print(f"{name}: missing"); continue
        b = find_first_boundary(path)
        if not b:
            print(f"{name}: no boundary found"); continue
        idx_a, lineA, lineB = b
        try:
            ya, yb, pga, pgb = measure_word(path, idx_a, idx_a + 1)
            if pga != pgb:
                print(f"{name[:45]:45} para {idx_a}→{idx_a+1}: spans pages {pga}→{pgb}, skip"); continue
            delta = yb - ya
            expect_a = lineA / 20.0
            expect_b = lineB / 20.0
            if abs(delta - expect_a) < 0.2:
                verdict = f"MATCHES A ({expect_a:.2f})"
            elif abs(delta - expect_b) < 0.2:
                verdict = f"matches B ({expect_b:.2f})"
            else:
                verdict = f"neither ({delta:.2f})"
            print(f"{name[:45]:45} {idx_a:>6} {lineA:>6} {lineB:>6} {expect_a:>8.2f} {expect_b:>8.2f} {delta:>9.2f} {verdict:>15}")
        except Exception as e:
            print(f"{name}: err {e}")


if __name__ == "__main__":
    main()
