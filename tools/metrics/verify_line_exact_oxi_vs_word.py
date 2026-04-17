"""Compare Oxi vs Word at line=exact boundaries.

If Oxi uses lineB (bug) instead of lineA (correct), Oxi's measured
delta at the boundary = Word's delta + (lineB - lineA) scaled by lines.

Simple check: per-boundary Oxi_delta vs Word_delta. Should differ by
exactly (lineB - lineA)/20 * n_A_lines.
"""
import os, sys, re, json, subprocess, time, zipfile
from pathlib import Path
import win32com.client
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOC_DIR = r"tools/golden-test/documents/docx"
OXI_PNG_DIR = "pipeline_data/oxi_png"
OXI_RENDERER = str(Path(r"tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe").resolve())

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
            return (i+1, a, b, lines)
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
        pga = doc.Paragraphs(idx_a).Range.Information(3)
        pgb = doc.Paragraphs(idx_b).Range.Information(3)
        doc.Close(False)
        return (round(ya, 2), round(yb, 2), pga, pgb)
    finally:
        try: word.Quit()
        except: pass


def render_oxi(path, out_prefix):
    Path(out_prefix).parent.mkdir(parents=True, exist_ok=True)
    dump = Path(out_prefix).parent / (Path(out_prefix).name + "_layout.json")
    trash = str(Path(out_prefix).parent / "_trash").replace("\\", "/")
    result = subprocess.run([OXI_RENDERER, str(Path(path).resolve()), trash, "150", f"--dump-layout={dump}"],
                            capture_output=True, timeout=60)
    if result.returncode != 0: return None
    return json.load(open(dump, encoding='utf-8'))


def find_oxi_para_y(dump, para_idx):
    """Find FIRST y of para_idx in Oxi dump."""
    for pg in dump['pages']:
        for e in pg['elements']:
            if e.get('type') == 'text' and e.get('para_idx') == para_idx:
                return e.get('y')
    return None


def main():
    print(f"{'doc':45} {'idx':>5} {'lineA':>6} {'lineB':>6} {'W_A':>7} {'W_B':>7} {'W_d':>7} {'O_A':>7} {'O_B':>7} {'O_d':>7} {'Δ':>7}")
    for name in TARGETS:
        path = os.path.join(DOC_DIR, name)
        if not os.path.exists(path): continue
        b = find_first_boundary(path)
        if not b: continue
        idx_a, lineA, lineB, _ = b
        w = measure_word(path, idx_a, idx_a + 1)
        if not w: continue
        ya, yb, pga, pgb = w
        # Check if A has multiple lines by examining range height
        dump = render_oxi(path, f"{OXI_PNG_DIR}/_{name[:10]}_test/p")
        if not dump:
            print(f"{name[:45]:45} oxi render failed"); continue
        # Oxi para_idx may differ from Word. In Oxi dump, para_idx is 0-indexed block.
        # Use the Oxi block corresponding to Word's idx_a. Approximation: subtract 1.
        oa = find_oxi_para_y(dump, idx_a - 1)
        ob = find_oxi_para_y(dump, idx_a)
        if oa is None or ob is None:
            print(f"{name[:45]:45} idx_a={idx_a} lineA={lineA}→{lineB} W={yb-ya:.2f} Oxi para not found"); continue
        wd = yb - ya
        od = ob - oa
        diff = od - wd
        print(f"{name[:45]:45} {idx_a:>5} {lineA:>6} {lineB:>6} {ya:>7.2f} {yb:>7.2f} {wd:>7.2f} {oa:>7.2f} {ob:>7.2f} {od:>7.2f} {diff:>+7.2f}")


if __name__ == "__main__":
    main()
