"""Measure first paragraph Y for all linesAndChars docs to find Y0 formula."""
import win32com.client, time, sys, os, glob, zipfile, re
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOCX_DIR = os.path.abspath("tools/golden-test/documents/docx")
WD_Y = 6

def get_doc_grid(docx_path):
    try:
        with zipfile.ZipFile(docx_path) as z:
            doc = z.read('word/document.xml').decode('utf-8')
        m = re.search(r'docGrid[^/]*/|docGrid[^>]*>', doc)
        if m:
            s = m.group(0)
            t = re.search(r'type="([^"]*)"', s)
            lp = re.search(r'linePitch="(\d+)"', s)
            cp = re.search(r'charSpace="([^"]*)"', s)
            return (t.group(1) if t else None,
                    int(lp.group(1)) if lp else None,
                    cp.group(1) if cp else None)
    except:
        pass
    return (None, None, None)

def get_top_margin(docx_path):
    try:
        with zipfile.ZipFile(docx_path) as z:
            doc = z.read('word/document.xml').decode('utf-8')
        m = re.search(r'pgMar[^/]*/', doc)
        if m:
            t = re.search(r'w:top="(\d+)"', m.group(0))
            if t: return int(t.group(1)) / 20.0
    except:
        pass
    return None

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False

results = []
for docx in sorted(glob.glob(os.path.join(DOCX_DIR, "*.docx"))):
    grid_type, line_pitch, char_space = get_doc_grid(docx)
    if grid_type != "linesAndChars":
        continue
    top_margin = get_top_margin(docx)
    if top_margin is None:
        continue

    try:
        doc = word.Documents.Open(os.path.abspath(docx), ReadOnly=True)
        time.sleep(0.3)
        p1 = doc.Paragraphs(1)
        y1 = p1.Range.Information(WD_Y)
        fs1 = p1.Range.Font.Size
        ls1 = p1.Format.LineSpacing
        doc.Close(SaveChanges=False)

        name = os.path.basename(docx)
        pitch_pt = line_pitch / 20.0 if line_pitch else 0
        delta = y1 - top_margin
        results.append((name, top_margin, pitch_pt, y1, delta, fs1, ls1))
        print(f"  {name}: top={top_margin:.1f} pitch={pitch_pt:.1f} P1_y={y1:.1f} delta={delta:.2f} fs={fs1} ls={ls1:.1f}")
    except Exception as e:
        print(f"  ERR {os.path.basename(docx)}: {e}")

word.Quit()
print(f"\n{len(results)} linesAndChars documents measured")
