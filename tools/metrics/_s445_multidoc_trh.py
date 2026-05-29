"""S445: across negative-drift form docs, measure Word's actual row pitch vs the
trHeight value for atLeast rows, to confirm the bump formula (3-doc rule).

For each doc: group cells by RowIndex, get each row's topmost cell text Y
(collapsed start, R30). Pitch(row k) = topY(k+1)-topY(k). Compare consecutive
rows that share the SAME trHeight value (read from XML order) -> clean pitch.
We can't easily map XML trHeight to COM row, so instead report the modal pitch
and compare to the doc's trHeight sample values.
"""
import sys, os, json, statistics, zipfile, re
import win32com.client as win32
sys.stdout.reconfigure(encoding="utf-8")

DOCXDIR = r"c:\Users\ryuji\oxi-main\tools\golden-test\documents\docx"
DOCS = {
    "7ead52b63f0e": "7ead52b63f0e_000067058.docx",
    "6514f214e482": None,
    "d4d126dfe1d9": None,
    "de6e32b5960b": None,
}

def find_file(stem):
    for f in os.listdir(DOCXDIR):
        if f.startswith(stem):
            return f
    return None

def trheights(path):
    with zipfile.ZipFile(path) as z:
        x = z.read("word/document.xml").decode("utf-8")
    return [(int(v), rule or "(none)") for v, rule in
            re.findall(r'<w:trHeight w:val="(\d+)"(?:\s+w:hRule="(\w+)")?', x)]

word = win32.gencache.EnsureDispatch("Word.Application")
word.Visible = False
try:
    for doc_id, fn in DOCS.items():
        if fn is None:
            fn = find_file(doc_id)
        if not fn:
            print(f"{doc_id}: file not found"); continue
        path = os.path.join(DOCXDIR, fn)
        trh = trheights(path)
        doc = word.Documents.Open(path, ReadOnly=True)
        print(f"\n=== {doc_id} ({fn}) ===")
        print(f"  trHeights (val_tw, rule), first 12: {trh[:12]}")
        try:
            for ti in range(1, doc.Tables.Count + 1):
                tbl = doc.Tables(ti)
                cells = tbl.Range.Cells
                rowy = {}
                for ci in range(1, cells.Count + 1):
                    c = cells(ci)
                    ri = c.RowIndex
                    sr = doc.Range(c.Range.Start, c.Range.Start)
                    y = float(sr.Information(6))
                    rowy.setdefault(ri, []).append(y)
                tops = [(k, min(v)) for k, v in sorted(rowy.items())]
                pitches = [round(b[1] - a[1], 2) for a, b in zip(tops, tops[1:])]
                # modal pitch
                from collections import Counter
                cnt = Counter(pitches)
                print(f"  table{ti}: nrows~{len(tops)} pitch modes: {cnt.most_common(6)}")
        finally:
            doc.Close(False)
finally:
    word.Quit()
