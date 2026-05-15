"""仮説 5 を minimal_a1d6_v2 (T5 R1-R3 のみ) で検証.

Word が cell first-block sb を行高さ計算に含めて docGrid linePitch にスナップするか:
- docGrid linePitch=292tw = 14.6pt
- R1 cells: sb=4.35pt (suppressed), line=11pt → row pitch = ceil(15.35/14.6)*14.6 = 29.20pt (if hypothesis holds for R1→R2)
- R2 cells: sb=7.30pt (suppressed), line=14pt → row pitch = ceil(21.30/14.6)*14.6 = 29.20pt
- R3 cells: sb=7.30pt (suppressed), line=14pt → row pitch = ceil(21.30/14.6)*14.6 = 29.20pt

Expected Word y for cells (R1→R2 と R2→R3):
- R1.C2 ("名称"): y0 (some baseline near table top)
- R2.C2: y0 + 29.20pt
- R3.C2: y0 + 29.20pt × 2

a1d6 original では R2 y=93.75 - R1 y=67.50 = 26.25pt (仮説 29.20 と違う)
R3 y=123.00 - R2 y=93.75 = 29.25pt (仮説と一致)

R1→R2 で 26.25pt がなぜ 29.20 と一致しないか、minimal v2 で確認。
"""
import os, sys, json
import win32com.client as wc
sys.stdout.reconfigure(encoding='utf-8')

def info6(d, pos):
    try: return float(d.Range(pos, pos).Information(6))
    except: return None

def measure_doc(path):
    word = wc.Dispatch('Word.Application')
    word.Visible = False
    doc = word.Documents.Open(os.path.abspath(path), ReadOnly=True)
    try:
        # T5 (= tables[5] 1-based) の cell content y を計測
        tables = doc.Tables
        if tables.Count < 5:
            print(f'{path}: only {tables.Count} tables, T5 not present')
            return
        T5 = tables(5)
        print(f'\n=== {os.path.basename(path)} ===')
        print(f'T5: rows={T5.Rows.Count} cols={T5.Columns.Count}')
        cells = T5.Range.Cells
        for ci in range(1, cells.Count + 1):
            try:
                c = cells(ci)
                ri = c.RowIndex
                col = c.ColumnIndex
                cs_y = info6(doc, c.Range.Start)
                cell_paras = c.Range.Paragraphs
                p1 = cell_paras(1)
                p1_y = info6(doc, p1.Range.Start)
                p1_sb = float(p1.Format.SpaceBefore or 0)
                p1_ls = float(p1.Format.LineSpacing or 0)
                p1_lr = int(p1.Format.LineSpacingRule or 0)
                text = (p1.Range.Text or '').replace('\r','').strip()[:30]
                print(f'  cell({ri},{col}) cs_y={cs_y:>7.2f} p1_y={p1_y:>7.2f} sb={p1_sb:>5.2f} ls={p1_ls:>5.1f} lr={p1_lr} text={text!r}')
            except Exception as e:
                pass
    finally:
        doc.Close(SaveChanges=False)
        word.Quit()

def main():
    measure_doc('c:/tmp/minimal_a1d6_v2.docx')

if __name__ == '__main__':
    main()
