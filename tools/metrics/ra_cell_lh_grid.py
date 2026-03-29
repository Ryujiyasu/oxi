"""
Ra: テーブルセル内行高さのgrid snap vs ls報告値の関係を完全確定
- ls=12(Single)報告なのに実際gap=18の理由
- gridなし文書でのセル内行高さ
- adjustLineHeightInTable の影響
- CJKフォントでのセル内行高さ
"""
import win32com.client, json, os, tempfile
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from lxml import etree

TEMPLATE = os.path.join(os.path.dirname(__file__), 'ja_gov_template.docx')

word = win32com.client.DispatchEx('Word.Application')
word.Visible = False
word.DisplayAlerts = False

results = []


def make_table_doc(grid_type=None, grid_pitch=None, font_name="Calibri", font_size=11):
    d = Document(TEMPLATE)
    sec = d.sections[0]
    sectPr = sec._sectPr
    # Remove existing grid
    for dg in sectPr.findall(qn('w:docGrid')):
        sectPr.remove(dg)
    if grid_type:
        dg = etree.SubElement(sectPr, qn('w:docGrid'))
        dg.set(qn('w:type'), grid_type)
        dg.set(qn('w:linePitch'), str(grid_pitch))
    # If no grid_type, no docGrid at all

    for p in d.paragraphs:
        p._element.getparent().remove(p._element)

    # Add table with 3 rows
    from docx.table import Table
    tbl_el = d.element.body.makeelement(qn('w:tbl'), {})
    d.element.body.append(tbl_el)
    tblPr = etree.SubElement(tbl_el, qn('w:tblPr'))
    tblBorders = etree.SubElement(tblPr, qn('w:tblBorders'))
    for side in ['top', 'bottom', 'left', 'right', 'insideH', 'insideV']:
        b = etree.SubElement(tblBorders, qn(f'w:{side}'))
        b.set(qn('w:val'), 'single')
        b.set(qn('w:sz'), '4')
        b.set(qn('w:color'), '000000')

    for r in range(3):
        tr = etree.SubElement(tbl_el, qn('w:tr'))
        for c in range(2):
            tc = etree.SubElement(tr, qn('w:tc'))
            p = etree.SubElement(tc, qn('w:p'))
            pPr = etree.SubElement(p, qn('w:pPr'))
            spacing = etree.SubElement(pPr, qn('w:spacing'))
            spacing.set(qn('w:before'), '0')
            spacing.set(qn('w:after'), '0')
            run = etree.SubElement(p, qn('w:r'))
            rPr = etree.SubElement(run, qn('w:rPr'))
            rFonts = etree.SubElement(rPr, qn('w:rFonts'))
            rFonts.set(qn('w:ascii'), font_name)
            rFonts.set(qn('w:eastAsia'), font_name)
            sz = etree.SubElement(rPr, qn('w:sz'))
            sz.set(qn('w:val'), str(int(font_size * 2)))
            t = etree.SubElement(run, qn('w:t'))
            t.text = f"R{r+1}C{c+1}"

    return d


def measure_table(doc_path, label):
    doc = word.Documents.Open(doc_path)
    try:
        data = {"label": label, "rows": []}

        tbl = doc.Tables(1)
        for r in range(1, tbl.Rows.Count + 1):
            cell = tbl.Cell(r, 1)
            para = cell.Range.Paragraphs(1)
            y = para.Range.Information(6)
            ls = para.Format.LineSpacing
            ls_rule = para.Format.LineSpacingRule
            data["rows"].append({
                "row": r,
                "text_y": round(y, 4),
                "line_spacing": round(ls, 4),
                "ls_rule": ls_rule,
            })

        for i in range(1, len(data["rows"])):
            data["rows"][i]["gap"] = round(
                data["rows"][i]["text_y"] - data["rows"][i-1]["text_y"], 4)

        return data
    finally:
        doc.Close(False)


try:
    configs = [
        ("grid_lines_360", "lines", 360, "Calibri", 11),
        ("grid_lines_300", "lines", 300, "Calibri", 11),
        ("no_grid", None, None, "Calibri", 11),
        ("grid_lines_360_msgothic", "lines", 360, "MS Gothic", 10.5),
        ("no_grid_msgothic", None, None, "MS Gothic", 10.5),
        ("grid_lines_360_9pt", "lines", 360, "Calibri", 9),
        ("no_grid_9pt", None, None, "Calibri", 9),
    ]

    for name, gtype, gpitch, font, fs in configs:
        d = make_table_doc(grid_type=gtype, grid_pitch=gpitch, font_name=font, font_size=fs)
        tmp = os.path.join(tempfile.gettempdir(), f"ra_cell_lh_{name}.docx")
        d.save(tmp)
        data = measure_table(tmp, name)
        results.append(data)
        os.unlink(tmp)

        print(f"\n=== {name} ({font} {fs}pt, grid={gtype}/{gpitch}) ===")
        for r in data["rows"]:
            gap = r.get("gap", "-")
            print(f"  R{r['row']}: y={r['text_y']}, ls={r['line_spacing']}(rule={r['ls_rule']}), gap={gap}")

finally:
    word.Quit()

out_path = os.path.join(os.path.dirname(__file__), 'output', 'ra_cell_lh_grid.json')
os.makedirs(os.path.dirname(out_path), exist_ok=True)
with open(out_path, 'w', encoding='utf-8') as f:
    json.dump(results, f, indent=2, ensure_ascii=False)
print(f"\nResults saved to {out_path}")

# Analysis
print("\n\n=== ANALYSIS ===")
for data in results:
    gaps = [r["gap"] for r in data["rows"] if "gap" in r]
    avg_gap = sum(gaps) / len(gaps) if gaps else 0
    ls = data["rows"][0]["line_spacing"]
    print(f"{data['label']}: ls_reported={ls}, avg_gap={avg_gap:.2f}")
