"""Measure all FT_* repros + 5 baseline floating-table docs via Word COM.

For each floating table, record:
  - tblpY (pt)
  - vertAnchor / horzAnchor
  - table_top (Information(6))
  - table_page (Information(3))
  - table_left (Information(7))   [horz, optional]
  - First-row first-cell content top (proxy of inner-content y)
  - Anchor paragraph: top, page, text snippet
  - For body paragraphs around the floating table: first body para after,
    and last body para before (raw)

Then group by (PreKind / Doc) and compute slope of table_top vs tblpY across
the variants in the same group. Output JSON + console summary.

Output: pipeline_data/ft_slope_measurements.json
"""
import json
import re
import zipfile
from pathlib import Path
import sys

import win32com.client as w32

REPRO_DIR = Path(r"C:\Users\ryuji\oxi-1\tools\metrics\ft_slope_repro")
BASELINE_DOCS_DIR = Path(r"C:\Users\ryuji\oxi-1\tools\golden-test\documents\docx")

BASELINE_DOC_NAMES = [
    "ed025cbecffb_index-23.docx",
    "2ea81a8441cc_0025006-192.docx",
    "3a4f9fbe1a83_001620506.docx",
    "459f05f1e877_kyodokenkyuyoushiki01.docx",  # vertAnchor=page (control)
    "1ec1091177b1_006.docx",
    "e201249db062_tokumei_08_05.docx",
]

OUT = Path(r"C:\Users\ryuji\oxi-1\pipeline_data\ft_slope_measurements.json")


def parse_floating_xml_order(docx_path):
    """Parse document.xml and return list of dicts per <w:tbl> in document
    order: {is_floating, tblpY_pt, vertAnchor, horzAnchor, tblpX_pt}.
    """
    with zipfile.ZipFile(docx_path) as z:
        xml = z.read("word/document.xml").decode("utf-8")
    out = []
    for tm in re.finditer(r"<w:tbl(?:\s[^>]*)?>", xml):
        start = tm.end()
        end = xml.find("</w:tbl>", start)
        if end < 0:
            continue
        block = xml[start:end]
        tp = re.search(r"<w:tblPr\b[^>]*>(.*?)</w:tblPr>", block, re.S)
        is_floating = False
        tblpY_pt = tblpX_pt = None
        vertAnchor = horzAnchor = None
        if tp:
            tppr = re.search(r"<w:tblpPr\b([^/>]*)/?>", tp.group(1))
            if tppr:
                is_floating = True
                a = tppr.group(1)
                m = re.search(r'w:tblpY="(-?\d+)"', a)
                if m: tblpY_pt = int(m.group(1)) / 20.0
                m = re.search(r'w:tblpX="(-?\d+)"', a)
                if m: tblpX_pt = int(m.group(1)) / 20.0
                m = re.search(r'w:vertAnchor="([^"]*)"', a)
                if m: vertAnchor = m.group(1)
                m = re.search(r'w:horzAnchor="([^"]*)"', a)
                if m: horzAnchor = m.group(1)
        out.append({
            "is_floating": is_floating,
            "tblpY_pt": tblpY_pt,
            "tblpX_pt": tblpX_pt,
            "vertAnchor": vertAnchor,
            "horzAnchor": horzAnchor,
        })
    return out


def measure_doc(word, docx_path):
    info = {"file": docx_path.name, "tables": []}
    xml_order = parse_floating_xml_order(docx_path)
    info["xml_table_order"] = xml_order

    doc = word.Documents.Open(str(docx_path.resolve()), ReadOnly=True)
    try:
        for ti, tbl in enumerate(doc.Tables, start=1):
            try:
                tt = tbl.Range.Information(6)
                tp = tbl.Range.Information(3)
                tl = tbl.Range.Information(7)
            except Exception:
                continue

            anchor_y = anchor_pg = None
            anchor_text = None
            try:
                pre_range = doc.Range(0, tbl.Range.Start)
                if pre_range.Paragraphs.Count > 0:
                    last_pre = pre_range.Paragraphs(pre_range.Paragraphs.Count)
                    anchor_y = last_pre.Range.Information(6)
                    anchor_pg = last_pre.Range.Information(3)
                    anchor_text = (last_pre.Range.Text or "")[:40].replace("\r", "\\r").replace("\x07", "\\x07")
            except Exception:
                pass

            # First cell first paragraph y
            cell_y = None
            try:
                first_para = tbl.Cell(1, 1).Range.Paragraphs(1)
                cell_y = first_para.Range.Information(6)
            except Exception:
                pass

            tinfo = {
                "table_idx_word": ti,
                "table_top_pt": round(tt, 3) if tt is not None else None,
                "table_left_pt": round(tl, 3) if tl is not None else None,
                "table_page": tp,
                "anchor_para_top_pt": round(anchor_y, 3) if anchor_y is not None else None,
                "anchor_para_page": anchor_pg,
                "anchor_para_text": anchor_text,
                "first_cell_first_para_top_pt": round(cell_y, 3) if cell_y is not None else None,
                "rows": tbl.Rows.Count,
                "cols": tbl.Columns.Count,
            }
            if ti - 1 < len(xml_order):
                tinfo.update(xml_order[ti - 1])
            info["tables"].append(tinfo)
    finally:
        doc.Close(SaveChanges=0)
    return info


def main():
    targets = []
    if REPRO_DIR.exists():
        targets += sorted(REPRO_DIR.glob("FT_*.docx"))
    for n in BASELINE_DOC_NAMES:
        p = BASELINE_DOCS_DIR / n
        if p.exists():
            targets.append(p)
        else:
            print(f"  warn: missing {p.name}", file=sys.stderr)

    print(f"Measuring {len(targets)} docs...", file=sys.stderr)

    word = w32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    results = []
    try:
        for d in targets:
            try:
                info = measure_doc(word, d)
                results.append(info)
            except Exception as e:
                results.append({"file": d.name, "error": str(e)})
            print(f"  done {d.name}", file=sys.stderr)
    finally:
        word.Quit()

    OUT.parent.mkdir(parents=True, exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, ensure_ascii=False)

    # Summary table
    print()
    print(f"{'doc':50s} {'va':>5} {'tblpY':>7} {'a_top':>7} {'t_top':>7} {'Δ':>7} {'page':>4} {'rows':>4}")
    print("-" * 110)
    for r in results:
        if "error" in r:
            print(f"  {r['file']:48s}  ERROR: {r['error']}")
            continue
        for t in r.get("tables", []):
            if not t.get("is_floating"):
                continue
            va = (t.get("vertAnchor") or "?")[:5]
            tblpY = t.get("tblpY_pt")
            apt = t.get("anchor_para_top_pt")
            tt = t.get("table_top_pt")
            tp = t.get("table_page") or 0
            rows = t.get("rows") or 0
            delta = (tt - apt) if (tt is not None and apt is not None) else None
            print(f"  {r['file'][:48]:48s} {va:>5s}"
                  f" {(f'{tblpY:6.2f}' if tblpY is not None else '   -  '):>7}"
                  f" {(f'{apt:6.2f}' if apt is not None else '   -  '):>7}"
                  f" {(f'{tt:6.2f}' if tt is not None else '   -  '):>7}"
                  f" {(f'{delta:+6.2f}' if delta is not None else '   -  '):>7}"
                  f" {tp:>4d} {rows:>4d}")

    # Slope analysis: group FT_<PreKind>_Y* and compute slope
    print()
    print("=== Slope analysis (FT_<PreKind>_Y*) ===")
    groups = {}
    for r in results:
        n = r["file"]
        m = re.match(r"FT_(\w+?)_Y(\d+)\.docx", n)
        if not m:
            continue
        pk, ytw = m.group(1), int(m.group(2))
        for t in r.get("tables", []):
            if not t.get("is_floating"):
                continue
            groups.setdefault(pk, []).append((ytw / 20.0, t.get("table_top_pt"), t.get("anchor_para_top_pt")))
            break  # one floating tbl per repro
    for pk, pts in sorted(groups.items()):
        pts.sort()
        ys, tts, apts = zip(*pts)
        if len(ys) >= 2 and tts[-1] is not None and tts[0] is not None:
            slope = (tts[-1] - tts[0]) / (ys[-1] - ys[0]) if (ys[-1] != ys[0]) else 0
            print(f"  {pk:10s} tblpY {ys}  table_top {tts}  anchor_top {apts}")
            print(f"             ==> slope = {slope:.3f}")


if __name__ == "__main__":
    main()
