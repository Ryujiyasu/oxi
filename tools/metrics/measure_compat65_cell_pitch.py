"""Verify whether Word grid-snaps intra-cell line pitch in docs with
adjustLineHeightInTable=True.

Hypothesis (per memories d77a_cell_grid_pitch + d77a_cell_pitch_mismatch):
Word DOES grid-snap intra-cell lines to docGrid pitch even when
adjustLineHeightInTable is set. If confirmed across 3+ docs, fixing Oxi's
in-table line height to grid-snap is safe for all 35 compat65 docs.

Method: for each doc, pick the first table, find a multi-line paragraph in
cell (1,1), measure char Y positions, compute pitch = median y-diff.
Compare to doc's docGrid linePitch (pt).
"""
import os, sys, json, time
import win32com.client
import zipfile
import re

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOCS = [
    "04b88e7e0b25_index-19.docx",
    "2ea81a8441cc_0025006-192.docx",
    "6514f214e482_tokumei_08_01-2.docx",  # compat65 + tables
]

DOC_DIR = r"tools\golden-test\documents\docx"


def docgrid_pitch_pt(path: str) -> float:
    with zipfile.ZipFile(path) as zf:
        xml = zf.read("word/document.xml").decode("utf-8")
    m = re.search(r'<w:docGrid[^>]*w:linePitch="(\d+)"', xml)
    if m:
        return int(m.group(1)) / 20.0
    return 0.0


def measure_cell_pitch(doc, tbl_idx: int):
    tbl = doc.Tables(tbl_idx)
    cell = tbl.Cell(1, 1)
    # Find the first paragraph with ≥3 lines (so pitch can be measured)
    # Scan ALL cells in table, pick paragraph with most lines
    best = None  # (ys, ri, ci, pi)
    for ri in range(1, tbl.Rows.Count + 1):
        for ci in range(1, tbl.Columns.Count + 1):
            try:
                cell = tbl.Cell(ri, ci)
            except Exception:
                continue
            for pi, p in enumerate(cell.Range.Paragraphs, 1):
                pr = p.Range
                n = pr.Characters.Count
                if n < 20:
                    continue
                ys = set()
                step = max(1, n // 80)
                for i in range(1, n + 1, step):
                    try:
                        ys.add(round(pr.Characters(i).Information(6), 1))
                    except Exception:
                        pass
                if len(ys) >= 2 and (best is None or len(ys) > len(best[0])):
                    best = (sorted(ys), ri, ci, pi)
    if best is None:
        return None, -1
    return (best[0], None), best[3]


def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    report = []
    try:
        for name in DOCS:
            path = os.path.abspath(os.path.join(DOC_DIR, name))
            grid_pt = docgrid_pitch_pt(path)
            result = {"doc": name, "grid_pitch_pt": grid_pt}
            doc = word.Documents.Open(path, ReadOnly=True)
            time.sleep(0.3)
            doc.Repaginate()
            try:
                ntbls = doc.Tables.Count
                result["n_tables"] = ntbls
                # Try first 3 tables to find a multi-line para
                found = False
                for ti in range(1, min(ntbls, 4) + 1):
                    chosen, pi = measure_cell_pitch(doc, ti)
                    if chosen:
                        ys, _p = chosen
                        diffs = [round(ys[i + 1] - ys[i], 2) for i in range(len(ys) - 1)]
                        # median-ish by taking middle values
                        diffs_sorted = sorted(diffs)
                        # filter out 0 (same-line samples)
                        nz = [d for d in diffs_sorted if d > 0.1]
                        if nz:
                            pitch = nz[len(nz) // 2]
                            result["tbl_idx"] = ti
                            result["para_idx"] = pi
                            result["n_lines"] = len(ys)
                            result["line_ys"] = ys
                            result["pitch_diffs"] = diffs
                            result["measured_pitch_pt"] = pitch
                            result["grid_snapped"] = abs(pitch - grid_pt) < 0.5
                            found = True
                            break
                if not found:
                    result["note"] = "no multi-line para in first tables"
            finally:
                doc.Close(False)

            report.append(result)
            pitch_s = result.get("measured_pitch_pt", "?")
            snap_s = result.get("grid_snapped", "?")
            print(
                f"{name}: grid={grid_pt}pt measured_cell_pitch={pitch_s}pt "
                f"grid_snapped={snap_s}",
                file=sys.stderr,
            )
    finally:
        word.Quit()

    out = "pipeline_data/compat65_cell_pitch.json"
    with open(out, "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=2)
    print(f"\n[OK] {out}")


if __name__ == "__main__":
    main()
