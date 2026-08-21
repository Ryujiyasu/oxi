# -*- coding: utf-8 -*-
# Top up the font default-row-height table with every (face, size) pair the
# corpus styles name that the sweep has not measured. The table is keyed by
# the REQUESTED face: when the face is not installed, what this measurement
# records is Excel's own substitution on this machine — exactly what Excel
# does when it opens the document here, which is what the gate compares.
import glob
import json
import re
import sys
import zipfile
import win32com.client

SWEEP = r"pipeline_data\com_measurements\xlsx_row_height_sweep.json"


def main():
    sweep = json.load(open(SWEEP, encoding="utf-8"))
    have = {(r["face"], float(r["size"])) for r in sweep}
    missing = set()
    for p in glob.glob(r"tools\golden-test\documents\xlsx\*.xlsx"):
        z = zipfile.ZipFile(p)
        try:
            st = z.read("xl/styles.xml").decode("utf-8")
        except KeyError:
            continue
        for f in re.findall(r"<font>.*?</font>|<font/>",
                            re.search(r"<fonts.*?</fonts>", st, re.S).group(0)):
            n = re.search(r'name val="([^"]+)"', f)
            s = re.search(r'sz val="([0-9.]+)"', f)
            if n and s and (n.group(1), float(s.group(1))) not in have:
                missing.add((n.group(1), float(s.group(1))))
    print("%d pairs to measure" % len(missing))

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    added = 0
    try:
        wb = excel.Workbooks.Add()
        ws = wb.Worksheets(1)
        font = wb.Styles("Normal").Font
        for face, size in sorted(missing):
            try:
                font.Name = face
                font.Size = size
            except Exception as e:
                print("  %-24s %5s SET FAILED %s" % (face, size, str(e)[:40]))
                continue
            entry = {
                "face": face,
                "size": size,
                "applied_name": font.Name,
                "standard_height_pt": ws.StandardHeight,
                "standard_width_chars": ws.StandardWidth,
                "height_px": ws.StandardHeight / 0.75,
            }
            sweep.append(entry)
            added += 1
            print("  %-24s %5s -> %6.2fpt %5.1fpx (applied %s)" % (
                face, size, entry["standard_height_pt"], entry["height_px"],
                entry["applied_name"]))
        wb.Close(False)
    finally:
        excel.Quit()
    with open(SWEEP, "w", encoding="utf-8") as f:
        json.dump(sweep, f, ensure_ascii=False, indent=1)
    print("added %d; sweep now %d entries" % (added, len(sweep)))


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
