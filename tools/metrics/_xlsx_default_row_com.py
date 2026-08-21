# -*- coding: utf-8 -*-
# Ground truth for the default row height question: for every corpus workbook,
# what the sheet DECLARES (defaultRowHeight / dyDescent / customHeight, Normal
# font) against what Excel actually DRAWS (StandardHeight, and the Height of a
# row far past the used range, which no <row> entry can have touched).
#
# Candidate under test (2026-08-21): drawn default px =
#   ceil((defaultRowHeight + (customHeight ? 0 : dyDescent)) / 0.75)
# against SX13's font-derived story, which the Normal-rewrite instrument
# contradicts (Meiryo UI 11 rewrites to 15.75, not 224e's 18.75).
import glob
import json
import re
import sys
import zipfile
import win32com.client


def first_sheet_xml(z):
    # Worksheets(1) is the first <sheet> in workbook.xml order, resolved
    # through the relationships — sheetN.xml file names do not carry order.
    wbxml = z.read("xl/workbook.xml").decode("utf-8")
    rid = re.search(r'<sheet [^>]*r:id="([^"]+)"', wbxml).group(1)
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    m = re.search(r'Id="%s"[^>]*Target="([^"]+)"' % rid, rels)
    if not m:
        m = re.search(r'Target="([^"]+)"[^>]*Id="%s"' % rid, rels)
    target = m.group(1).replace("../", "").lstrip("/")
    if not target.startswith("xl/"):
        target = "xl/" + target
    return z.read(target).decode("utf-8")


def declared(path):
    z = zipfile.ZipFile(path)
    styles = z.read("xl/styles.xml").decode("utf-8")
    f = re.search(r"<font>(.*?)</font>", styles, re.S).group(1)
    name = re.search(r'name val="([^"]+)"', f)
    sz = re.search(r'sz val="([^"]+)"', f)
    sheet = first_sheet_xml(z)
    m = re.search(r"<sheetFormatPr[^>]*>", sheet)
    out = {
        "font": name.group(1) if name else None,
        "font_size": float(sz.group(1)) if sz else None,
        "sheetFormatPr": m.group(0) if m else None,
        "defaultRowHeight": None,
        "dyDescent": None,
        "customHeight": False,
    }
    if m:
        drh = re.search(r'defaultRowHeight="([^"]+)"', m.group(0))
        dyd = re.search(r'dyDescent="([^"]+)"', m.group(0))
        ch = re.search(r'customHeight="(1|true)"', m.group(0))
        out["defaultRowHeight"] = float(drh.group(1)) if drh else None
        out["dyDescent"] = float(dyd.group(1)) if dyd else None
        out["customHeight"] = bool(ch)
    return out


def main():
    docs = sorted(glob.glob(r"tools\golden-test\documents\xlsx\*.xlsx"))
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    rows = []
    try:
        for path in docs:
            d = declared(path)
            try:
                wb = excel.Workbooks.Open(
                    r"C:\Users\ryuji\oxi-main" + "\\" + path, 0, False)
            except Exception as e:
                print(path, "OPEN FAILED", str(e)[:80])
                continue
            ws = wb.Worksheets(1)
            d["doc"] = path.split("\\")[-1][:24]
            # The TRUE Normal font, from Excel itself — the first <font> in
            # styles.xml is a heuristic that openpyxl-authored files break.
            nf = wb.Styles("Normal").Font
            d["normal_font"] = nf.Name
            d["normal_size"] = nf.Size
            d["standard_height_pt"] = ws.StandardHeight
            # A row the file cannot have stated: far past any used range.
            d["far_row_pt"] = ws.Rows(1048570).Height
            # Rewrite the Normal size with its own value: if the write makes
            # Excel recompute the standard height from the font, this reads
            # the FONT-DERIVED height in this document's own context.
            try:
                nf.Size = d["normal_size"]
                d["rederived_pt"] = ws.StandardHeight
            except Exception:
                d["rederived_pt"] = None
            wb.Close(False)

            # The candidate's prediction.
            if d["defaultRowHeight"] is not None:
                pad = 0.0 if d["customHeight"] else (d["dyDescent"] or 0.0)
                import math
                pred_px = math.ceil((d["defaultRowHeight"] + pad) / 0.75 - 1e-9)
                d["predicted_pt"] = pred_px * 0.75
            else:
                d["predicted_pt"] = None
            rows.append(d)
            verdict = ("OK" if d["predicted_pt"] is not None
                       and abs(d["predicted_pt"] - d["far_row_pt"]) < 1e-6
                       else "----")
            print("%-24s %-16s %4s drh=%-6s dyd=%-5s ch=%d  std=%-6s far=%-6s rederived=%-6s %s" % (
                d["doc"], d["normal_font"][:16], d["normal_size"],
                d["defaultRowHeight"], d["dyDescent"], d["customHeight"],
                d["standard_height_pt"], d["far_row_pt"], d["rederived_pt"],
                verdict))
    finally:
        excel.Quit()
    out = r"pipeline_data\com_measurements\xlsx_default_row_truth.json"
    with open(out, "w", encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False, indent=1)
    print("wrote %d rows to %s" % (len(rows), out))


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
