# -*- coding: utf-8 -*-
# What Excel makes of a sheet's default row height when the sheet does not
# claim one (customHeight absent): the height is derived from the standard
# font. SX13/SX14 measured four faces at 11pt; hhea, DirectWrite and OS/2
# usWin were ruled out, OS/2 sTypo is the lead but the residue is not a
# constant. This sweep widens the input space — many faces, many sizes — to
# separate a per-em rule from a per-point one and to expose the quantisation.
#
# Method: one workbook, rewrite its Normal style font, read StandardHeight.
# The instrument is validated first against the four SX13 anchors; if any
# anchor disagrees the sweep aborts rather than record numbers from a broken
# instrument.
import json
import sys
import win32com.client

FACES = [
    "ＭＳ Ｐゴシック",  # ＭＳ Ｐゴシック
    "ＭＳ ゴシック",        # ＭＳ ゴシック
    "ＭＳ Ｐ明朝",              # ＭＳ Ｐ明朝
    "ＭＳ 明朝",                    # ＭＳ 明朝
    "MS UI Gothic",
    "メイリオ",                     # メイリオ
    "Meiryo UI",
    "游ゴシック",               # 游ゴシック
    "Yu Gothic UI",
    "游明朝",                           # 游明朝
    "BIZ UDPゴシック",              # BIZ UDPゴシック
    "BIZ UDゴシック",               # BIZ UDゴシック
    "Arial",
    "Calibri",
    "Aptos Narrow",
    "Times New Roman",
    "Segoe UI",
    "Verdana",
    "Tahoma",
]
SIZES = [6, 8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 24, 36]

# face, size, expected height. 游/ＭＳ Ｐゴシック/Calibri agree between the
# fresh-workbook rewrite and in-document column-style derivation (battery E,
# 2026-08-21); ＭＳ 明朝 10/14 come from the 0887f mutations. The SX13 value
# "Meiryo UI 11 -> 18.75" was a misattribution (224e's column style carries
# a different font) — the rewrite instrument's 15.75 is NOT a defect.
ANCHORS = [
    ("游ゴシック", 11, 18.75),
    ("ＭＳ Ｐゴシック", 11, 13.5),
    ("Calibri", 11, 15.0),
    ("ＭＳ 明朝", 10, 12.0),
    ("ＭＳ 明朝", 14, 17.25),
]


def measure(style_font, ws, face, size):
    style_font.Name = face
    style_font.Size = size
    applied_name = style_font.Name
    applied_size = style_font.Size
    return {
        "face": face,
        "size": size,
        "applied_name": applied_name,
        "applied_size": applied_size,
        "standard_height_pt": ws.StandardHeight,
        "standard_width_chars": ws.StandardWidth,
        "height_px": ws.StandardHeight / 0.75,
    }


def main():
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    rows = []
    try:
        wb = excel.Workbooks.Add()
        ws = wb.Worksheets(1)
        style_font = wb.Styles("Normal").Font

        for face, size, expected in ANCHORS:
            r = measure(style_font, ws, face, size)
            got = r["standard_height_pt"]
            ok = abs(got - expected) < 1e-6
            print("anchor %-20s %4s -> %-8s expected %-8s %s" % (
                face, size, got, expected, "OK" if ok else "MISMATCH"))
            if not ok:
                print("instrument invalid: Normal-style rewrite does not "
                      "drive StandardHeight the way open documents do; abort")
                sys.exit(2)

        for face in FACES:
            for size in SIZES:
                r = measure(style_font, ws, face, size)
                if r["applied_name"] != face:
                    r["note"] = "font not applied (missing?)"
                rows.append(r)
                print("%-24s %5s -> %8.2fpt %6.1fpx  (applied: %s)" % (
                    face, size, r["standard_height_pt"], r["height_px"],
                    r["applied_name"]))
        wb.Close(False)
    finally:
        excel.Quit()

    out = r"pipeline_data\com_measurements\xlsx_row_height_sweep.json"
    with open(out, "w", encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False, indent=1)
    print("wrote %d measurements to %s" % (len(rows), out))


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
