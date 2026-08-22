# -*- coding: utf-8 -*-
r"""Which property makes one level of indent 12 pixels instead of 15?

`_xlsx_indent.py` gets 15 pixels a level from Excel for every face, size and
alignment it asks about, and again when the workbook's Normal font is changed.
The `h2daa*dendeba_kmc` trio gets 12: its indent="1" rows sit 3 pixels left of
ours and its indent="2" rows 6. The cell font is the same ＭＳ Ｐゴシック the
probe asks about, so the difference is somewhere else in the workbook.

This takes the workbook itself and puts one property back to what a plain book
has, one variant at a time, then asks Excel to draw each. The variant whose
indent goes back to 15 names the property. The `std:` variants then sweep the
workbook's standard font, to see what a level is made of.

Read the level off row 8, which carries indent="1" in ＭＳ Ｐゴシック 11:
`level = x8 - 5`. Row 5 carries indent="2" in the same face at 10 point:
`level = (x5 - 11) / 2`.

    python tools\metrics\_xlsx_indent_bisect.py
    python tools\metrics\_xlsx_indent_bisect.py --reuse
"""
import argparse
import json
import os
import re
import subprocess
import sys
import zipfile
from pathlib import Path

import numpy as np
from PIL import Image

REPO = Path(__file__).resolve().parents[2]
SHOOTER = Path(__file__).resolve().parent / "_xlsx_screen_shot.ps1"
RENDERER = REPO / "tools" / "oxi-xlsx-renderer" / "target" / "release" / "oxi-xlsx-renderer.exe"
SOURCE = (REPO / "tools" / "golden-test" / "documents" / "xlsx"
          / "5c74ec72c6e1_h2daa2023_dendeba_kmc.xlsx")
SCRATCH = Path(r"C:\tmp\xlsx_indent_bisect")
# The rows to read: 4 and 5 carry indent="2", 8 carries indent="1", and 3, 6
# and 7 carry none, so a move that is not the indent's shows up too.
ROWS = [2, 3, 4, 5, 6, 7, 8, 9]

MSP = "ＭＳ Ｐゴシック"
GOTHIC = "ＭＳ ゴシック"
YU = "游ゴシック"
MEIRYO = "メイリオ"

VARIANTS = ["asis", "normal_msp", "normal_style_msp", "normal_calibri",
            "old_theme_version"]
# What one level is worth when the workbook's standard font is each of these.
# If it is three of that font's spaces, ＭＳ ゴシック — whose space is as wide
# as its digit — must be worth 24, and ＭＳ Ｐゴシック at 8 point 9.
VARIANTS += ["std:" + face + ":" + points for face, points in
             [(GOTHIC, "11"), ("Arial", "11")]]


def surgery(name, parts):
    """The parts of the workbook this variant rewrites."""
    theme = parts["xl/theme/theme1.xml"].decode("utf-8")
    styles = parts["xl/styles.xml"].decode("utf-8")
    sheet = parts["xl/worksheets/sheet1.xml"].decode("utf-8")
    shared = parts["xl/sharedStrings.xml"].decode("utf-8")
    book = parts["xl/workbook.xml"].decode("utf-8")
    first_font = re.search(r"<font>.*?</font>|<font/>", styles, re.S).group(0)
    if name == "asis":
        return {}
    if name == "theme_minor_latin":
        # The minor font a `scheme="minor"` run resolves to.
        held = re.search(r"<a:minorFont>.*?</a:minorFont>", theme, re.S).group(0)
        swapped = re.sub(r'<a:latin typeface="[^"]*"',
                         '<a:latin typeface="Calibri"', held, count=1)
        swapped = re.sub(r'<a:ea typeface="[^"]*"', '<a:ea typeface=""',
                         swapped, count=1)
        return {"xl/theme/theme1.xml": theme.replace(held, swapped)}
    if name == "normal_msp":
        # The Normal style's font — the workbook's standard font.
        return {"xl/styles.xml": styles.replace(
            first_font,
            '<font><sz val="11"/><color theme="1"/><name val="' + MSP + '"/>'
            '<family val="3"/><charset val="128"/></font>', 1)}
    if name == "normal_calibri":
        return {"xl/styles.xml": styles.replace(
            first_font,
            '<font><sz val="11"/><color theme="1"/><name val="Calibri"/>'
            '<family val="2"/><scheme val="minor"/></font>', 1)}
    if name.startswith("std:"):
        # The standard font named outright: no scheme, so the theme has no
        # say in what it resolves to.
        _, face, points = name.split(":")
        return {"xl/styles.xml": styles.replace(
            first_font,
            '<font><sz val="' + points + '"/><color theme="1"/><name val="'
            + face + '"/><family val="3"/><charset val="128"/></font>', 1)}
    if name == "normal_style_msp":
        # Leaves the first font alone and points the Normal style at another:
        # tells the font the standard is taken from apart from the first one
        # in the list, which the workbook happens to have them agree on.
        held = re.search(r"<cellStyleXfs count=\"\d+\"><xf [^>]*?fontId=\"0\"",
                         styles).group(0)
        return {"xl/styles.xml": styles.replace(
            held, held.replace('fontId="0"', 'fontId="8"'), 1)}
    if name == "xf_plain":
        # The indented cells hang off a named style (xfId="8"); this hangs
        # them off Normal instead, everything else the same.
        return {"xl/styles.xml": styles.replace('xfId="8"', 'xfId="0"')}
    if name == "row_plain":
        # The rows carry a style of their own, with customFormat.
        return {"xl/worksheets/sheet1.xml":
                re.sub(r'(<row [^>]*?) s="368" customFormat="1"', r"\1", sheet)}
    if name == "cols_plain":
        return {"xl/worksheets/sheet1.xml": re.sub(r' style="366"', "", sheet)}
    if name == "no_default_width":
        return {"xl/worksheets/sheet1.xml":
                re.sub(r' defaultColWidth="[^"]*"', "", sheet)}
    if name == "no_phonetic":
        stripped = re.sub(r"<rPh[^>]*>.*?</rPh>", "", shared, flags=re.S)
        stripped = re.sub(r"<phoneticPr[^>]*/>", "", stripped)
        return {"xl/sharedStrings.xml": stripped,
                "xl/worksheets/sheet1.xml": re.sub(r"<phoneticPr[^>]*/>", "", sheet)}
    if name == "old_theme_version":
        return {"xl/workbook.xml": book.replace('defaultThemeVersion="166925"',
                                                'defaultThemeVersion="124226"')}
    if name == "no_theme_version":
        return {"xl/workbook.xml": re.sub(r' defaultThemeVersion="[^"]*"', "", book)}
    raise SystemExit("no surgery named " + name)


def build():
    SCRATCH.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(SOURCE) as source:
        order = source.namelist()
        parts = {name: source.read(name) for name in order}
    made = []
    for name in VARIANTS:
        changed = surgery(name, parts)
        out = SCRATCH / (name.replace(":", "_") + ".xlsx")
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as book:
            for part in order:
                if part in changed:
                    book.writestr(part, changed[part].encode("utf-8"))
                else:
                    book.writestr(part, parts[part])
        made.append((name, out))
    return made


def shoot(made):
    listing = SCRATCH / "_batch.txt"
    lines = []
    for _, book in made:
        picture = book.with_suffix(".excel.png")
        picture.unlink(missing_ok=True)
        lines.append(str(book.resolve()) + "\t" + str(picture.resolve()))
    # With a mark at the front: PowerShell 5.1 reads a plain UTF-8 file as the
    # system codepage, which loses every workbook whose name is not ASCII.
    listing.write_text("\n".join(lines), encoding="utf-8-sig")
    done = subprocess.run(["powershell", "-NoProfile", "-File", str(SHOOTER),
                           "-ListFile", str(listing.resolve())],
                          capture_output=True, text=True, encoding="utf-8",
                          errors="replace", timeout=1800)
    failed = [line for line in done.stdout.splitlines() if not line.startswith("ok")]
    print("\n".join(failed[-6:]) if failed else "Excel drew them all")
    listing.unlink(missing_ok=True)


def rows_of(book):
    environment = dict(os.environ, OXI_XLSX_DUMP_ROWS="1")
    held = json.loads(
        (REPO / "pipeline_data" / "xlsx_used_range.json").read_text(encoding="utf-8"))
    found = held.get(SOURCE.stem, {}).get("excel")
    if found:
        environment["OXI_XLSX_RANGE"] = ",".join(str(number) for number in found)
    ours = book.with_suffix(".oxi.png")
    done = subprocess.run([str(RENDERER), str(book), str(ours), "96"],
                          capture_output=True, timeout=300, env=environment)
    heights = {}
    for line in done.stdout.decode("utf-8", "replace").splitlines():
        parts = line.split()
        if len(parts) == 4 and parts[0] == "row":
            heights[int(parts[1])] = int(float(parts[3]))
    return ours, heights


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--reuse", action="store_true")
    args = parser.parse_args()
    made = build()
    if not args.reuse:
        shoot(made)
    print(f"{'variant':<24}" + "".join(f"{'row ' + str(r):>8}" for r in ROWS)
          + f"{'level(8)':>10}{'level(5)':>10}")
    for name, book in made:
        picture = book.with_suffix(".excel.png")
        if not picture.exists():
            print(f"{name:<24}(no picture)")
            continue
        ours_png, heights = rows_of(book)
        truth = np.asarray(Image.open(picture).convert("L"))
        mine = np.asarray(Image.open(ours_png).convert("L"))
        edges, at = {}, 0
        for index in sorted(heights):
            edges[index] = (at, at + heights[index])
            at += heights[index]
        line, ourline, seen = "", "", {}
        for row in ROWS:
            if row not in edges or edges[row][1] > truth.shape[0]:
                line += f"{'-':>8}"
                ourline += f"{'-':>8}"
                continue
            top, foot = edges[row]
            lit = np.flatnonzero((truth[top:foot] < 128).sum(axis=0))
            mylit = np.flatnonzero((mine[top:foot] < 128).sum(axis=0))
            seen[row] = int(lit[0]) if lit.size else None
            line += f"{seen[row] if seen[row] is not None else -1:>8}"
            ourline += f"{int(mylit[0]) if mylit.size else -1:>8}"
        level8 = f"{seen[8] - 5}" if seen.get(8) is not None else "-"
        level5 = f"{(seen[5] - 11) / 2:g}" if seen.get(5) is not None else "-"
        print(f"{name:<24}{line}{level8:>10}{level5:>10}")
        if name == "asis":
            print(f"{'  (ours)':<24}{ourline}")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
