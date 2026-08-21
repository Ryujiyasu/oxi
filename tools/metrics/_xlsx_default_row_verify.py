# -*- coding: utf-8 -*-
# Corpus-wide verification of the default-row-height law derived 2026-08-21:
#   customHeight="1"  -> round(defaultRowHeight to nearest 0.75pt pixel)
#   otherwise         -> defaultRowHeight is IGNORED; height = max of the
#                        per-font default heights over every column-style
#                        font, plus the Normal font if any of the 16384
#                        columns has no <col> style.
# Font default heights come from the measured table
# (xlsx_row_height_sweep.json, fresh-workbook Normal-rewrite instrument,
# validated against in-document column-style derivation in battery E).
import glob
import json
import re
import sys
import zipfile


def first_sheet_xml(z):
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


def parse_styles(z):
    st = z.read("xl/styles.xml").decode("utf-8")
    fonts = []
    mf = re.search(r"<fonts[^>]*>(.*?)</fonts>", st, re.S)
    for f in re.findall(r"<font\s*/>|<font>.*?</font>|<font\s[^>]*>.*?</font>",
                        mf.group(1), re.S):
        name = re.search(r'name val="([^"]+)"', f)
        sz = re.search(r'sz val="([^"]+)"', f)
        fonts.append((name.group(1) if name else None,
                      float(sz.group(1)) if sz else None))
    xf_fonts = []
    mx = re.search(r"<cellXfs[^>]*>(.*?)</cellXfs>", st, re.S)
    if mx:
        for xf in re.findall(r"<xf [^>]*/?>", mx.group(1)):
            fid = re.search(r'fontId="(\d+)"', xf)
            xf_fonts.append(int(fid.group(1)) if fid else 0)
    return fonts, xf_fonts


def main():
    table = {}
    for r in json.load(open(
            r"pipeline_data\com_measurements\xlsx_row_height_sweep.json",
            encoding="utf-8")):
        if r.get("applied_name") == r["face"]:
            table[(r["face"], float(r["size"]))] = r["standard_height_pt"]

    truth = {d["doc"]: d for d in json.load(open(
        r"pipeline_data\com_measurements\xlsx_default_row_truth.json",
        encoding="utf-8"))}

    missing = set()
    n_ok = n_bad = n_skip = 0
    bad = []
    for path in sorted(glob.glob(r"tools\golden-test\documents\xlsx\*.xlsx")):
        doc = path.split("\\")[-1][:24]
        t = truth.get(doc)
        if not t or "far_row_pt" not in t:
            continue
        z = zipfile.ZipFile(path)
        try:
            sheet = first_sheet_xml(z)
            fonts, xf_fonts = parse_styles(z)
        except Exception as e:
            print(doc, "PARSE FAIL", e)
            continue

        m = re.search(r"<sheetFormatPr[^>]*>", sheet)
        drh = dyd = None
        ch = False
        if m:
            g = re.search(r'defaultRowHeight="([^"]+)"', m.group(0))
            drh = float(g.group(1)) if g else None
            ch = bool(re.search(r'customHeight="(1|true)"', m.group(0)))

        if ch and drh is not None:
            # customHeight: add 0.05pt, floor to the 96dpi pixel (battery G:
            # 14.93 -> 19px but 14.95 -> 20px; 17.18 -> 22px, 17.2 -> 23px)
            import math
            pred = math.floor((drh + 0.05) / 0.75 + 1e-9) * 0.75
        else:
            # candidate fonts: every col-style font + Normal if uncovered
            normal = (t["normal_font"], float(t["normal_size"]))
            cand = set()
            covered = []
            for col in re.findall(r"<col [^>]*/>", sheet):
                mn = int(re.search(r'min="(\d+)"', col).group(1))
                mx_ = int(re.search(r'max="(\d+)"', col).group(1))
                s = re.search(r'style="(\d+)"', col)
                if s:
                    fid = xf_fonts[int(s.group(1))] if int(s.group(1)) < len(xf_fonts) else 0
                    face, sz = fonts[fid] if fid < len(fonts) else (None, None)
                    if face:
                        cand.add((face, sz))
                        covered.append((mn, mx_))
                # col without style leaves those columns on Normal
            full = False
            if covered:
                covered.sort()
                lo = 1
                full = True
                for mn, mx_ in covered:
                    if mn > lo:
                        full = False
                        break
                    lo = max(lo, mx_ + 1)
                if lo <= 16384:
                    full = False
            if not full:
                cand.add(normal)
            heights = []
            for face, sz in cand:
                key = (face, float(sz))
                if key in table:
                    heights.append(table[key])
                else:
                    missing.add(key)
            if not heights or missing & cand_keys(cand):
                n_skip += 1
                continue
            pred = max(heights)

        actual = t["far_row_pt"]
        if abs(pred - actual) < 1e-6:
            n_ok += 1
        else:
            n_bad += 1
            bad.append((doc, pred, actual, drh, ch))

    print("OK=%d  BAD=%d  SKIP(missing font)=%d" % (n_ok, n_bad, n_skip))
    for doc, pred, actual, drh, ch in bad[:40]:
        print("  %-24s pred=%-6s actual=%-6s drh=%-6s ch=%d" % (
            doc, pred, actual, drh, ch))
    if missing:
        print("missing font table entries:")
        for k in sorted(missing):
            print("  %r" % (k,))
        with open(r"pipeline_data\com_measurements\xlsx_row_missing_fonts.json",
                  "w", encoding="utf-8") as f:
            json.dump(sorted([list(k) for k in missing]), f,
                      ensure_ascii=False, indent=1)


def cand_keys(cand):
    return {(face, float(sz)) for face, sz in cand}


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
