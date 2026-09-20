# -*- coding: utf-8 -*-
"""Static feature census of every corpus docx: which documents CAN a change
touch? A layout change gated on a feature cannot move a document that lacks
it, so the iteration gate only needs the documents whose census row matches
the change's predicate (plus a byte-identity sample of the rest to catch a
wrong predicate). Full gates stay for commits.

  python feature_census.py build            # scan corpus -> pipeline_data/feature_census.json
  python feature_census.py query <expr>     # e.g. "cols>1 and wrap_tb>0"
  python feature_census.py sample <expr> N  # N random docs NOT matching expr

Rows are keyed "set/type__sha" style ids used by the gates; each row carries
plain counts/flags read from document.xml / styles.xml / settings.xml /
theme1.xml with regexes (no rendering).
"""
import os, sys, re, json, glob, zipfile, random
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "feature_census.json"
ROOTS = [REPO / "pipeline_data" / "docx_corpus", REPO / "tools" / "golden-test"]


def part(z, name):
    try:
        return z.read(name).decode("utf-8", "replace")
    except KeyError:
        return ""


def census(path):
    z = zipfile.ZipFile(path)
    d = part(z, "word/document.xml"); st = part(z, "word/styles.xml"); se = part(z, "word/settings.xml"); th = part(z, "word/theme/theme1.xml")
    anchors = re.findall(r"<wp:anchor[^>]*>.*?</wp:anchor>", d, re.S)
    r = {
        "paras": len(re.findall(r"<w:p[ >/]", d)),
        "tables": d.count("<w:tbl>"),
        "cols_max": max([int(v) for v in re.findall(r'<w:cols[^>]*w:num="(\d+)"', d)] + [1]),
        "sections": d.count("<w:sectPr"),
        "anchors": len(anchors),
        "wrap_tb": sum("<wp:wrapTopAndBottom" in a for a in anchors),
        "wrap_tb_para": sum("<wp:wrapTopAndBottom" in a and 'positionV relativeFrom="paragraph"' in a for a in anchors),
        "wrap_square": sum("<wp:wrapSquare" in a for a in anchors),
        "wrap_none": sum("<wp:wrapNone" in a for a in anchors),
        "behind_doc": sum('behindDoc="1"' in a for a in anchors),
        "sp_autofit": sum("<a:spAutoFit" in a for a in anchors),
        "txbx": d.count("<w:txbxContent>"),
        "vml": len(re.findall(r"<w:pict[ >]", d)),
        "inline_pics": d.count("<wp:inline"),
        "docgrid_type": (re.search(r'<w:docGrid[^>]*w:type="(\w+)"', d) or [None, ""])[1],
        "char_space": (lambda m: int(m.group(1)) if m else 0)(re.search(r'<w:docGrid[^>]*w:charSpace="(-?\d+)"', d)),
        "line_pitch": (lambda m: int(m.group(1)) if m else 0)(re.search(r'<w:docGrid[^>]*w:linePitch="(\d+)"', d)),
        "compat": (lambda m: int(m.group(1)) if m else 0)(re.search(r'w:name="compatibilityMode"[^>]*w:val="(\d+)"', se)),
        "fe_layout": "<w:useFELayout" in se,
        "csc": (re.search(r'<w:characterSpacingControl w:val="(\w+)"', se) or [None, ""])[1],
        "theme_font_lang_ea": (re.search(r'<w:themeFontLang[^>]*w:eastAsia="([^"]+)"', se) or [None, ""])[1],
        "theme_ea_empty": '<a:ea typeface=""/>' in th,
        "theme_jpan": (re.search(r'<a:font script="Jpan" typeface="([^"]+)"', th) or [None, ""])[1],
        "kern": "<w:kern" in st or "<w:kern" in d,
        "lrpb": d.count("<w:lastRenderedPageBreak"),
        "page_breaks": d.count('w:type="page"'),
        "keep_next": d.count("<w:keepNext"),
        "pbdr": d.count("<w:pBdr>"),
        "ruby": d.count("<w:ruby>"),
        "vertical": 'w:orient="landscape"' in d or "<w:textDirection" in d,
        "fields": d.count("<w:instrText") + d.count("<w:fldSimple"),
        "toc": "TOC" in d and "instrText" in d,
        "footnotes": d.count("<w:footnoteReference"),
        "ea_latin_run": len(re.findall(r'<w:rFonts[^>]*w:eastAsia="(Times New Roman|Arial|Calibri|Cambria|Century|Georgia|Verdana|Tahoma)"', d)),
        "ea_theme_ref": len(re.findall(r'w:eastAsiaTheme="minorEastAsia"', d + st)),
        "fonts": sorted(set(re.findall(r'w:eastAsia="([^"]+)"', d + st)))[:12],
        "has_cjk": bool(re.search(r"[぀-ヿ一-鿿]", d)),
        "pair_split_runs": len(re.findall(r"[、。）」]</w:t></w:r>\s*<w:r[ >]", d)) + len(re.findall(r"[（「]</w:t></w:r>\s*<w:r[ >]", d)),
    }
    return r


def doc_id(path):
    p = Path(path)
    if "docx_corpus" in p.parts:
        i = p.parts.index("docx_corpus")
        return f"{p.parts[i+1]}/{p.parts[i+2]}__{p.stem}"
    return f"golden/{p.stem}"


def build():
    rows = {}
    files = []
    for root in ROOTS:
        files += glob.glob(str(root / "**" / "*.docx"), recursive=True)
    for f in files:
        try:
            rows[doc_id(f)] = census(f)
        except Exception as e:
            rows[doc_id(f)] = {"error": str(e)[:80]}
        rows[doc_id(f)]["path"] = f
    OUT.write_text(json.dumps(rows, ensure_ascii=False, indent=0), encoding="utf-8")
    print("census:", len(rows), "docs ->", OUT)


def load():
    return json.loads(OUT.read_text(encoding="utf-8"))


def query(expr):
    rows = load(); hit = []
    for did, r in rows.items():
        if "error" in r:
            continue
        try:
            if eval(expr, {}, dict(r)):
                hit.append(did)
        except Exception:
            pass
    return hit


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "build"
    if cmd == "build":
        build()
    elif cmd == "query":
        h = query(sys.argv[2]); print(len(h)); [print(x) for x in h]
    elif cmd == "sample":
        h = set(query(sys.argv[2])); rest = [d for d in load() if d not in h]
        random.seed(int(sys.argv[4]) if len(sys.argv) > 4 else 0)
        for x in random.sample(rest, min(int(sys.argv[3]), len(rest))):
            print(x)
