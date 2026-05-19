"""V2: finer-grained bisection on the 3 triggers found in S109f v1.

S1 (rows 3-13 strip) killed the bug. Bisect WHICH rows trigger it.
S8 (settings.xml strip) killed the bug. Bisect WHICH setting.
S9 (theme1.xml strip) killed the bug. Test theme font effect.
"""
import os, re, zipfile

SRC = os.path.abspath("tools/golden-test/documents/docx/b35123fe8efc_tokumei_08_01.docx")
OUT_DIR = os.path.abspath("tools/metrics/b35123_strip_variants_v2")
os.makedirs(OUT_DIR, exist_ok=True)


def read_docx(path):
    parts = {}
    with zipfile.ZipFile(path, "r") as z:
        for name in z.namelist():
            parts[name] = z.read(name)
    return parts


def write_docx(parts, out_path):
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in parts.items():
            z.writestr(name, data)


def keep_first_n_rows(doc_xml: str, n: int) -> str:
    tables = list(re.finditer(r"<w:tbl\b[^>]*>.*?</w:tbl>", doc_xml, re.DOTALL))
    if not tables:
        return doc_xml
    tbl_m = tables[0]
    tbl_xml = tbl_m.group(0)
    rows = list(re.finditer(r"<w:tr\b[^>]*>.*?</w:tr>", tbl_xml, re.DOTALL))
    if len(rows) <= n:
        return doc_xml
    new_tbl = tbl_xml[: rows[0].start()] + "".join(r.group(0) for r in rows[:n]) + tbl_xml[rows[-1].end() :]
    return doc_xml[: tbl_m.start()] + new_tbl + doc_xml[tbl_m.end() :]


def strip_settings_element(parts: dict, pattern: str) -> dict:
    parts = dict(parts)
    if "word/settings.xml" not in parts:
        return parts
    s = parts["word/settings.xml"].decode("utf-8")
    s = re.sub(pattern, "", s, flags=re.DOTALL)
    parts["word/settings.xml"] = s.encode("utf-8")
    return parts


def main():
    parts = read_docx(SRC)
    doc_xml_orig = parts["word/document.xml"].decode("utf-8")
    # Row bisection: keep first N rows for various N
    for n in [3, 4, 5, 6, 7, 8, 9, 10]:
        new_parts = dict(parts)
        new_xml = keep_first_n_rows(doc_xml_orig, n)
        new_parts["word/document.xml"] = new_xml.encode("utf-8")
        out = os.path.join(OUT_DIR, f"R{n:02d}.docx")
        write_docx(new_parts, out)
        print(f"Built {out} (keep first {n} rows)")
    # Settings bisection: strip individual settings (charSpace, compat, autoSpace, etc.)
    settings_strips = {
        "Nset_compat": r"<w:compat>.*?</w:compat>",
        "Nset_compat_only_v15_kept": r"<w:compatSetting[^>]*w:name=\"(?!compatibilityMode)[^\"]+\"[^/]*/>",
        "Nset_charspacing": r"<w:characterSpacingControl[^/]*/>",
        "Nset_autospace": r"<w:autoSpaceLikeWord95[^/]*/>",
        "Nset_kinsoku": r"<w:noLineBreaksAfter[^/]*/?>|<w:noLineBreaksBefore[^/]*/?>",
        "Nset_themefonts": r"<w:themeFontLang[^/]*/>",
    }
    for label, pat in settings_strips.items():
        new_parts = strip_settings_element(parts, pat)
        out = os.path.join(OUT_DIR, f"{label}.docx")
        write_docx(new_parts, out)
        print(f"Built {out}")
    # Theme strip variants (already had S9 = strip entire theme1.xml; here try minimal modifications)
    # Just confirm S1 + R variants are the main payload
    print("Done.")


if __name__ == "__main__":
    main()
