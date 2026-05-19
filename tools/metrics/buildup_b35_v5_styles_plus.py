"""V5: styles.xml triggers PARTIAL bug. Need to find what pushes partial→full.

Hypothesis: full bug needs styles + something else (theme/fontTable/specific paragraph structure).
"""
import os, re, zipfile

V1_BASE = os.path.abspath("tools/metrics/b35_parametric_repro/B_baseline_match_b35.docx")
R05 = os.path.abspath("tools/metrics/b35123_strip_variants_v2/R05.docx")
OUT_DIR = os.path.abspath("tools/metrics/b35_buildup_v5")
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


# Same helpers as v4
CT_OVERRIDES = {
    "word/styles.xml": "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
    "word/theme/theme1.xml": "application/vnd.openxmlformats-officedocument.theme+xml",
    "word/fontTable.xml": "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml",
    "word/numbering.xml": "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
}
DOC_REL_TYPES = {
    "word/styles.xml": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
    "word/theme/theme1.xml": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
    "word/fontTable.xml": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable",
    "word/numbering.xml": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
}


def add_files(parts, r05_parts, files):
    parts = dict(parts)
    ct = parts["[Content_Types].xml"].decode("utf-8")
    rels_path = "word/_rels/document.xml.rels"
    rels = parts.get(rels_path, b'<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>').decode("utf-8")
    rid = 100
    for f in files:
        if f not in r05_parts:
            continue
        parts[f] = r05_parts[f]
        if (ct_type := CT_OVERRIDES.get(f)):
            if f'PartName="/{f}"' not in ct:
                ct = ct.replace("</Types>", f'<Override PartName="/{f}" ContentType="{ct_type}"/></Types>')
        if (rel_type := DOC_REL_TYPES.get(f)):
            target = f[len("word/"):] if f.startswith("word/") else f"../{f}"
            if target not in rels:
                rels = rels.replace("</Relationships>", f'<Relationship Id="rId{rid}" Type="{rel_type}" Target="{target}"/></Relationships>')
                rid += 1
    parts["[Content_Types].xml"] = ct.encode("utf-8")
    parts[rels_path] = rels.encode("utf-8")
    return parts


def patch_doc(parts: dict, patch_fn) -> dict:
    parts = dict(parts)
    d = parts["word/document.xml"].decode("utf-8")
    d = patch_fn(d)
    parts["word/document.xml"] = d.encode("utf-8")
    return parts


def main():
    base = read_docx(V1_BASE)
    r05 = read_docx(R05)

    # Variant A: styles + theme
    write_docx(add_files(base, r05, ["word/styles.xml", "word/theme/theme1.xml"]),
               os.path.join(OUT_DIR, "v5_styles_theme.docx"))
    # Variant B: styles + fontTable
    write_docx(add_files(base, r05, ["word/styles.xml", "word/fontTable.xml"]),
               os.path.join(OUT_DIR, "v5_styles_fontTable.docx"))
    # Variant C: styles + theme + fontTable
    write_docx(add_files(base, r05, ["word/styles.xml", "word/theme/theme1.xml", "word/fontTable.xml"]),
               os.path.join(OUT_DIR, "v5_styles_theme_fontTable.docx"))

    # Variant D: ALL files (cumul_ALL = partial in v4). Modify document.xml to use minorEastAsia theme fonts.
    parts_all = add_files(base, r05, list(set(r05.keys()) - set(base.keys())))
    # Patch all rPr <w:rFonts ...> to use minorEastAsia theme refs (like R05)
    def use_theme_fonts(d: str) -> str:
        return re.sub(
            r'<w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝" w:hint="eastAsia"/>',
            '<w:rFonts w:asciiTheme="minorEastAsia" w:eastAsiaTheme="minorEastAsia" w:hAnsiTheme="minorEastAsia" w:hint="eastAsia"/>',
            d,
        )
    write_docx(patch_doc(parts_all, use_theme_fonts),
               os.path.join(OUT_DIR, "v5_all_themefonts.docx"))

    # Variant E: ALL files + use tblStyle "af" instead of inline tblBorders
    def use_tblstyle_af(d: str) -> str:
        # Replace tblPr to use style
        d = re.sub(
            r'<w:tblPr><w:tblW[^/]*/><w:tblBorders>.*?</w:tblBorders></w:tblPr>',
            '<w:tblPr><w:tblStyle w:val="af"/><w:tblW w:w="9070" w:type="dxa"/><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr>',
            d, flags=re.DOTALL,
        )
        return d
    write_docx(patch_doc(parts_all, use_tblstyle_af),
               os.path.join(OUT_DIR, "v5_all_tblstyle_af.docx"))

    # Variant F: ALL files + theme fonts + tblStyle af (combine D+E)
    parts_f = patch_doc(parts_all, lambda d: use_tblstyle_af(use_theme_fonts(d)))
    write_docx(parts_f, os.path.join(OUT_DIR, "v5_all_themefonts_tblstyle.docx"))

    print("Done.")


if __name__ == "__main__":
    main()
