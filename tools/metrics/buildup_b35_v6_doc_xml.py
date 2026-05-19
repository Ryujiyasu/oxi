"""V6: take R05's document.xml directly, vary which OTHER files are present.
Tests if R05's doc.xml structure is necessary for FULL bug, and what minimal
file set + doc.xml triggers full bug."""
import os, zipfile, re

V1_BASE = os.path.abspath("tools/metrics/b35_parametric_repro/B_baseline_match_b35.docx")
R05 = os.path.abspath("tools/metrics/b35123_strip_variants_v2/R05.docx")
OUT_DIR = os.path.abspath("tools/metrics/b35_buildup_v6")
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


def main():
    base = read_docx(V1_BASE)
    r05 = read_docx(R05)

    # First: copy R05's _rels file too AND strip refs to files I don't have
    def copy_r05_doc_minimal(p: dict, file_set_keep: set) -> dict:
        """Copy R05's doc.xml + rels, strip references to files not in file_set_keep."""
        p = dict(p)
        p["word/document.xml"] = r05["word/document.xml"]
        rels = r05["word/_rels/document.xml.rels"].decode("utf-8")
        # Strip <Relationship Target="X"/> where X isn't in file_set_keep
        def keep_rel(m):
            target = re.search(r'Target="([^"]+)"', m.group(0))
            if not target:
                return m.group(0)
            t = target.group(1)
            full = "word/" + t if not t.startswith("../") else t[3:]
            if full in file_set_keep:
                return m.group(0)
            return ""
        rels = re.sub(r'<Relationship\b[^>]+/>', keep_rel, rels)
        p["word/_rels/document.xml.rels"] = rels.encode("utf-8")
        # Also strip headerRef/footerRef from document.xml since we removed those files
        d = p["word/document.xml"].decode("utf-8")
        # Strip <w:headerReference .../> and <w:footerReference .../>
        d = re.sub(r'<w:headerReference[^/]*/>', '', d)
        d = re.sub(r'<w:footerReference[^/]*/>', '', d)
        p["word/document.xml"] = d.encode("utf-8")
        return p

    # V6a: copy R05's document.xml into v1 baseline (no headers/footers)
    p = dict(base)
    p = copy_r05_doc_minimal(p, set(base.keys()))
    write_docx(p, os.path.join(OUT_DIR, "v6a_r05doc_only.docx"))

    # V6b: + styles.xml
    p = dict(base)
    p = add_files(p, r05, ["word/styles.xml"])
    p = copy_r05_doc_minimal(p, set(p.keys()) | {"word/styles.xml"})
    write_docx(p, os.path.join(OUT_DIR, "v6b_r05doc_styles.docx"))

    # V6c: + styles + theme
    p = dict(base)
    p = add_files(p, r05, ["word/styles.xml", "word/theme/theme1.xml"])
    p = copy_r05_doc_minimal(p, set(p.keys()) | {"word/styles.xml", "word/theme/theme1.xml"})
    write_docx(p, os.path.join(OUT_DIR, "v6c_r05doc_styles_theme.docx"))

    # V6d: + styles + theme + fontTable
    p = dict(base)
    p = add_files(p, r05, ["word/styles.xml", "word/theme/theme1.xml", "word/fontTable.xml"])
    p = copy_r05_doc_minimal(p, set(p.keys()) | {"word/styles.xml", "word/theme/theme1.xml", "word/fontTable.xml"})
    write_docx(p, os.path.join(OUT_DIR, "v6d_r05doc_styles_theme_fontTable.docx"))

    # V6e: also use R05 settings.xml
    p = dict(base)
    p["word/settings.xml"] = r05["word/settings.xml"]
    p = add_files(p, r05, ["word/styles.xml", "word/theme/theme1.xml", "word/fontTable.xml"])
    p = copy_r05_doc_minimal(p, set(p.keys()))
    write_docx(p, os.path.join(OUT_DIR, "v6e_r05doc_r05settings_styles_theme_fontTable.docx"))

    print("Done.")


if __name__ == "__main__":
    main()
