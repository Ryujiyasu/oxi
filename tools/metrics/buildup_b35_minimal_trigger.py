"""V4 systematic build-up: start from minimal v1 baseline (no bug), add ONE
R05-only file at a time, identify which addition first triggers the bug.

R05 has 14 files my minimal v1 doesn't have:
  customXml/_rels/item1.xml.rels, customXml/item1.xml, customXml/itemProps1.xml,
  docProps/app.xml, docProps/core.xml,
  word/endnotes.xml, word/fontTable.xml, word/footer1.xml, word/footnotes.xml,
  word/header1.xml, word/numbering.xml, word/styles.xml, word/theme/theme1.xml,
  word/webSettings.xml

For each, add the file + update [Content_Types].xml + word/_rels/document.xml.rels
to reference it. Test if bug appears.
"""
import os, re, zipfile

V1_BASE = os.path.abspath("tools/metrics/b35_parametric_repro/B_baseline_match_b35.docx")
R05 = os.path.abspath("tools/metrics/b35123_strip_variants_v2/R05.docx")
OUT_DIR = os.path.abspath("tools/metrics/b35_buildup")
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


# Content type for each part name (Override path)
CT_OVERRIDES = {
    "word/styles.xml": "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
    "word/theme/theme1.xml": "application/vnd.openxmlformats-officedocument.theme+xml",
    "word/fontTable.xml": "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml",
    "word/numbering.xml": "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
    "word/header1.xml": "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
    "word/footer1.xml": "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
    "word/footnotes.xml": "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
    "word/endnotes.xml": "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml",
    "word/webSettings.xml": "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml",
    "docProps/app.xml": "application/vnd.openxmlformats-officedocument.extended-properties+xml",
    "docProps/core.xml": "application/vnd.openxmlformats-package.core-properties+xml",
    "customXml/item1.xml": "application/xml",
    "customXml/itemProps1.xml": "application/vnd.openxmlformats-officedocument.customXmlProperties+xml",
}

# Rel type for relationships that go in word/_rels/document.xml.rels
DOC_REL_TYPES = {
    "word/styles.xml": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
    "word/theme/theme1.xml": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
    "word/fontTable.xml": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable",
    "word/numbering.xml": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
    "word/header1.xml": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
    "word/footer1.xml": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
    "word/footnotes.xml": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes",
    "word/endnotes.xml": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes",
    "word/webSettings.xml": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings",
}


def add_files(base_parts: dict, r05_parts: dict, files_to_add: list[str]) -> dict:
    """Add the specified file(s) from R05 to base. Update [Content_Types].xml and rels."""
    parts = dict(base_parts)
    ct = parts["[Content_Types].xml"].decode("utf-8")
    rels_path = "word/_rels/document.xml.rels"
    if rels_path not in parts:
        rels = '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
    else:
        rels = parts[rels_path].decode("utf-8")

    next_rid = 100
    for fname in files_to_add:
        # Add file content from R05
        if fname not in r05_parts:
            continue
        parts[fname] = r05_parts[fname]

        # Add ContentType Override
        ct_type = CT_OVERRIDES.get(fname)
        if ct_type and f'PartName="/{fname}"' not in ct:
            ct = ct.replace("</Types>", f'<Override PartName="/{fname}" ContentType="{ct_type}"/></Types>')

        # Add document.xml.rels entry
        rel_type = DOC_REL_TYPES.get(fname)
        if rel_type:
            # Need a relative target from /word/_rels/document.xml.rels
            if fname.startswith("word/"):
                target = fname[len("word/"):]
            else:
                target = f"../{fname}"
            if target not in rels:
                rels = rels.replace(
                    "</Relationships>",
                    f'<Relationship Id="rId{next_rid}" Type="{rel_type}" Target="{target}"/></Relationships>',
                )
                next_rid += 1

    # Also need word/theme/theme1.xml's rels if styles uses it (theme has its own _rels in some cases)
    # For now skip — theme is typically self-contained

    parts["[Content_Types].xml"] = ct.encode("utf-8")
    parts[rels_path] = rels.encode("utf-8")
    return parts


def main():
    base = read_docx(V1_BASE)
    r05 = read_docx(R05)

    # All R05-only files
    r05_only_files = sorted(set(r05.keys()) - set(base.keys()))
    print(f"R05-only files ({len(r05_only_files)}):")
    for f in r05_only_files:
        print(f"  {f}")

    # Single-file additions: add each ONE file
    for fname in r05_only_files:
        new_parts = add_files(base, r05, [fname])
        label = f"add_{fname.replace('/', '_').replace('.xml', '').replace('.rels', '_rels')}"
        out = os.path.join(OUT_DIR, f"{label}.docx")
        write_docx(new_parts, out)
        print(f"Built {out}")

    # Cumulative additions (pairs / groups of likely suspects)
    cumulative_groups = [
        ("S+T", ["word/styles.xml", "word/theme/theme1.xml"]),
        ("S+T+F", ["word/styles.xml", "word/theme/theme1.xml", "word/fontTable.xml"]),
        ("S+T+F+N", ["word/styles.xml", "word/theme/theme1.xml", "word/fontTable.xml", "word/numbering.xml"]),
        ("ALL", r05_only_files),
    ]
    for label, files in cumulative_groups:
        new_parts = add_files(base, r05, files)
        out = os.path.join(OUT_DIR, f"cumul_{label}.docx")
        write_docx(new_parts, out)
        print(f"Built {out}")

    print("Done.")


if __name__ == "__main__":
    main()
