"""V4: aggressive strip-down beyond compat flags. R05 baseline (confirmed trigger),
strip each non-compat element individually."""
import os, re, zipfile

SRC = os.path.abspath("tools/metrics/b35123_strip_variants_v2/R05.docx")  # confirmed bug-triggering
OUT_DIR = os.path.abspath("tools/metrics/b35123_strip_variants_v4")
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


def strip_in_settings(parts: dict, pattern: str) -> dict:
    parts = dict(parts)
    if "word/settings.xml" not in parts:
        return parts
    s = parts["word/settings.xml"].decode("utf-8")
    s = re.sub(pattern, "", s, flags=re.DOTALL)
    parts["word/settings.xml"] = s.encode("utf-8")
    return parts


def strip_in_doc(parts: dict, pattern: str) -> dict:
    parts = dict(parts)
    d = parts["word/document.xml"].decode("utf-8")
    d = re.sub(pattern, "", d, flags=re.DOTALL)
    parts["word/document.xml"] = d.encode("utf-8")
    return parts


def main():
    parts = read_docx(SRC)

    variants = {
        # Strip non-compat settings.xml elements
        "S_compressPunct": (strip_in_settings, r"<w:characterSpacingControl[^/]*/>"),
        "S_drawingGrid": (strip_in_settings, r"<w:drawingGrid\w*[^/]*/>"),
        "S_displayDrawingGrid": (strip_in_settings, r"<w:displayHorizontal[^/]*/>|<w:displayVertical[^/]*/>"),
        "S_bordersDoNotSurround": (strip_in_settings, r"<w:bordersDoNotSurround\w+[^/]*/>"),
        "S_defaultTabStop": (strip_in_settings, r"<w:defaultTabStop[^/]*/>"),
        "S_hdrShapeDefaults": (strip_in_settings, r"<w:hdrShapeDefaults>.*?</w:hdrShapeDefaults>"),
        "S_footnotePr": (strip_in_settings, r"<w:footnotePr>.*?</w:footnotePr>|<w:endnotePr>.*?</w:endnotePr>"),
        "S_zoom": (strip_in_settings, r"<w:zoom[^/]*/>"),
        # Strip document.xml elements
        "D_tblStyle": (strip_in_doc, r"<w:tblStyle[^/]*/>"),
        "D_tblLook": (strip_in_doc, r"<w:tblLook[^/]*/>"),
        # Empty cell paragraph (vMerge continuation rows have empty paragraph) — replace with nothing? probably can't
    }

    for label, (func, pat) in variants.items():
        new_parts = func(parts, pat)
        out = os.path.join(OUT_DIR, f"R05_{label}.docx")
        write_docx(new_parts, out)
        print(f"Built {out}")

    # Combination: strip drawingGrid AND characterSpacingControl
    p1 = strip_in_settings(parts, r"<w:drawingGrid\w*[^/]*/>")
    p1 = strip_in_settings(p1, r"<w:displayHorizontal[^/]*/>|<w:displayVertical[^/]*/>")
    p1 = strip_in_settings(p1, r"<w:characterSpacingControl[^/]*/>")
    out = os.path.join(OUT_DIR, "R05_combo_NoGridNoSpacing.docx")
    write_docx(p1, out)
    print(f"Built {out}")

    print("Done.")


if __name__ == "__main__":
    main()
