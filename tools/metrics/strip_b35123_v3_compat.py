"""V3: strip individual compat flags to identify the trigger."""
import os, re, zipfile

SRC = os.path.abspath("tools/golden-test/documents/docx/b35123fe8efc_tokumei_08_01.docx")
OUT_DIR = os.path.abspath("tools/metrics/b35123_strip_variants_v3")
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


COMPAT_FLAGS = [
    "spaceForUL",
    "balanceSingleByteDoubleByteWidth",
    "doNotLeaveBackslashAlone",
    "ulTrailSpace",
    "doNotExpandShiftReturn",
    "adjustLineHeightInTable",
    "useFELayout",
]


def strip_flag(parts: dict, flag: str) -> dict:
    parts = dict(parts)
    s = parts["word/settings.xml"].decode("utf-8")
    s = re.sub(rf"<w:{flag}\b[^/]*/>", "", s)
    parts["word/settings.xml"] = s.encode("utf-8")
    return parts


def main():
    parts = read_docx(SRC)
    # 1. Strip each compat flag individually
    for flag in COMPAT_FLAGS:
        new_parts = strip_flag(parts, flag)
        out = os.path.join(OUT_DIR, f"C_strip_{flag}.docx")
        write_docx(new_parts, out)
        print(f"Built {out}")
    # 2. Also: strip ALL the non-compatSetting flags (the standalone <w:flag/> ones)
    new_parts = dict(parts)
    s = new_parts["word/settings.xml"].decode("utf-8")
    for flag in COMPAT_FLAGS:
        s = re.sub(rf"<w:{flag}\b[^/]*/>", "", s)
    new_parts["word/settings.xml"] = s.encode("utf-8")
    out = os.path.join(OUT_DIR, "C_strip_all_flags.docx")
    write_docx(new_parts, out)
    print(f"Built {out}")
    # 3. Keep ONLY useFELayout (strip everything else)
    new_parts = dict(parts)
    s = new_parts["word/settings.xml"].decode("utf-8")
    for flag in COMPAT_FLAGS:
        if flag != "useFELayout":
            s = re.sub(rf"<w:{flag}\b[^/]*/>", "", s)
    new_parts["word/settings.xml"] = s.encode("utf-8")
    out = os.path.join(OUT_DIR, "C_keep_only_useFELayout.docx")
    write_docx(new_parts, out)
    print(f"Built {out}")
    # 4. Keep ONLY adjustLineHeightInTable
    new_parts = dict(parts)
    s = new_parts["word/settings.xml"].decode("utf-8")
    for flag in COMPAT_FLAGS:
        if flag != "adjustLineHeightInTable":
            s = re.sub(rf"<w:{flag}\b[^/]*/>", "", s)
    new_parts["word/settings.xml"] = s.encode("utf-8")
    out = os.path.join(OUT_DIR, "C_keep_only_adjustLineHeightInTable.docx")
    write_docx(new_parts, out)
    print(f"Built {out}")
    # 5. Keep ONLY balanceSingleByteDoubleByteWidth
    new_parts = dict(parts)
    s = new_parts["word/settings.xml"].decode("utf-8")
    for flag in COMPAT_FLAGS:
        if flag != "balanceSingleByteDoubleByteWidth":
            s = re.sub(rf"<w:{flag}\b[^/]*/>", "", s)
    new_parts["word/settings.xml"] = s.encode("utf-8")
    out = os.path.join(OUT_DIR, "C_keep_only_balanceSBDB.docx")
    write_docx(new_parts, out)
    print(f"Built {out}")


if __name__ == "__main__":
    main()
