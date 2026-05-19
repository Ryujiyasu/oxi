"""V2: use R05 as base (confirmed trigger), modify only:
- font size (sz) on all runs in target cell
- docGrid charSpace
- col1 tcW
Keep all other files (styles.xml / theme1.xml / fontTable.xml / etc.) intact.
"""
import os, re, zipfile

SRC = os.path.abspath("tools/metrics/b35123_strip_variants_v2/R05.docx")
OUT_DIR = os.path.abspath("tools/metrics/b35_parametric_repro_v2")
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


def patch(parts: dict, sz_halfpt: int, charspace: int, tcw_col1: int):
    parts = dict(parts)
    # 1. Patch document.xml: change ALL sz/szCs to target, change tcW for col1, change charSpace
    d = parts["word/document.xml"].decode("utf-8")
    d = re.sub(r'<w:sz w:val="\d+"/>', f'<w:sz w:val="{sz_halfpt}"/>', d)
    d = re.sub(r'<w:szCs w:val="\d+"/>', f'<w:szCs w:val="{sz_halfpt}"/>', d)
    # Change col1 tcW (the smaller tcW; in R05 it's 1271)
    d = re.sub(r'<w:tcW w:w="1271" w:type="dxa"/>', f'<w:tcW w:w="{tcw_col1}" w:type="dxa"/>', d)
    # Adjust col2 to keep total table width = 9067
    d = re.sub(r'<w:tcW w:w="7796" w:type="dxa"/>', f'<w:tcW w:w="{9067-tcw_col1}" w:type="dxa"/>', d)
    # Update tblGrid col1 width too (in R05: 1197 was the original tblGrid; keep but update if needed)
    d = re.sub(r'<w:gridCol w:w="1197"/>', f'<w:gridCol w:w="{tcw_col1-74}"/>', d)
    d = re.sub(r'<w:gridCol w:w="7870"/>', f'<w:gridCol w:w="{9067-tcw_col1+74}"/>', d)
    # Update tblW total
    d = re.sub(r'<w:tblW w:w="9067" w:type="dxa"/>', f'<w:tblW w:w="9067" w:type="dxa"/>', d)
    # Change docGrid charSpace
    d = re.sub(r'<w:docGrid w:type="linesAndChars" w:linePitch="350" w:charSpace="-?\d+"/>',
               f'<w:docGrid w:type="linesAndChars" w:linePitch="350" w:charSpace="{charspace}"/>', d)
    parts["word/document.xml"] = d.encode("utf-8")
    # 2. Also patch styles.xml's docDefaults sz (it has sz=21 default)
    if "word/styles.xml" in parts:
        s = parts["word/styles.xml"].decode("utf-8")
        # docDefault sz
        s = re.sub(r'(<w:rPrDefault>.*?<w:sz w:val=")\d+(".*?</w:rPrDefault>)',
                   lambda m: f"{m.group(1)}{sz_halfpt}{m.group(2)}", s, flags=re.DOTALL)
        parts["word/styles.xml"] = s.encode("utf-8")
    return parts


def main():
    parts = read_docx(SRC)
    # vary fs at fixed cs=-2714, tcW=1271
    for sz in [18, 21, 24, 28]:
        new_parts = patch(parts, sz, -2714, 1271)
        out = os.path.join(OUT_DIR, f"R05P_fs{sz:02d}_csn02714_tcw1271.docx")
        write_docx(new_parts, out)
        print(f"Built {out}")
    # vary charSpace at fixed fs=21, tcW=1271
    for cs in [0, -500, -1500, -2714, -4000]:
        sign = 'p' if cs >= 0 else 'n'
        new_parts = patch(parts, 21, cs, 1271)
        out = os.path.join(OUT_DIR, f"R05P_fs21_cs{sign}{abs(cs):05d}_tcw1271.docx")
        write_docx(new_parts, out)
        print(f"Built {out}")
    # vary tcW at fixed fs=21, cs=-2714
    for tcw in [800, 1000, 1271, 1500, 2000]:
        new_parts = patch(parts, 21, -2714, tcw)
        out = os.path.join(OUT_DIR, f"R05P_fs21_csn02714_tcw{tcw:04d}.docx")
        write_docx(new_parts, out)
        print(f"Built {out}")
    print("Done.")


if __name__ == "__main__":
    main()
