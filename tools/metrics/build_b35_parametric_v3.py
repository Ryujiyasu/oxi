"""V3: vary TEXT LENGTH and vMerge GROUP ROWS to triangulate N selection rule.

Hypothesis to test:
- N depends on text length (3, 5, 7, 10, 15 chars in col1)
- N depends on vMerge group rows (1, 2, 3, 4, 5 continuation rows)

Keep R05's full file set, only modify table row count + text content.
"""
import os, re, zipfile

SRC = os.path.abspath("tools/metrics/b35123_strip_variants_v2/R05.docx")
OUT_DIR = os.path.abspath("tools/metrics/b35_parametric_repro_v3")
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


def patch(parts: dict, text_chars: int, n_rows_total: int) -> dict:
    """Modify table to have exactly n_rows_total rows + replace col1 text."""
    parts = dict(parts)
    d = parts["word/document.xml"].decode("utf-8")
    # Replace text in col1 cell of row 1 (the vMerge restart cell with 組織的管理措置)
    # The text content is between <w:t>...</w:t>; find the one with 組織的管理措置
    base_text = "アイウエオカキクケコサシスセソタチツテト"  # 20 fullwidth katakana available
    new_text = base_text[:text_chars]
    d = re.sub(r"<w:t>組織的管理措置</w:t>", f"<w:t>{new_text}</w:t>", d, count=1)
    # Now adjust the table row count to n_rows_total (header + restart + (n-2) continuation)
    tbl_m = re.search(r"<w:tbl\b[^>]*>.*?</w:tbl>", d, re.DOTALL)
    if tbl_m:
        tbl_xml = tbl_m.group(0)
        rows = list(re.finditer(r"<w:tr\b[^>]*>.*?</w:tr>", tbl_xml, re.DOTALL))
        if len(rows) >= 2:
            # R05 has 5 rows (header + restart + 3 continue). We want n_rows_total rows.
            # Keep header (rows[0]) + restart (rows[1]) and duplicate continuation row as needed.
            header = rows[0].group(0)
            restart_row = rows[1].group(0)
            continue_row = rows[2].group(0) if len(rows) > 2 else None
            if continue_row is None:
                return parts  # can't make continue rows
            # Build new table with n_rows_total rows
            n_continue = max(0, n_rows_total - 2)
            new_rows = header + restart_row + (continue_row * n_continue)
            new_tbl = tbl_xml[: rows[0].start()] + new_rows + tbl_xml[rows[-1].end():]
            d = d[: tbl_m.start()] + new_tbl + d[tbl_m.end():]
    parts["word/document.xml"] = d.encode("utf-8")
    return parts


def main():
    parts = read_docx(SRC)
    # 5x5 grid: text_chars x n_rows_total
    for text_chars in [3, 5, 7, 10, 15]:
        for n_rows in [2, 3, 4, 5, 7]:  # 2=header+restart only (no continuation)
            new_parts = patch(parts, text_chars, n_rows)
            out = os.path.join(OUT_DIR, f"T{text_chars:02d}_R{n_rows}.docx")
            write_docx(new_parts, out)
            print(f"Built {out} (text {text_chars} chars, {n_rows} rows)")
    print("Done.")


if __name__ == "__main__":
    main()
