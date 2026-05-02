"""§4.7c Mech 3 — formula investigation via controlled clone-doc probes.

Session 51 found Mech 3 fires only with real-doc supporting files. We
clone 7f272a's supporting files and inject controlled probe text varying:
  - n_yak in line (1, 2, 3, 4, 6)
  - probe length to control overflow vs no-overflow
  - alignment {left, both}

Hypotheses:
  (a) slack-distribution: total_compression = overflow, per-yak = overflow/n_yak
  (b) grid-snap: per-yak ∈ {0.5pt × k}, target snaps to some grid
  (c) font-size×const: per-yak = font_size × constant_ratio

For each probe we record:
  - per-line: n_chars, n_yak, observed line width, font sizes
  - per-yak: ch, position, advance, ratio
"""
import json, os, sys, time, zipfile, shutil, tempfile
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

SRC_REAL = os.path.abspath(
    r"C:\Users\ryuji\oxi-1\tools\golden-test\documents\docx\7f272a2dfd3b_index-21.docx")
OUT_DIR = os.path.abspath(r"C:\Users\ryuji\oxi-1\pipeline_data\mech3_compress_repro_docs")
RESULT_PATH = os.path.abspath(
    r"C:\Users\ryuji\oxi-1\pipeline_data\mech3_compression_formula.json")

os.makedirs(OUT_DIR, exist_ok=True)

YAKUMONO = set("（「『【〔｛〈《［）」』】〕｝〉》］、。，．—")


def make_doc(text, jc, page_w_tw=11906, margin_lr_tw=1304):
    """Wrap probe text in a single paragraph with given jc + page size."""
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
            ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<w:body><w:p>'
            f'<w:pPr><w:jc w:val="{jc}"/>'
            '<w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr></w:pPr>'
            '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>'
            f'<w:t>{text}</w:t></w:r></w:p>'
            f'<w:sectPr><w:pgSz w:w="{page_w_tw}" w:h="16838"/>'
            f'<w:pgMar w:top="1134" w:right="{margin_lr_tw}" w:bottom="1134" w:left="{margin_lr_tw}"'
            ' w:header="851" w:footer="992" w:gutter="0"/>'
            '<w:cols w:space="720"/>'
            '<w:docGrid w:type="lines" w:linePitch="360"/>'
            '</w:sectPr></w:body></w:document>')


def clone_with_doc(label, text, jc, page_w_tw=11906, margin_lr_tw=1304):
    out_path = os.path.join(OUT_DIR, f"{label}.docx")
    tmp = tempfile.mkdtemp(prefix="m3_")
    try:
        with zipfile.ZipFile(SRC_REAL) as z:
            z.extractall(tmp)
        with open(os.path.join(tmp, "word", "document.xml"), "w", encoding="utf-8") as f:
            f.write(make_doc(text, jc, page_w_tw, margin_lr_tw))
        with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
            for root, _, files in os.walk(tmp):
                for fn in files:
                    full = os.path.join(root, fn)
                    arc = os.path.relpath(full, tmp).replace("\\", "/")
                    z.write(full, arc)
        return out_path
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def measure_one(word, path, content_w_pt):
    d = word.Documents.Open(path, ReadOnly=True)
    time.sleep(0.2)
    try:
        chars = d.Range().Characters
        xs = []
        for ci in range(1, chars.Count + 1):
            try:
                c = chars(ci)
                t = c.Text
                if t in ("\r", "\x07"):
                    continue
                xs.append((t,
                           float(c.Information(5)),
                           float(c.Information(6)),
                           float(c.Font.Size if c.Font.Size else 0)))
            except Exception:
                continue
    finally:
        try: d.Close(SaveChanges=False)
        except: pass
    if not xs: return None
    # Group by line
    lines_buckets = {}
    for t, x, y, sz in xs:
        ykey = round(y, 0)
        lines_buckets.setdefault(ykey, []).append((t, x, y, sz))
    out = []
    for ykey in sorted(lines_buckets.keys()):
        items = sorted(lines_buckets[ykey], key=lambda v: v[1])
        advs = []
        for i in range(len(items) - 1):
            advs.append((items[i][0], round(items[i+1][1] - items[i][1], 3),
                         items[i][3]))
        # Yak analysis
        yaks = []
        for i, (ch, adv, sz) in enumerate(advs):
            if ch in YAKUMONO:
                ratio = adv / sz if sz > 0 else None
                yaks.append({
                    "pos": i,
                    "ch": ch,
                    "adv": round(adv, 3),
                    "font_size": sz,
                    "ratio": round(ratio, 4) if ratio else None,
                    "compression_pt": round(sz - adv, 3) if sz else 0,
                    "is_compressed": (ratio is not None and ratio < 0.99),
                })
        # actual line width: sum of advs + last char's font size (approx)
        line_x_start = items[0][1]
        line_x_last = items[-1][1]
        line_width = (line_x_last - line_x_start) + items[-1][3]
        # natural_full = sum(font_sizes) for all chars (assume CJK fullwidth)
        # — this is approx, ASCII digits would be half-width but our probes are CJK
        natural_full = sum(it[3] for it in items)
        # observed compression total
        total_comp = sum(y["compression_pt"] for y in yaks if y["is_compressed"])
        out.append({
            "y": ykey,
            "n_chars": len(items),
            "advances_first15": [(t, round(a, 2), round(s, 1)) for t, a, s in advs[:15]],
            "yaks": yaks,
            "natural_full_pt": round(natural_full, 3),
            "line_width_actual_pt": round(line_width, 3),
            "n_yak": len(yaks),
            "n_yak_compressed": sum(1 for y in yaks if y["is_compressed"]),
            "total_compression_pt": round(total_comp, 3),
            "content_w_pt": content_w_pt,
            "implied_natural_pre_compression": round(line_width + total_comp, 3),
            "overflow_vs_content_w": round((line_width + total_comp) - content_w_pt, 3),
        })
    return out


# Probe templates: each ends in `…次` or similar non-yak so we can measure
# all yak in line. Inject specific yak counts at controlled positions.
PROBES = {
    # Vary n_yak: each yak between CJK (Mech 1 wouldn't fire on B→CJK)
    "P_yak1_short":  "卸売市場法第６条第１項（第14条と同法第６条第１項）次の文",        # 2 yak (`（` `）`)
    "P_yak2_long":   "卸売市場法第６条第１項（第14条と同法第６条第１項）の規定により、次の文章を続け",    # 3 yak (+ `、`)
    "P_yak3_overflow":  "卸売市場法第６条第１項（第14条において準用する同法第６条第１項）の規定により、中央卸売市場（地方卸売市場）に係る認定事項の変更について認定を受けたいので、次の",   # ~6 yak
    # short probes — fits in 1 line, no overflow
    "P_yak2_no_overflow": "市場法（１項）規定",     # 2 yak, very short
    "P_yak2_padded":  "卸売市場法第６条第１項（第14条同法第６条第１項）と次の文章による情報定義",  # 2 yak, line-near-full
}


def main():
    word = w32.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    out = {}
    # 7f272a's content_w
    PAGE_W_TW = 11906; MARGIN_LR_TW = 1304
    content_w_pt = (PAGE_W_TW - 2 * MARGIN_LR_TW) / 20.0
    print(f"content_w = {content_w_pt}pt")

    try:
        for label, text in PROBES.items():
            for jc in ["left", "both"]:
                try:
                    pname = f"{label}_jc{jc}"
                    print(f"\n=== {pname} ===")
                    print(f"  text: {text[:60]!r}")
                    p = clone_with_doc(pname, text, jc, PAGE_W_TW, MARGIN_LR_TW)
                    lines = measure_one(word, p, content_w_pt)
                    if lines is None:
                        out[pname] = {"error": "no measurement"}
                        continue
                    out[pname] = {
                        "text": text, "jc": jc, "lines": lines,
                    }
                    for li, ln in enumerate(lines, start=1):
                        print(f"  L{li}: n={ln['n_chars']} natural~{ln['natural_full_pt']} "
                              f"actual={ln['line_width_actual_pt']} "
                              f"natural_implied={ln['implied_natural_pre_compression']} "
                              f"overflow={ln['overflow_vs_content_w']:+.2f} "
                              f"n_yak={ln['n_yak']} comp={ln['n_yak_compressed']} "
                              f"total_comp={ln['total_compression_pt']}")
                        for y in ln["yaks"]:
                            mark = "*" if y["is_compressed"] else " "
                            print(f"     {mark} pos={y['pos']:2d} {y['ch']!r} adv={y['adv']} ratio={y['ratio']} comp_pt={y['compression_pt']}")
                except Exception as e:
                    out[f"{label}_jc{jc}"] = {"error": str(e)}
                    print(f"  ERR: {e}")
    finally:
        try: word.Quit()
        except: pass

    os.makedirs(os.path.dirname(RESULT_PATH), exist_ok=True)
    with open(RESULT_PATH, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"\nWrote {RESULT_PATH}")


if __name__ == "__main__":
    main()
