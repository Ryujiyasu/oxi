"""§4.7c Mech 3 — measure 7f272a_p1 P13 + sibling paragraphs per-line per-char.

Goal: gather data points for regression of compression amount against:
  candidates: slack-distribution, grid-snap, font-size const

For each measured line:
  - text characters + advances (Information(5))
  - n_yak, n_compressible, total_yak_compression
  - line_y, content_w (via doc PageWidth - margins)
  - font sizes per char
  - linePitch from docGrid
"""
import json, sys, time
from pathlib import Path
import win32com.client as w32
import zipfile, re

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOC = Path(r"C:\Users\ryuji\oxi-1\tools\golden-test\documents\docx\7f272a2dfd3b_index-21.docx")
OUT = Path(r"C:\Users\ryuji\oxi-1\pipeline_data\mech3_7f272a_perline.json")

# Type A/B yakumono (compressible per Mech 1/2/3)
TYPE_A = set("（「『【〔｛〈《［" "‘")
TYPE_B = set("）」』】〕｝〉》］、。，．" "—")
ALL_YAK = TYPE_A | TYPE_B


def parse_section_geometry(docx_path):
    with zipfile.ZipFile(docx_path) as z:
        xml = z.read("word/document.xml").decode("utf-8")
    pgsz = re.search(r'<w:pgSz w:w="(\d+)" w:h="(\d+)"', xml)
    pgmar = re.search(
        r'<w:pgMar w:top="(\d+)" w:right="(\d+)" w:bottom="(\d+)" w:left="(\d+)"',
        xml
    )
    grid = re.search(r'<w:docGrid[^/>]*w:linePitch="(\d+)"', xml)
    pw = int(pgsz.group(1)) / 20.0 if pgsz else None
    mr = int(pgmar.group(2)) / 20.0 if pgmar else 0
    ml = int(pgmar.group(4)) / 20.0 if pgmar else 0
    content_w = pw - mr - ml if pw else None
    line_pitch_tw = int(grid.group(1)) if grid else None
    return {
        "page_w_pt": pw,
        "margin_left_pt": ml,
        "margin_right_pt": mr,
        "content_w_pt": content_w,
        "linePitch_tw": line_pitch_tw,
        "linePitch_pt": (line_pitch_tw / 20.0) if line_pitch_tw else None,
    }


def measure_paragraph(word, doc, pi):
    para = doc.Paragraphs(pi)
    align = para.Alignment
    txt_full = para.Range.Text or ""
    chars = para.Range.Characters
    xs = []
    for ci in range(1, min(chars.Count + 1, 200)):
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
    if not xs:
        return None
    # Group by line
    lines = {}
    for t, x, y, sz in xs:
        # Bucket y by 0.5pt
        ykey = round(y, 0)
        lines.setdefault(ykey, []).append((t, x, y, sz))
    # Within each line, sort by x
    line_results = []
    for ykey in sorted(lines.keys()):
        items = sorted(lines[ykey], key=lambda v: v[1])
        # Advances
        advs = []
        for i in range(len(items) - 1):
            advs.append((items[i][0], round(items[i+1][1] - items[i][1], 3),
                         items[i][3]))
        # Last char advance unknown unless we have extent
        chars_text = "".join(it[0] for it in items)
        n_chars = len(items)
        # Yak analysis
        yaks = []
        for i, (ch, adv, sz) in enumerate(advs):
            if ch in ALL_YAK:
                ratio = adv / sz if sz > 0 else None
                yaks.append({
                    "pos": i,
                    "ch": ch,
                    "adv": round(adv, 3),
                    "font_size": sz,
                    "ratio": round(ratio, 4) if ratio else None,
                    "compressed": (ratio is not None and ratio < 0.95),
                })
        # Natural width estimate: each non-yak = font_size, each yak = font_size
        # Sum of advances tells us actual line width minus last char
        line_x_start = items[0][1]
        line_x_last = items[-1][1]
        last_adv = items[-1][3]  # assume last char fullwidth
        line_width_actual = (line_x_last - line_x_start) + last_adv
        # Sum of font sizes (natural CJK fullwidth assumption)
        natural = sum(it[3] for it in items)
        line_results.append({
            "y": ykey,
            "n_chars": n_chars,
            "text": chars_text[:80],
            "advances": [(t, round(a, 3), round(s, 1)) for t, a, s in advs],
            "yaks": yaks,
            "natural_estimate_pt": round(natural, 3),
            "line_width_actual_pt": round(line_width_actual, 3),
            "n_yak": len(yaks),
            "n_yak_compressed": sum(1 for y in yaks if y["compressed"]),
            "total_yak_compression": round(sum(
                (y["font_size"] - y["adv"]) for y in yaks if y["compressed"]
            ), 3),
            "line_x_start": round(line_x_start, 3),
            "line_x_last": round(line_x_last, 3),
        })
    return {
        "para_idx": pi,
        "alignment": align,
        "text": txt_full[:120].replace("\r", "\\r"),
        "lines": line_results,
    }


def main():
    geom = parse_section_geometry(DOC)
    print(f"Section: page_w={geom['page_w_pt']}pt margins=L{geom['margin_left_pt']}/R{geom['margin_right_pt']} "
          f"content_w={geom['content_w_pt']}pt linePitch={geom['linePitch_pt']}pt")

    # Target paragraphs: P13 (the wrap-fit example), surrounding contexts
    TARGET_P = [11, 13, 16, 18, 19, 21, 22, 25, 27, 28, 30, 32, 34]

    word = w32.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    out = {"geometry": geom, "paragraphs": {}}
    try:
        doc = word.Documents.Open(str(DOC.resolve()), ReadOnly=True)
        try:
            n_total = doc.Paragraphs.Count
            print(f"Word total paragraphs: {n_total}")
            for pi in TARGET_P:
                if pi > n_total: continue
                try:
                    r = measure_paragraph(word, doc, pi)
                    if r is None: continue
                    out["paragraphs"][f"P{pi}"] = r
                    print(f"\n=== P{pi} (align={r['alignment']}, lines={len(r['lines'])}) ===")
                    print(f"  text: {r['text'][:80]!r}")
                    for li, ln in enumerate(r["lines"], start=1):
                        comp_yaks = [y for y in ln["yaks"] if y["compressed"]]
                        print(f"  L{li}: n_chars={ln['n_chars']}  width={ln['line_width_actual_pt']}pt  natural={ln['natural_estimate_pt']}pt  "
                              f"slack={round(ln['natural_estimate_pt']-ln['line_width_actual_pt'],3)}  "
                              f"yak={ln['n_yak']} comp={ln['n_yak_compressed']} "
                              f"total_comp={ln['total_yak_compression']}pt")
                        for y in comp_yaks:
                            print(f"       {y['ch']!r} adv={y['adv']} sz={y['font_size']} ratio={y['ratio']}")
                except Exception as e:
                    print(f"P{pi} ERR: {e}")
                    out["paragraphs"][f"P{pi}"] = {"error": str(e)}
        finally:
            try: doc.Close(SaveChanges=0)
            except: pass
    finally:
        try: word.Quit()
        except: pass

    OUT.parent.mkdir(parents=True, exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"\nWrote {OUT}")


if __name__ == "__main__":
    main()
