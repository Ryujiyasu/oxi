"""§4.7 R17 yakumono gate validation — per-char advance measurement
of 4 paragraphs (2 R17 winners, 2 R17 losers) to identify the exact
yakumono pairs Word compresses.

Goal: pin-point which yakumono pairs in real docs Word compresses,
so we can reverse-engineer the proper Oxi gate (replacing R17's
list_marker workaround).

Targets (per user 2026-05-02):
  ed025_p1 paragraph 12 (卸売市場法第４条第５項第５号...)
    R17 big_loser, fLCh=100, jc無し, numPr無し
  7f272a_p1 paragraph 12 (卸売市場法第６条第１項...)
    R17 big_loser, fLCh=100, jc=left, numPr無し
  683f_p2 paragraph 0 (５　前項に基づき著作権が...)
    R17 winner = Word does NOT compress, twip-only hanging, numPr無し
  3a4f_p23 paragraph 4 (また、労基法第３２条第２項...)
    R17 winner = Word does NOT compress, fLCh=100, numPr無し

Method: open each docx via Word COM, find target paragraphs by
matching opening text, measure per-char Information(5) (horizontal
position relative to page), compute advances. Identify chars where
advance ≠ natural width (10.5pt for default body font).
"""
import win32com.client
import os
import time
import json
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# (label, docx_path, target_text_prefix (15-30 chars to identify para))
TARGETS = [
    # Skip the two already-completed losers; re-run only for winners
    ("683f_p2_para0_winner",
     "tools/golden-test/documents/docx/683ffcab86e2_20230331_resources_open_data_contract_addon_00.docx",
     "前項に基づき著作権"),
    ("3a4f_p23_para4_winner",
     "tools/golden-test/documents/docx/3a4f9fbe1a83_001620506.docx",
     "労基法第３２条第２項"),
    # Also re-run losers in same script for combined output
    ("ed025_p1_para12_loser",
     "tools/golden-test/documents/docx/ed025cbecffb_index-23.docx",
     "卸売市場法第４条第５項第５号"),
    ("7f272a_p1_para12_loser",
     "tools/golden-test/documents/docx/7f272a2dfd3b_index-21.docx",
     "卸売市場法第６条第１項"),
]

# Yakumono characters to flag in output
YAKUMONO_A = set("（「『【〔｛〈《［")  # Type A — open
YAKUMONO_B = set("）」』】〕｝〉》］、。，．—")  # Type B — close/punct
YAKUMONO_C = set("・：；！？ー―／＼")  # Type C — non-compressing

RESULT_PATH = os.path.abspath(
    "pipeline_data/r17_per_char_advances_2026-05-02.json")


def find_paragraph_by_prefix(doc, prefix, max_paragraphs=2000):
    """Find paragraph containing prefix as substring (lenient)."""
    for pi in range(1, min(doc.Paragraphs.Count + 1, max_paragraphs + 1)):
        try:
            p = doc.Paragraphs(pi)
            text = p.Range.Text.strip()
            if prefix in text:
                return pi
        except Exception:
            continue
    return None


def measure_para(word, docx_path, prefix):
    d = word.Documents.Open(os.path.abspath(docx_path), ReadOnly=True)
    try:
        time.sleep(0.5)
        pi = find_paragraph_by_prefix(d, prefix)
        if pi is None:
            return {"error": f"paragraph starting with {prefix!r} not found"}
        p = d.Paragraphs(pi)
        rng = p.Range
        text = rng.Text
        page = int(rng.Information(3))
        # Per-char measurement
        chars = rng.Characters
        per_char = []
        for ci in range(1, chars.Count + 1):
            try:
                c = chars(ci)
                t = c.Text
                if t in ("\r", "\x07"):
                    continue
                per_char.append({
                    "i": ci,
                    "ch": t,
                    "x": round(float(c.Information(5)), 4),
                    "y": round(float(c.Information(6)), 4),
                    "font": c.Font.Name,
                    "size": c.Font.Size,
                })
            except Exception as e:
                per_char.append({"i": ci, "err": str(e)})
        # Compute advances grouped by line
        lines = {}
        for r in per_char:
            if "err" in r:
                continue
            y = r.get("y")
            if y is None:
                continue
            lines.setdefault(round(y, 1), []).append(r)
        line_data = []
        for y in sorted(lines.keys()):
            sorted_chars = sorted(lines[y], key=lambda r: r["x"])
            advs = []
            for i in range(len(sorted_chars) - 1):
                ch = sorted_chars[i]["ch"]
                adv = round(sorted_chars[i + 1]["x"]
                             - sorted_chars[i]["x"], 4)
                size = sorted_chars[i]["size"]
                expected_full = size  # CJK natural
                # Flag if compressed: adv < 0.85 × size for yakumono
                is_yak = (ch in YAKUMONO_A or ch in YAKUMONO_B
                           or ch in YAKUMONO_C)
                ratio = round(adv / size, 3) if size else None
                advs.append({
                    "i": sorted_chars[i]["i"],
                    "ch": ch,
                    "adv": adv,
                    "size": size,
                    "ratio": ratio,
                    "yakumono_class": (
                        "A" if ch in YAKUMONO_A
                        else ("B" if ch in YAKUMONO_B
                              else ("C" if ch in YAKUMONO_C else None))),
                    "compressed": (is_yak and ratio is not None
                                    and ratio < 0.85),
                })
            line_data.append({
                "y": y,
                "n_chars": len(sorted_chars),
                "first_x": sorted_chars[0]["x"] if sorted_chars else None,
                "advances": advs,
            })
        return {
            "paragraph_index": pi,
            "page": page,
            "text": text.strip(),
            "n_lines": len(line_data),
            "lines": line_data,
        }
    finally:
        d.Close(SaveChanges=False)


def main():
    results = {}
    # Restart Word per doc to avoid RPC death
    for label, path, prefix in TARGETS:
        print(f"\n=== {label} ===", flush=True)
        print(f"  doc: {path}", flush=True)
        print(f"  prefix: {prefix!r}", flush=True)
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False
        try:
            try:
                data = measure_para(word, path, prefix)
            except Exception as e:
                data = {"error": str(e)}
            results[label] = data
            if "error" in data:
                print(f"  ERROR: {data['error']}", flush=True)
                continue
            print(f"  found paragraph {data['paragraph_index']} on page "
                  f"{data['page']}", flush=True)
            print(f"  text: {data['text'][:80]!r}", flush=True)
            print(f"  n_lines: {data['n_lines']}", flush=True)
            for line in data["lines"]:
                yak_advs = [a for a in line["advances"]
                             if a["yakumono_class"] is not None]
                compressed = [a for a in yak_advs if a["compressed"]]
                if not yak_advs:
                    continue
                print(f"  L y={line['y']} n={line['n_chars']}: "
                      f"yakumono={len(yak_advs)} compressed={len(compressed)}",
                      flush=True)
                for a in yak_advs:
                    marker = "  ← COMPRESSED" if a["compressed"] else ""
                    print(f"    [{a['i']:3d}] {a['ch']!r:>4} "
                          f"({a['yakumono_class']}) adv={a['adv']} "
                          f"r={a['ratio']}{marker}", flush=True)
        finally:
            try:
                word.Quit()
            except Exception:
                pass
            time.sleep(1.5)

    os.makedirs(os.path.dirname(RESULT_PATH), exist_ok=True)
    with open(RESULT_PATH, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f"\nWrote results to {RESULT_PATH}")


if __name__ == "__main__":
    main()
