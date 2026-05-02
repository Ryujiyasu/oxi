"""依頼 3: 3a4f_p64 (R31 winner +0.032) と 3a4f_p42 (R31 loser -0.008)
の per-char advance 測定.

Goal: R31 が p64 で何を追加 compress したか specific pair 単位で確認、
Mech 1 (FINAL RULE) なのか Mech 2 なのか判定。

3a4f は kern=None (audit 確認済) なので、session 51 spec によれば
NO compression のはずだが、p23 では Mech 2 圧縮確認済。p64 でも
同様か、別パターンか。
"""
import win32com.client
import os
import time
import json
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOCX = os.path.abspath(
    "tools/golden-test/documents/docx/3a4f9fbe1a83_001620506.docx")
RESULT_PATH = os.path.abspath(
    "pipeline_data/3a4f_p64_p42_per_char_2026-05-02.json")

YAKUMONO_A = set("（「『【〔｛〈《［")
YAKUMONO_B = set("）」』】〕｝〉》］、。，．—")
YAKUMONO_C = set("・：；！？ー―／＼")


def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    results = {"page_64_paragraphs": [], "page_42_paragraphs": []}
    try:
        d = word.Documents.Open(DOCX, ReadOnly=True)
        time.sleep(0.8)
        # Find paragraphs on page 64 and page 42
        for target_page in [64, 42]:
            label = f"page_{target_page}_paragraphs"
            print(f"\n=== Scanning page {target_page} ===")
            paragraphs_on_page = []
            for pi in range(1, d.Paragraphs.Count + 1):
                try:
                    p = d.Paragraphs(pi)
                    page = int(p.Range.Information(3))
                    if page > target_page + 1:
                        break
                    if page == target_page:
                        text = p.Range.Text.strip()
                        if not text:
                            continue
                        paragraphs_on_page.append((pi, text))
                except Exception:
                    continue
            print(f"  found {len(paragraphs_on_page)} paragraphs on page "
                  f"{target_page}")

            # Measure each paragraph
            for pi, text in paragraphs_on_page:
                try:
                    p = d.Paragraphs(pi)
                    rng = p.Range
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
                                "size": c.Font.Size,
                                "font": c.Font.Name,
                            })
                        except Exception:
                            pass
                    # Group by line
                    lines = {}
                    for r in per_char:
                        lines.setdefault(round(r["y"], 1), []).append(r)
                    line_data = []
                    for y in sorted(lines.keys()):
                        sorted_chars = sorted(lines[y], key=lambda r: r["x"])
                        advs = []
                        for i in range(len(sorted_chars) - 1):
                            ch = sorted_chars[i]["ch"]
                            adv = round(sorted_chars[i + 1]["x"]
                                         - sorted_chars[i]["x"], 4)
                            sz = sorted_chars[i]["size"]
                            yclass = ("A" if ch in YAKUMONO_A
                                       else ("B" if ch in YAKUMONO_B
                                             else ("C" if ch in YAKUMONO_C
                                                   else None)))
                            ratio = (round(adv / sz, 3) if sz else None)
                            next_ch = sorted_chars[i + 1]["ch"]
                            prev_ch = (sorted_chars[i - 1]["ch"]
                                        if i > 0 else None)
                            advs.append({
                                "i": sorted_chars[i]["i"],
                                "ch": ch,
                                "prev_ch": prev_ch,
                                "next_ch": next_ch,
                                "adv": adv,
                                "size": sz,
                                "ratio": ratio,
                                "yakumono_class": yclass,
                            })
                        line_data.append({
                            "y": y,
                            "n_chars": len(sorted_chars),
                            "advances": advs,
                        })
                    para_record = {
                        "para_index": pi,
                        "page": target_page,
                        "text": text,
                        "n_lines": len(line_data),
                        "lines": line_data,
                    }
                    results[label].append(para_record)
                    # Print summary
                    yak_compressed = []
                    for ln in line_data:
                        for a in ln["advances"]:
                            if (a["yakumono_class"]
                                    and a["ratio"] is not None
                                    and a["ratio"] < 0.85):
                                yak_compressed.append(a)
                    if yak_compressed:
                        print(f"\n  [P{pi}] (text={text[:60]!r})")
                        print(f"    compressed yakumono ({len(yak_compressed)}):")
                        for a in yak_compressed:
                            print(f"      [{a['i']:3d}] {a['ch']!r:>3} "
                                  f"({a['yakumono_class']}) prev={a['prev_ch']!r} "
                                  f"next={a['next_ch']!r} adv={a['adv']} "
                                  f"r={a['ratio']}")
                    elif any(a["yakumono_class"]
                              for ln in line_data for a in ln["advances"]):
                        print(f"  [P{pi}] yakumono present, NONE compressed")
                except Exception as e:
                    print(f"  [P{pi}] ERR: {e}")
        d.Close(SaveChanges=False)
    finally:
        try:
            word.Quit()
        except Exception:
            pass
    with open(RESULT_PATH, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f"\nWrote results to {RESULT_PATH}")


if __name__ == "__main__":
    main()
