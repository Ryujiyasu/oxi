"""依頼 (B): 3a4f_p74 per-char advance Word COM measurement.

3a4f_p74 = R33 big_loser (-0.0288 vs R32). Hypothesis: Oxi's Mech 2
Stage 3 distribution differs from Word's actual per-char compression
positions.

Strategy: GoTo page 74, iterate forward measuring each paragraph until
we leave page 74. Single Word session (page-jump avoids full-doc scan).
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
    "pipeline_data/3a4f_p74_per_char_2026-05-02.json")

YAKUMONO_A = set("（「『【〔｛〈《［")
YAKUMONO_B = set("）」』】〕｝〉》］、。，．—")
YAKUMONO_C = set("・：；！？ー―／＼")


def classify(ch):
    if ch in YAKUMONO_A:
        return "A"
    if ch in YAKUMONO_B:
        return "B"
    if ch in YAKUMONO_C:
        return "C"
    return None


def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    out = {"page_74": []}
    try:
        d = word.Documents.Open(DOCX, ReadOnly=True)
        time.sleep(1.0)
        sel = word.Selection
        sel.GoTo(What=1, Which=1, Count=74)  # wdGoToPage = 1
        time.sleep(0.5)
        cur = d.Range(sel.Start, sel.Start).Paragraphs(1)
        for offset in range(40):
            try:
                rng = cur.Range
                page = int(rng.Information(3))
                if page > 74:
                    break
                if page < 74:
                    cur = cur.Next()
                    if cur is None:
                        break
                    continue
                text = rng.Text.strip()
                if not text:
                    cur = cur.Next()
                    if cur is None:
                        break
                    continue
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
                        continue
                if not per_char:
                    cur = cur.Next()
                    if cur is None:
                        break
                    continue
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
                        ratio = round(adv / sz, 3) if sz else None
                        yclass = classify(ch)
                        next_ch = sorted_chars[i + 1]["ch"]
                        prev_ch = (sorted_chars[i - 1]["ch"]
                                    if i > 0 else None)
                        rule_match = "none"
                        if yclass == "A":
                            pc = classify(prev_ch) if prev_ch else None
                            if pc == "A":
                                rule_match = "A_after_A"
                        elif yclass == "B":
                            nc = classify(next_ch) if next_ch else None
                            if nc in ("A", "B"):
                                rule_match = f"B_before_{nc}"
                        advs.append({
                            "i": sorted_chars[i]["i"],
                            "ch": ch,
                            "prev_ch": prev_ch,
                            "next_ch": next_ch,
                            "adv": adv,
                            "size": sz,
                            "font": sorted_chars[i].get("font"),
                            "ratio": ratio,
                            "yakumono_class": yclass,
                            "rule_match": rule_match,
                            "compressed_word": (ratio is not None
                                                 and ratio < 0.85
                                                 and yclass is not None),
                        })
                    line_data.append({
                        "y": y,
                        "n_chars": len(sorted_chars),
                        "first_x": sorted_chars[0]["x"],
                        "last_x": sorted_chars[-1]["x"],
                        "advances": advs,
                    })
                out["page_74"].append({
                    "page": 74,
                    "text": text,
                    "n_lines": len(line_data),
                    "lines": line_data,
                })
                # Summary
                yak_compressed = []
                yak_not_compressed = []
                for ln in line_data:
                    for a in ln["advances"]:
                        if a["yakumono_class"]:
                            if a["compressed_word"]:
                                yak_compressed.append(a)
                            else:
                                yak_not_compressed.append(a)
                rule_YES = [a for a in yak_compressed
                             if a["rule_match"] != "none"]
                rule_NOT = [a for a in yak_not_compressed
                             if a["rule_match"] != "none"]
                print(f"\n  text: {text[:60]!r}", flush=True)
                print(f"    n_lines={len(line_data)} "
                      f"yak_total={len(yak_compressed) + len(yak_not_compressed)} "
                      f"compressed={len(yak_compressed)} "
                      f"M1_hits={len(rule_YES)} "
                      f"NOT_M1_compressed_via_M2={len(yak_compressed) - len(rule_YES)} "
                      f"M1_anomaly_NOT_compressed={len(rule_NOT)}",
                      flush=True)
                # Print all compressed chars
                for a in yak_compressed:
                    cls = "Mech1" if a["rule_match"] != "none" else "Mech2"
                    print(f"    [{cls}] [{a['i']:3d}] {a['ch']!r:>3} "
                          f"prev={a['prev_ch']!r} next={a['next_ch']!r} "
                          f"adv={a['adv']} r={a['ratio']}", flush=True)
                # Print uncompressed B chars (relevant for Mech 2 candidates)
                if rule_NOT:
                    print("    ANOMALY (M1 rule match but NOT compressed):",
                          flush=True)
                    for a in rule_NOT:
                        print(f"      [{a['i']:3d}] {a['ch']!r:>3} "
                              f"prev={a['prev_ch']!r} next={a['next_ch']!r} "
                              f"rule={a['rule_match']} adv={a['adv']}",
                              flush=True)
                cur = cur.Next()
                if cur is None:
                    break
            except Exception as e:
                print(f"  ERR offset={offset}: {e}", flush=True)
                break
        try:
            d.Close(SaveChanges=False)
        except Exception:
            pass
    finally:
        try:
            word.Quit()
        except Exception:
            pass
    with open(RESULT_PATH, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"\nWrote {RESULT_PATH}: page_74 = {len(out['page_74'])} paragraphs",
          flush=True)


if __name__ == "__main__":
    main()
