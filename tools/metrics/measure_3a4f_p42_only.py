"""依頼 3 — measure 3a4f page 42 only (p64 already done in v2 log).

R31 loser at p42. Compare to p64 (winner) to identify what R31 changes
broke p42.
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
    "pipeline_data/3a4f_p42_per_char_2026-05-02.json")

YAKUMONO_A = set("（「『【〔｛〈《［")
YAKUMONO_B = set("）」』】〕｝〕》］、。，．—")
YAKUMONO_C = set("・：；！？ー―／＼")


def measure_page(word, doc, target_page, max_paras=15):
    sel = word.Selection
    try:
        sel.GoTo(What=1, Which=1, Count=target_page)
    except Exception as e:
        return {"error": f"GoTo failed: {e}"}
    time.sleep(0.5)
    start_para = doc.Range(sel.Start, sel.Start).Paragraphs(1)
    results = []
    cur = start_para
    for offset in range(max_paras):
        try:
            rng = cur.Range
            page = int(rng.Information(3))
            if page > target_page:
                break
            if page < target_page:
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
                    })
                except Exception:
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
                    yclass = ("A" if ch in YAKUMONO_A
                               else ("B" if ch in YAKUMONO_B
                                     else ("C" if ch in YAKUMONO_C else None)))
                    ratio = round(adv / sz, 3) if sz else None
                    next_ch = sorted_chars[i + 1]["ch"]
                    prev_ch = sorted_chars[i - 1]["ch"] if i > 0 else None
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
                line_data.append({"y": y, "n_chars": len(sorted_chars),
                                   "advances": advs})
            results.append({"page": target_page, "text": text,
                             "n_lines": len(line_data), "lines": line_data})
            cur = cur.Next()
            if cur is None:
                break
        except Exception as e:
            return {"results": results, "error_at_offset": offset,
                    "error": str(e)}
    return {"results": results}


def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    try:
        doc = word.Documents.Open(DOCX, ReadOnly=True)
        time.sleep(1.0)
        print("\n=== Page 42 ===", flush=True)
        res = measure_page(word, doc, 42, max_paras=20)
        results = res.get("results", [])
        print(f"  paragraphs: {len(results)}", flush=True)
        for para in results:
            yak_compressed = []
            for ln in para["lines"]:
                for a in ln["advances"]:
                    if (a["yakumono_class"]
                            and a["ratio"] is not None
                            and a["ratio"] < 0.85):
                        yak_compressed.append(a)
            if yak_compressed:
                print(f"\n  text: {para['text'][:60]!r}", flush=True)
                print(f"    compressed ({len(yak_compressed)}):", flush=True)
                for a in yak_compressed:
                    print(f"      [{a['i']:3d}] {a['ch']!r:>3} "
                          f"({a['yakumono_class']}) prev={a['prev_ch']!r} "
                          f"next={a['next_ch']!r} adv={a['adv']} "
                          f"r={a['ratio']}", flush=True)
        doc.Close(SaveChanges=False)
        out = {"page_42": res}
        with open(RESULT_PATH, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)
        print(f"\nWrote to {RESULT_PATH}", flush=True)
    finally:
        try:
            word.Quit()
        except Exception:
            pass


if __name__ == "__main__":
    main()
