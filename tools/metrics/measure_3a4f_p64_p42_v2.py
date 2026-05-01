"""依頼 3 v2: 3a4f_p64 / p42 measurement using GoTo for page jump.

The full-doc scan v1 hit RPC death due to 2386 paragraphs.
v2: use Selection.GoTo wdGoToPage for direct page jump, then iterate
forward from that page only.
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


def measure_page(word, doc, target_page, max_paras=15):
    """Use Selection.GoTo to jump to page, measure paragraphs there."""
    sel = word.Selection
    # wdGoToPage = 1
    try:
        sel.GoTo(What=1, Which=1, Count=target_page)  # absolute page jump
    except Exception as e:
        return {"error": f"GoTo failed: {e}"}
    time.sleep(0.3)
    # Find the paragraph at this position
    start_para = doc.Range(sel.Start, sel.Start).Paragraphs(1)
    try:
        # Get paragraph index from doc start
        first_pi = None
        # Use Selection's Range
        r0 = sel.Range
        # Try to determine para index
        # Iterate from doc start: too slow. Use Range.End comparison
        # More robust: iterate paragraphs starting from this range
        results = []
        cur = start_para
        for offset in range(max_paras):
            try:
                rng = cur.Range
                page = int(rng.Information(3))
                if page > target_page:
                    break  # passed the page
                if page < target_page:
                    cur = cur.Next()
                    if cur is None:
                        break
                    continue
                # On target page
                text = rng.Text.strip()
                if not text:
                    cur = cur.Next()
                    if cur is None:
                        break
                    continue
                # Measure per-char
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
                        ratio = round(adv / sz, 3) if sz else None
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
                results.append({
                    "page": target_page,
                    "text": text,
                    "n_lines": len(line_data),
                    "lines": line_data,
                })
                cur = cur.Next()
                if cur is None:
                    break
            except Exception as e:
                return {"results": results, "error_at_offset": offset,
                        "error": str(e)}
        return {"results": results}
    except Exception as e:
        return {"error": str(e)}


def main():
    out = {}
    # Restart Word per page to dodge RPC death
    for page in [42]:  # only p42 — p64 already done
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False
        try:
            doc = word.Documents.Open(DOCX, ReadOnly=True)
            time.sleep(1.0)
            print(f"\n=== Page {page} ===", flush=True)
            res = measure_page(word, doc, page, max_paras=20)
            out[f"page_{page}"] = res
            doc.Close(SaveChanges=False)
        except Exception as e:
            out[f"page_{page}"] = {"error": str(e)}
        finally:
            try:
                word.Quit()
            except Exception:
                pass
            time.sleep(2.0)
    out_orig_p64 = {}
    if False:  # keep p64 from prior log
        pass
            if "error" in res and "results" not in res:
                print(f"  ERROR: {res['error']}", flush=True)
                continue
            results = res.get("results", [])
            print(f"  paragraphs measured on page {page}: {len(results)}",
                  flush=True)
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
                    print(f"    compressed yakumono ({len(yak_compressed)}):",
                          flush=True)
                    for a in yak_compressed:
                        print(f"      [{a['i']:3d}] {a['ch']!r:>3} "
                              f"({a['yakumono_class']}) "
                              f"prev={a['prev_ch']!r} next={a['next_ch']!r} "
                              f"adv={a['adv']} r={a['ratio']}",
                              flush=True)
        doc.Close(SaveChanges=False)
    finally:
        try:
            word.Quit()
        except Exception:
            pass

    with open(RESULT_PATH, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"\nWrote results to {RESULT_PATH}")


if __name__ == "__main__":
    main()
