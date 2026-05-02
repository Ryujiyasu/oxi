"""
Measure per-char advance on body paragraphs containing consecutive yakumono
in no-kern baseline docs (R32 premise validation).

For each of 5 picked docs:
  1. Open in Word COM
  2. Walk every body paragraph
  3. For each pair of CONSECUTIVE yakumono characters (e.g. 」「), measure
     the X position of each char and compute per-char advance.
  4. Compare to natural width (≈ fontSize × 1.0pt for fullwidth CJK).
  5. Compression detected if advance < fontSize × 0.75.

Expected (R32 hypothesis "kern alone is trigger"):
  - All compressions are absent (every advance ≈ fontSize)
  - If ANY doc compresses → R32 design WRONG.

Picked from `no_kern_candidates.json`:
  - ruby_text_lineheight_11.docx (19 yakumono)
  - paragraph_spacing_grid_04.docx (17)
  - drop_cap_japanese_17.docx (15)
  - page_break_paragraph_spacing.docx (15)
  - textbox_wrap_07.docx (15)

Output: pipeline_data/r32_no_kern_yakumono_advance.json
"""
import os
import json
import time

import win32com.client
import pythoncom

DOCX_DIR = os.path.join(os.path.dirname(__file__), "..", "..", "pipeline_data", "docx")
OUT_JSON = os.path.join(os.path.dirname(__file__), "output", "r32_no_kern_yakumono_advance.json")

YAKUMONO = set(
    "（「『【〔｛〈《［"
    "）」』】〕｝〉》］、。，．"
    "・：；！？ー―／＼"
    "“”‘’"
)

PICKED = [
    # Re-picked 2026-05-02 by consecutive-pair count (compression candidates)
    "special_chars_spacing_01.docx",        # 8 pairs (highest density)
    "ruby_text_lineheight_11.docx",         # 6 pairs
    "complex_border_table_12.docx",         # 3 pairs
    "kinsoku_line_break_01.docx",           # 3 pairs (kinsoku-themed)
    "kinsoku_table_cell_overflow.docx",     # 3 pairs (kinsoku-themed)
]

RPC_REJECTED_CODES = {-2147418111, -2147023174, -2147023170}

def retry(fn, *args, retries=15, delay=0.3, **kwargs):
    last_exc = None
    for i in range(retries):
        try:
            return fn(*args, **kwargs)
        except Exception as e:
            last_exc = e
            code = e.args[0] if hasattr(e, "args") and len(e.args) >= 1 else None
            if code in RPC_REJECTED_CODES or "rejected" in str(e).lower():
                pythoncom.PumpWaitingMessages()
                time.sleep(delay * (1.3 ** i))
                continue
            raise
    raise last_exc


def find_yakumono_pairs(text):
    """Yield (i, char_pair) where text[i] and text[i+1] are both yakumono."""
    pairs = []
    for i in range(len(text) - 1):
        a, b = text[i], text[i + 1]
        if a in YAKUMONO and b in YAKUMONO:
            pairs.append((i, a + b))
    return pairs


def measure_char_x(wdoc, para, idx_in_para):
    """Measure X position of paragraph's char at idx_in_para."""
    rng = wdoc.Range(para.Range.Start + idx_in_para,
                     para.Range.Start + idx_in_para + 1)
    try:
        return round(rng.Information(5), 4)
    except Exception:
        return None


def get_run_font_size(para, idx_in_para):
    """Best-effort font size at character index — find first run that
    contains the index."""
    try:
        rng = para.Range.Document.Range(para.Range.Start + idx_in_para,
                                          para.Range.Start + idx_in_para + 1)
        size = rng.Font.Size
        return float(size) if size is not None else None
    except Exception:
        return None


def measure_doc(word, path):
    wdoc = retry(lambda: word.Documents.Open(path, ReadOnly=True))
    try:
        wdoc.Repaginate()
        time.sleep(0.05)
        results = []

        for pi in range(1, wdoc.Paragraphs.Count + 1):
            try:
                p = wdoc.Paragraphs(pi)
                # Skip non-body paragraphs (header/footer/footnotes are separate)
                in_table = bool(p.Range.Information(12))  # wdWithInTable
                text = p.Range.Text  # includes paragraph mark
                pairs = find_yakumono_pairs(text)
                if not pairs:
                    continue
                for idx, pair in pairs:
                    x_a = measure_char_x(wdoc, p, idx)
                    x_b = measure_char_x(wdoc, p, idx + 1)
                    x_c = measure_char_x(wdoc, p, idx + 2) if idx + 2 < len(text) else None
                    sz = get_run_font_size(p, idx)
                    advance_a = (x_b - x_a) if (x_b is not None and x_a is not None) else None
                    advance_b = (x_c - x_b) if (x_c is not None and x_b is not None) else None
                    natural_w = sz if sz else None
                    is_compressed_a = (advance_a is not None and natural_w is not None
                                       and advance_a < natural_w * 0.75)
                    results.append({
                        "para_idx": pi,
                        "char_idx": idx,
                        "pair": pair,
                        "in_table": in_table,
                        "x_a": x_a, "x_b": x_b, "x_c": x_c,
                        "advance_a": round(advance_a, 4) if advance_a is not None else None,
                        "advance_b": round(advance_b, 4) if advance_b is not None else None,
                        "font_size_pt": sz,
                        "natural_w_pt": natural_w,
                        "compressed_a": is_compressed_a,
                    })
            except Exception as e:
                results.append({"para_idx": pi, "err": str(e)})
        return results
    finally:
        wdoc.Close(False)


def main():
    word = win32com.client.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    time.sleep(2.0)
    try:
        retry(lambda: word.Documents.Count)
    except Exception:
        pass

    output = {}
    try:
        for fname in PICKED:
            path = os.path.join(DOCX_DIR, fname)
            if not os.path.exists(path):
                print(f"  SKIP {fname}: not found")
                continue
            print(f"\n=== {fname} ===")
            try:
                records = measure_doc(word, path)
            except Exception as e:
                print(f"  ERR: {e}")
                output[fname] = {"error": str(e)}
                continue

            output[fname] = records
            n_pairs = sum(1 for r in records if "pair" in r)
            n_compressed = sum(1 for r in records if r.get("compressed_a"))
            print(f"  {n_pairs} yakumono pairs measured, {n_compressed} compressed (advance < 0.75 × font)")

            # Print first 10 detailed
            shown = 0
            for r in records:
                if "pair" not in r:
                    continue
                if shown >= 12:
                    break
                shown += 1
                pair = r["pair"]
                adv = r["advance_a"]
                sz = r["font_size_pt"]
                ratio = (adv / sz) if (adv is not None and sz) else None
                tag = "COMPR" if r.get("compressed_a") else "natur"
                tbl = " (in_table)" if r.get("in_table") else ""
                ratio_str = f"{ratio:.2f}" if ratio is not None else "?"
                print(f"    P{r['para_idx']:3} idx={r['char_idx']:3} pair='{pair}' "
                      f"sz={sz}pt advance_a={adv} ratio={ratio_str} [{tag}]{tbl}")
    finally:
        with open(OUT_JSON, "w", encoding="utf-8") as f:
            json.dump(output, f, indent=2, ensure_ascii=False)
        print(f"\nSaved to {OUT_JSON}")
        try:
            word.Quit()
        except Exception as e:
            print(f"  (word.Quit failed: {e})")


if __name__ == "__main__":
    main()
