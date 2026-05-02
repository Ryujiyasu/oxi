"""b35123 per-char COM measurement — BATCHED with Word restart per N paragraphs.

RPC server dies frequently on long iteration. Restart Word every 5 paragraphs.
"""
import os
import sys
import time
import json
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOC = os.path.abspath(
    "tools/golden-test/documents/docx/b35123fe8efc_tokumei_08_01.docx")
OUT = os.path.abspath("pipeline_data/b35123_per_char_2026-05-02.json")

YAKUMONO_A = set("（「『【〔｛〈《［")
YAKUMONO_B = set("）」』】〕｝〉》］、。，．—")


def cls(ch):
    if ch in YAKUMONO_A: return "A"
    if ch in YAKUMONO_B: return "B"
    return "X"


def measure_paragraph(d, pi):
    """Try to measure one paragraph, return data or None on error."""
    para = d.Paragraphs(pi)
    rng = para.Range
    page = int(rng.Information(3))
    if page > 1:
        return "PAGE_DONE"
    if page < 1:
        return None
    txt = rng.Text
    if not txt or txt == "\r":
        return None
    fs = float(rng.Font.Size) if rng.Font.Size > 0 else 12.0
    chars = rng.Characters
    char_data = []
    for ci in range(1, min(chars.Count, 80) + 1):
        try:
            c = chars(ci)
            ct = c.Text
            if ct in ("\r", "\x07"):
                continue
            x = float(c.Information(5))
            y = float(c.Information(6))
            char_data.append({
                "ch": ct, "x": x, "y": y,
                "cls": cls(ct),
                "size": float(c.Font.Size),
            })
        except Exception:
            continue
    return {
        "index": pi,
        "text_preview": txt[:50],
        "y": float(rng.Information(6)),
        "font_size": fs,
        "chars": char_data,
    }


def main():
    out = {"doc": DOC, "page": 1, "paragraphs": []}

    # Load existing partial data
    if os.path.exists(OUT):
        try:
            with open(OUT, encoding="utf-8") as f:
                old = json.load(f)
            existing = {p["index"] for p in old.get("paragraphs", [])}
            out["paragraphs"] = old["paragraphs"]
            print(f"Resuming, {len(existing)} paragraphs already done", flush=True)
        except Exception:
            existing = set()
    else:
        existing = set()

    # Determine total paragraphs (rough — we'll detect page boundary)
    # b35123 has 99 paragraphs total per earlier scan
    target_paragraphs = list(range(1, 100))
    skip_indices = existing

    BATCH_SIZE = 5
    pi_idx = 0
    while pi_idx < len(target_paragraphs):
        # Start Word fresh
        word = win32com.client.Dispatch("Word.Application")
        word.DisplayAlerts = False
        try:
            d = word.Documents.Open(DOC, ReadOnly=True)
            time.sleep(0.5)
            page_done = False
            for _ in range(BATCH_SIZE):
                if pi_idx >= len(target_paragraphs):
                    break
                pi = target_paragraphs[pi_idx]
                if pi in skip_indices:
                    pi_idx += 1
                    continue
                try:
                    res = measure_paragraph(d, pi)
                    if res == "PAGE_DONE":
                        page_done = True
                        pi_idx = len(target_paragraphs)  # exit outer loop
                        break
                    if res is not None:
                        out["paragraphs"].append(res)
                        yak_count = sum(1 for c in res["chars"] if c["cls"] in ("A","B"))
                        if yak_count > 0:
                            print(f"p{pi}: {yak_count} yak chars, text={res['text_preview']!r}", flush=True)
                    pi_idx += 1
                except Exception as e:
                    print(f"p{pi}: ERR {e}", flush=True)
                    # Save partial result + restart Word
                    break
            try:
                d.Close(SaveChanges=False)
            except: pass
            if page_done:
                break
        finally:
            try: word.Quit()
            except: pass
            time.sleep(2)
        # Save partial result every batch
        with open(OUT, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2, default=str)

    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2, default=str)
    print(f"\nWrote {OUT}, {len(out['paragraphs'])} paragraphs", flush=True)


if __name__ == "__main__":
    main()
