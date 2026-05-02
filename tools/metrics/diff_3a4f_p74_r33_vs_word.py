"""
Task C: 3a4f_p74 全行 R33 actual vs Word actual diff.

Inputs:
  Word actual per-char: oxi-3/pipeline_data/3a4f_p74_per_char_2026-05-02.json
  R33 implementation: ported from crates/oxidocs-core/src/layout/mod.rs
                     break_into_lines yakumono compression block (line 4181-4264)

Output: per-paragraph, per-line, per-char table:
  index | char | Word ratio | Word compressed | R33 compressed | match?

Goal: identify mismatch patterns where R33 vs Word disagree.
The user noted: paragraph 9 line 3 pos 87 （ was Mech-2-compressed by Word
but is NOT adjacent to a Mech-1 hit → partial hypothesis insufficient.

R33 compression decision (port of mod.rs:4181-4264):
  Stage 1 — yakumono_compressed[i] (alignment-agnostic):
    YAKUMONO_CLOSING = {）」』〕】》〙〗｝］、。，．}
    YAKUMONO_OPENING = {（「『〔【《〘〖｛［}
    YAKUMONO_TRIGGER = openers ∪ closers ∪ {・：；}
    For each i:
      if closing(c[i]) and i+1 < n and trigger(c[i+1]): v[i] = True
      elif opening(c[i]) and i > 0 and trigger(c[i-1]) and not v[i-1]: v[i] = True

  Stage 2 — apply (yakumono_enabled gated by kern/list_marker/chars+pair):
    is_opening_bracket = c in {（「『〔【《〈｛［}
    if v[i] and not is_opening_bracket:
      char_width *= 0.5  ← Mech 1 hit

  Stage 3 — expand-pair-rule (alignment-agnostic):
    is_yakumono_any = c in {（）「」『』〔〕【】《》〈〉｛｝［］、。，．}
    if is_yakumono_any:
      prev_compressed = v[i-1] if i > 0 else False
      next_compressed = v[i+1] if i+1 < n else False
      if (prev_compressed or next_compressed) and not is_opening_bracket:
        char_width *= 0.5  ← expand-pair

  Stage 4 — Mech 2 standalone hack (Justify/Distribute only):
    elif c in {、。，．}:
      if Justify or Distribute:
        prev_non_tr = ... and next_non_tr = ...
        if both: char_width *= 0.583  ← Mech 2 hack

  Stage 5 — Line-start ・/、/。 demand-driven (compat>=15):
    [skipped here — orthogonal to the question, only fires on first char of line]

For 3a4f_p74 we have:
  - kern in docDefaults? unknown without parsing styles, but likely True
    (3a4f is in the 109 docs with docDefaults kern per audit)
  - alignment? assume Justify (typical for kyodokenkyuyoushiki-style docs);
    fall back to read from docx if needed.
"""
import json
import os
import sys
import zipfile
import re

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

WORD_JSON = r"C:\Users\ryuji\oxi-3\pipeline_data\3a4f_p74_per_char_2026-05-02.json"
DOCX_PATH = r"C:\Users\ryuji\oxi-main\pipeline_data\golden_per_page\3a4f9fbe1a83_001620506_p74.docx"
OUT_JSON = os.path.join(os.path.dirname(__file__), "output", "ra2_3a4f_p74_r33_vs_word.json")
os.makedirs(os.path.dirname(OUT_JSON), exist_ok=True)


YAKUMONO_CLOSING = set("）」』〕】》〙〗｝］、。，．")
YAKUMONO_OPENING = set("（「『〔【《〘〖｛［")
YAKUMONO_TRIGGER = YAKUMONO_OPENING | YAKUMONO_CLOSING | set("・：；")
YAKUMONO_ANY = set("（）「」『』〔〕【】《》〈〉｛｝［］、。，．")
IS_OPENING_BRACKET = set("（「『〔【《〈｛［")
COMMA_PERIOD = set("、。，．")


def stage1_yakumono_compressed(text):
    """v[i] = True if char i is a Mech 1 hit (closing+next-trigger or opening+prev-trigger)."""
    n = len(text)
    v = [False] * n
    for i in range(n):
        c = text[i]
        if c in YAKUMONO_CLOSING:
            if i + 1 < n and text[i + 1] in YAKUMONO_TRIGGER:
                v[i] = True
        elif c in YAKUMONO_OPENING:
            if i > 0 and text[i - 1] in YAKUMONO_TRIGGER and not v[i - 1]:
                v[i] = True
    return v


def r33_decide_compressed(text, alignment="justify", kern_present=True):
    """Return list of dicts: {ch, mech1, expand_pair, mech2_standalone, compressed_r33}."""
    if not kern_present:
        return [{"ch": c, "mech1": False, "expand_pair": False,
                 "mech2_standalone": False, "compressed_r33": False} for c in text]

    v = stage1_yakumono_compressed(text)
    n = len(text)
    is_justify = alignment.lower() in ("justify", "distribute", "both")
    out = []
    for i in range(n):
        c = text[i]
        is_opening_b = c in IS_OPENING_BRACKET
        mech1 = v[i] and not is_opening_b
        expand_pair = False
        mech2_standalone = False
        if not mech1:
            if c in YAKUMONO_ANY:
                prev_compressed = v[i - 1] if i > 0 else False
                next_compressed = v[i + 1] if i + 1 < n else False
                if (prev_compressed or next_compressed) and not is_opening_b:
                    expand_pair = True
                elif c in COMMA_PERIOD and is_justify:
                    prev_non_tr = i == 0 or text[i - 1] not in YAKUMONO_TRIGGER
                    next_non_tr = i + 1 >= n or text[i + 1] not in YAKUMONO_TRIGGER
                    if prev_non_tr and next_non_tr:
                        mech2_standalone = True
        compressed = mech1 or expand_pair or mech2_standalone
        out.append({
            "ch": c,
            "mech1": mech1,
            "expand_pair": expand_pair,
            "mech2_standalone": mech2_standalone,
            "compressed_r33": compressed,
        })
    return out


def load_word_data():
    with open(WORD_JSON, encoding="utf-8") as f:
        return json.load(f)


def get_alignment_from_docx(path):
    """Read first body paragraph alignment from docx XML (rough heuristic)."""
    try:
        with zipfile.ZipFile(path) as z:
            with z.open("word/document.xml") as f:
                data = f.read().decode("utf-8")
        # Find first jc value
        m = re.search(r'<w:jc\s+w:val="([^"]+)"', data)
        return m.group(1) if m else "left"
    except Exception:
        return "unknown"


def main():
    word = load_word_data()
    paras = word["page_74"]
    print(f"Loaded {len(paras)} paragraphs from p74 Word per-char data")

    alignment = get_alignment_from_docx(DOCX_PATH)
    print(f"3a4f_p74 alignment (first jc): {alignment}")

    # We assume kern is present (3a4f is in the 109 docs with docDefaults kern).
    # Verify by inspecting styles.xml.
    kern_present = check_kern(DOCX_PATH)
    print(f"3a4f_p74 docDefaults kern: {kern_present}\n")

    summary = []
    mismatches_total = 0
    matches_total = 0

    for pi, para in enumerate(paras):
        text = para.get("text", "")
        if not text:
            continue
        print(f"\n=== Paragraph {pi+1} (n_lines={para.get('n_lines')}) ===")
        print(f"   text: {text[:80]}{'...' if len(text) > 80 else ''}")

        # Build word-side compressed flags by concatenating advances across lines
        word_chars = []
        for line_idx, ln in enumerate(para.get("lines", [])):
            for adv in ln.get("advances", []):
                word_chars.append({
                    "line": line_idx + 1,
                    "i": adv["i"],
                    "ch": adv["ch"],
                    "ratio": adv.get("ratio"),
                    "compressed_word": adv.get("compressed_word"),
                    "yakumono_class": adv.get("yakumono_class"),
                })

        # Word's text per char (paragraph-level)
        word_text = "".join(c["ch"] for c in word_chars)
        if word_text != text[:len(word_text)]:
            # Sometimes text includes trailing chars that are off-page; tolerate prefix match
            pass

        # Run R33 logic on the same text
        r33_out = r33_decide_compressed(word_text, alignment=alignment,
                                         kern_present=kern_present)

        # Diff per char
        n = len(word_chars)
        para_summary = {"para_idx": pi + 1, "alignment": alignment, "rows": [], "n_total": n,
                        "n_match": 0, "n_mismatch": 0,
                        "mismatch_word_compresses_r33_doesnt": [],
                        "mismatch_r33_compresses_word_doesnt": []}
        for i in range(n):
            wd = word_chars[i]
            r33 = r33_out[i]
            wc = wd["compressed_word"]
            rc = r33["compressed_r33"]
            match = wc == rc
            if match:
                para_summary["n_match"] += 1
                matches_total += 1
            else:
                para_summary["n_mismatch"] += 1
                mismatches_total += 1
                if wc and not rc:
                    para_summary["mismatch_word_compresses_r33_doesnt"].append(i)
                elif rc and not wc:
                    para_summary["mismatch_r33_compresses_word_doesnt"].append(i)
            para_summary["rows"].append({
                "i": i, "ch": wd["ch"],
                "line": wd["line"],
                "word_ratio": wd["ratio"],
                "word_compressed": wc,
                "r33_compressed": rc,
                "r33_mech1": r33["mech1"],
                "r33_expand_pair": r33["expand_pair"],
                "r33_mech2_standalone": r33["mech2_standalone"],
                "yakumono_class": wd["yakumono_class"],
                "match": match,
            })
        summary.append(para_summary)

        # Print mismatch summary
        n_mis = para_summary["n_mismatch"]
        if n_mis > 0:
            print(f"  MISMATCHES: {n_mis}/{n} chars disagree")
            print(f"    Word compresses, R33 doesn't: {len(para_summary['mismatch_word_compresses_r33_doesnt'])}")
            print(f"    R33 compresses, Word doesn't: {len(para_summary['mismatch_r33_compresses_word_doesnt'])}")
            for idx in (para_summary["mismatch_word_compresses_r33_doesnt"]
                        + para_summary["mismatch_r33_compresses_word_doesnt"])[:8]:
                row = para_summary["rows"][idx]
                ctx_l = word_text[max(0, idx-2):idx]
                ctx_r = word_text[idx+1:min(len(word_text), idx+3)]
                print(f"      i={idx} '{row['ch']}' (ctx '{ctx_l}[{row['ch']}]{ctx_r}') "
                      f"L{row['line']} word={row['word_compressed']}/r={row['word_ratio']} "
                      f"R33={row['r33_compressed']} (m1={row['r33_mech1']} ep={row['r33_expand_pair']} m2={row['r33_mech2_standalone']})")
        else:
            print(f"  ALL MATCH ({n} chars)")

    print(f"\n=== TOTAL ===")
    print(f"Paragraphs: {len(summary)}")
    print(f"Matches: {matches_total}, Mismatches: {mismatches_total}")
    print(f"Mismatch rate: {mismatches_total / max(1, matches_total + mismatches_total):.1%}")

    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump({"alignment": alignment, "kern_present": kern_present,
                   "matches_total": matches_total, "mismatches_total": mismatches_total,
                   "paragraphs": summary}, f, indent=2, ensure_ascii=False)
    print(f"\nSaved to {OUT_JSON}")


def check_kern(path):
    """Return True if word/styles.xml has <w:kern> in docDefaults rPrDefault."""
    try:
        with zipfile.ZipFile(path) as z:
            with z.open("word/styles.xml") as f:
                data = f.read().decode("utf-8")
        m = re.search(r"<w:docDefaults>.*?<w:rPrDefault>.*?<w:kern\s+w:val", data, re.DOTALL)
        return m is not None
    except Exception:
        return False


if __name__ == "__main__":
    main()
