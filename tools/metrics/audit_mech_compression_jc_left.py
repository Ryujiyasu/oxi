"""Cross-doc audit: how prevalent is yakumono compression on jc=left
paragraphs in the kern-active baseline?

Per Session 51 R0 (yakumono_kern_trigger) + Q6 alignment-agnostic:
  Mech 1 fires under any alignment when <w:kern> is in docDefaults.
  jc=left + kern → still expect Mech 1 compression for Type A/B/C
  trigger pairs.

Steps:
  1. Scan all 437 baseline docs for:
     - styles.xml has <w:kern>  (Mech 1 enabled at doc level)
     - document has at least one para with jc=left or no jc
  2. Pick 5 representative docs (varying yakumono density / topic).
  3. For each, measure first 3 paragraphs per-char via Word COM.
  4. Tag: did Word compress any yakumono? (advance < fontSize × 0.7)
  5. Summary table.
"""
import json, re, sys, time, zipfile, random, os
from pathlib import Path
from collections import Counter
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOCX_DIR = Path(r"C:\Users\ryuji\oxi-1\tools\golden-test\documents\docx")
RESULT_PATH = Path(r"C:\Users\ryuji\oxi-1\pipeline_data\mech_audit_jc_left.json")
YAKUMONO = set("「」（）［］【】〔〕、。，．" "'")
TYPE_C_NONCOMPRESS = set("ー―・：；！？／＼")


def scan_doc(p: Path):
    """Return dict with kern_in_styles, jc_left_paras, n_paras_with_yak,
    or None if can't read."""
    try:
        with zipfile.ZipFile(p) as z:
            try:
                docxml = z.read("word/document.xml").decode("utf-8")
            except KeyError:
                return None
            try:
                styles = z.read("word/styles.xml").decode("utf-8")
            except KeyError:
                styles = ""
    except Exception:
        return None

    # docDefaults kern presence (case-insensitive on attr value)
    kern_match = re.search(
        r'<w:docDefaults>.*?<w:rPrDefault>.*?<w:rPr>.*?<w:kern\b',
        styles, re.S
    )
    has_kern = kern_match is not None

    # Count paragraphs with jc=left or no jc, and any with yakumono
    n_jc_left_or_none = 0
    n_with_yak = 0
    for m in re.finditer(r'<w:p\b[^>]*>(.*?)</w:p>', docxml, re.S):
        body = m.group(1)
        ppr_m = re.search(r'<w:pPr\b[^>]*>(.*?)</w:pPr>', body, re.S)
        ppr = ppr_m.group(1) if ppr_m else ""
        jc_m = re.search(r'<w:jc w:val="([^"]+)"', ppr)
        jc_val = jc_m.group(1) if jc_m else None
        if jc_val in (None, "left", "start"):
            n_jc_left_or_none += 1
            txt = ''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', body))
            if any(c in YAKUMONO for c in txt):
                n_with_yak += 1
    return {
        "has_kern": has_kern,
        "n_jc_left_or_none": n_jc_left_or_none,
        "n_with_yak_jc_left": n_with_yak,
    }


def select_candidates():
    """Return docs that have kern + ≥3 jc=left/none paras with yak."""
    out = []
    docs = sorted(DOCX_DIR.glob("*.docx"))
    print(f"Scanning {len(docs)} baseline docs...", file=sys.stderr)
    for p in docs:
        info = scan_doc(p)
        if not info: continue
        if not info["has_kern"]: continue
        if info["n_with_yak_jc_left"] < 3: continue
        out.append((p.name, info["n_with_yak_jc_left"]))
    out.sort(key=lambda x: -x[1])
    return out


def measure_doc_first_paras(word, p: Path, max_paras=8):
    """Open doc, measure per-char advance for first max_paras paragraphs.
    Tag compression status per paragraph."""
    d = word.Documents.Open(str(p.resolve()), ReadOnly=True)
    out = []
    try:
        n_total = d.Paragraphs.Count
        # Find first 5 paragraphs that have yakumono content + jc=left/none
        para_idx_to_check = []
        for pi in range(1, min(n_total + 1, 80)):  # bound search
            try:
                para = d.Paragraphs(pi)
                if para.Alignment in (3,):  # 3=justify, skip
                    continue
                txt = (para.Range.Text or "")[:200]
                if not any(c in YAKUMONO for c in txt):
                    continue
                para_idx_to_check.append(pi)
                if len(para_idx_to_check) >= max_paras: break
            except Exception:
                continue

        for pi in para_idx_to_check:
            try:
                para = d.Paragraphs(pi)
                align = para.Alignment   # 0=left, 1=center, 2=right, 3=just
                txt_full = (para.Range.Text or "")
                # Per-char advance via Information(5)
                chars = para.Range.Characters
                xs = []
                for ci in range(1, min(chars.Count + 1, 60)):  # cap 60
                    try:
                        c = chars(ci)
                        t = c.Text
                        if t in ("\r", "\x07"):
                            continue
                        xs.append((t, float(c.Information(5)),
                                   float(c.Information(6)),
                                   float(c.Font.Size if c.Font.Size else 0)))
                    except Exception:
                        continue
                if not xs: continue
                # Group by line and compute advances
                xs_sorted = sorted(xs, key=lambda v: (v[2], v[1]))
                advs = []
                for i in range(len(xs_sorted) - 1):
                    a, b = xs_sorted[i], xs_sorted[i+1]
                    if abs(a[2] - b[2]) > 0.5:  # different line
                        continue
                    advs.append((a[0], round(b[1] - a[1], 3), a[3]))
                # Detect compressed yakumono: advance < fontSize × 0.7
                comp_yak = []
                full_yak = []
                for ch, adv, sz in advs:
                    if ch in YAKUMONO and sz > 0:
                        if adv < sz * 0.7:
                            comp_yak.append((ch, adv, sz))
                        else:
                            full_yak.append((ch, adv, sz))
                out.append({
                    "para_idx": pi,
                    "alignment_int": align,
                    "alignment": ["left","center","right","justify"][align] if align in (0,1,2,3) else f"?{align}",
                    "text_preview": txt_full[:50].replace("\r","\\r"),
                    "n_chars_measured": len(advs),
                    "n_yak_total": len(comp_yak) + len(full_yak),
                    "n_yak_compressed": len(comp_yak),
                    "compressed_yak_examples": [(c, round(a, 2), round(s, 1)) for c, a, s in comp_yak[:5]],
                    "compression_detected": len(comp_yak) > 0,
                })
            except Exception as e:
                out.append({"para_idx": pi, "error": str(e)})
    finally:
        try: d.Close(SaveChanges=0)
        except: pass
    return out


def main():
    print("Step 1: scan baseline for kern-active + jc=left/none docs...", file=sys.stderr)
    candidates = select_candidates()
    print(f"\nFound {len(candidates)} docs with kern + jc=left/none + yakumono content")
    print(f"Top 15 by yak-content count:")
    for n, count in candidates[:15]:
        print(f"  {count:>4}  {n}")

    # Pick 5 with diverse densities
    if len(candidates) >= 5:
        # Top, mid, low + 2 random
        sample = [candidates[0]]                        # most yak
        sample.append(candidates[len(candidates)//4])    # 75% percentile
        sample.append(candidates[len(candidates)//2])    # median
        sample.append(candidates[3*len(candidates)//4])  # 25% percentile
        sample.append(candidates[-1])                    # least
        # Dedupe
        seen = set()
        sampled = []
        for n, c in sample:
            if n not in seen:
                seen.add(n); sampled.append((n, c))
    else:
        sampled = candidates

    print(f"\n=== Sampling {len(sampled)} docs ===")
    for n, c in sampled:
        print(f"  {n} ({c} jc=left/none paras with yak)")

    print("\nStep 2: per-char measurement of first 5-8 paragraphs each...", file=sys.stderr)
    word = w32.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    results = {}
    try:
        for n, c in sampled:
            print(f"\n[{n}]", file=sys.stderr)
            try:
                p = DOCX_DIR / n
                results[n] = {
                    "scan_yak_para_count": c,
                    "measurements": measure_doc_first_paras(word, p, max_paras=8),
                }
            except Exception as e:
                results[n] = {"error": str(e)}
    finally:
        try: word.Quit()
        except: pass

    RESULT_PATH.parent.mkdir(parents=True, exist_ok=True)
    with open(RESULT_PATH, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)

    # Print summary
    print("\n=== Audit summary ===\n")
    for doc, info in results.items():
        if "error" in info:
            print(f"{doc[:50]:50s} ERROR: {info['error']}")
            continue
        meas = info["measurements"]
        n_compressed = sum(1 for m in meas if isinstance(m, dict) and m.get("compression_detected"))
        n_total = len([m for m in meas if isinstance(m, dict) and "compression_detected" in m])
        n_yak_compressed = sum(m.get("n_yak_compressed", 0) for m in meas if isinstance(m, dict))
        n_yak_total = sum(m.get("n_yak_total", 0) for m in meas if isinstance(m, dict))
        print(f"{doc[:50]:50s} paras_with_compress={n_compressed}/{n_total}  "
              f"yak_compressed={n_yak_compressed}/{n_yak_total}")
        for m in meas:
            if not isinstance(m, dict) or "error" in m: continue
            tag = "✓ COMPRESSED" if m["compression_detected"] else "  no compress"
            print(f"  p{m['para_idx']:3d} {m['alignment']:>8s} {tag}  yak={m['n_yak_compressed']}/{m['n_yak_total']}  ex={m['compressed_yak_examples']}  txt={m['text_preview']!r}")


if __name__ == "__main__":
    main()
