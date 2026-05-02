"""R35 Phase 1: re-analyze existing R33-era per-char JSONs.

Data shape (d77a_p10, 7f272a_p1):
  {paragraphs: [{para, n_lines, n_chars, text, lines: [
      {n_char, y, x_start, x_end, chars: [
          {ch, next, x, adv, sz, yak, next_yak, is_cjk, next_cjk}
      ]}
  ]}]}

Goals (R35 brief):
  Q1: standalone B-class char (、。，．) — overflow なしで jc=both で圧縮?
  Q2: A-class (（「) は同条件で圧縮?
  Q3: 圧縮量 — 0.583x 固定 / slack-driven / 混在?
  Q4: jc=left でも同じ trigger?
  B1: autoSpaceDE が overflow 計算に含まれる?
  B2: yakumono advance ≠ ink width?
  B3: line-end implicit padding?
"""
import json
import os
import sys
from collections import defaultdict

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DATA_DIR = "pipeline_data/r35_data"
OUT = "pipeline_data/r35_phase1_findings_2026-05-02.json"

YAKUMONO_A = set("（「『【〔｛〈《［")
YAKUMONO_B = set("）」』】〕｝〉》］、。，．—")
TRIGGER_NEIGHBOR = YAKUMONO_A | YAKUMONO_B  # FINAL RULE neighbor classes


def cls(ch):
    if ch in YAKUMONO_A: return "A"
    if ch in YAKUMONO_B: return "B"
    return "X"


def line_natural_width(line, content_w):
    """Compute the natural (un-compressed) width of a line.
    Sum of (sz if yak else adv) for B-class chars where adv < sz.
    """
    n = 0.0
    for c in line["chars"]:
        sz = c.get("sz") or 12.0
        # If yak, use full sz as natural; else use observed adv
        if cls(c["ch"]) in ("A", "B"):
            n += sz  # natural full-width
        else:
            n += c.get("adv") or sz
    return n


def analyze_doc(label, path, out_root):
    """Walk paragraphs[].lines[].chars[] structure, classify each yak."""
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    paras = data.get("paragraphs", [])
    findings = {
        "label": label, "path": path,
        "n_lines": 0, "n_yak": 0, "n_compressed": 0,
        "compressed_records": [],
        "uncompressed_yak": 0,
        "by_neighbor": defaultdict(list),
        "by_class": defaultdict(int),
        "by_class_compressed": defaultdict(int),
        "lines_with_compression": [],
    }
    for p_idx, para in enumerate(paras):
        for l_idx, line in enumerate(para.get("lines", [])):
            findings["n_lines"] += 1
            chars = line["chars"]
            content_w = (line.get("x_end") or 0) - (line.get("x_start") or 0)
            natural = line_natural_width(line, content_w)
            slack = content_w - natural
            line_has_comp = False
            for c_idx, c in enumerate(chars):
                ch = c["ch"]
                cl = cls(ch)
                if cl == "X":
                    continue
                findings["n_yak"] += 1
                findings["by_class"][cl] += 1
                adv = c.get("adv") or 0
                sz = c.get("sz") or 12.0
                ratio = round(adv / sz, 4) if sz else None
                # Determine prev/next chars
                prev_ch = chars[c_idx - 1]["ch"] if c_idx > 0 else "^"
                next_ch = chars[c_idx + 1]["ch"] if c_idx + 1 < len(chars) else "$"
                pcls = cls(prev_ch)
                ncls = cls(next_ch)
                # FINAL RULE detection
                final_rule = "none"
                if cl == "A" and pcls == "A":
                    final_rule = "A_after_A"
                elif cl == "B" and ncls in ("A", "B"):
                    final_rule = f"B_before_{ncls}"
                rec = {
                    "doc": label,
                    "para": para["para"], "line": l_idx + 1,
                    "char_pos": c_idx,
                    "n_chars_line": len(chars),
                    "ch": ch, "class": cl,
                    "prev_ch": prev_ch, "prev_cls": pcls,
                    "next_ch": next_ch, "next_cls": ncls,
                    "adv": adv, "sz": sz, "ratio": ratio,
                    "content_w": round(content_w, 2),
                    "natural": round(natural, 2),
                    "slack": round(slack, 2),
                    "final_rule": final_rule,
                }
                if ratio is not None and ratio < 0.85:
                    findings["n_compressed"] += 1
                    findings["by_class_compressed"][cl] += 1
                    findings["compressed_records"].append(rec)
                    key = f"{cl}_after_{pcls}_before_{ncls}"
                    findings["by_neighbor"][key].append(ratio)
                    line_has_comp = True
                else:
                    findings["uncompressed_yak"] += 1
            if line_has_comp:
                findings["lines_with_compression"].append({
                    "para": para["para"], "line_idx": l_idx + 1,
                    "n_chars": len(chars),
                    "content_w": round(content_w, 2),
                    "natural": round(natural, 2),
                    "slack": round(slack, 2),
                })

    # Summarize
    summary = {
        "n_lines": findings["n_lines"],
        "n_yak": findings["n_yak"],
        "n_compressed": findings["n_compressed"],
        "uncompressed_yak": findings["uncompressed_yak"],
        "by_class": dict(findings["by_class"]),
        "by_class_compressed": dict(findings["by_class_compressed"]),
        "by_neighbor": {k: {"n": len(v), "min": min(v), "max": max(v),
                              "ratios": v[:8]}
                         for k, v in findings["by_neighbor"].items()},
        "lines_with_compression": findings["lines_with_compression"],
        "compressed_records_sample": findings["compressed_records"][:50],
    }
    out_root[label] = summary

    print(f"\n=== {label} ===")
    print(f"  lines={summary['n_lines']} yak={summary['n_yak']} "
          f"compressed={summary['n_compressed']} ({100*summary['n_compressed']/max(1,summary['n_yak']):.1f}%)")
    print(f"  by_class: {summary['by_class']}")
    print(f"  by_class_compressed: {summary['by_class_compressed']}")
    print(f"  Neighbor classes (compressed):")
    for k, v in sorted(summary["by_neighbor"].items()):
        print(f"    {k}: n={v['n']:>3} min={v['min']:.4f} max={v['max']:.4f} "
              f"ratios={v['ratios']}")

    # Slack analysis: was Word's Mech compressing on lines with negative slack?
    pos_slack = [l for l in findings["lines_with_compression"] if l["slack"] > 0]
    neg_slack = [l for l in findings["lines_with_compression"] if l["slack"] <= 0]
    print(f"  lines_with_compression: pos_slack={len(pos_slack)} "
          f"neg_or_zero_slack={len(neg_slack)}")
    if pos_slack:
        print(f"    POSITIVE slack lines compressed (B1/B3 hypothesis evidence):")
        for l in pos_slack[:10]:
            print(f"      para={l['para']} line={l['line_idx']} "
                  f"n={l['n_chars']} content={l['content_w']} "
                  f"natural={l['natural']} slack=+{l['slack']:.2f}pt")
    print(f"  Slack distribution of compressed lines:")
    slack_buckets = defaultdict(int)
    for l in findings["lines_with_compression"]:
        b = round(l["slack"])
        slack_buckets[b] += 1
    for s in sorted(slack_buckets.keys()):
        print(f"    slack={s:>+4}pt: {slack_buckets[s]:>3} lines")
    return findings


def analyze_m2_position():
    """D4: m2_position_sweep — position-1 exemption."""
    p = os.path.join(DATA_DIR, "m2_position_sweep.json")
    with open(p, encoding="utf-8") as f:
        data = json.load(f)
    rows = []
    for key, v in data.items():
        pos = v.get("yak_pos")
        n_chars = v.get("n_chars_line1")
        for adv in v.get("yak_advs", []):
            idx, ch, advval = adv
            ratio = round(advval / 12.0, 3)
            rows.append({"yak_pos": pos, "ch": ch, "adv": advval,
                          "ratio": ratio, "n_chars_line1": n_chars,
                          "compressed": ratio < 0.85})
    print(f"\n=== D4 m2_position_sweep ===")
    for r in rows:
        print(f"  yak_pos={r['yak_pos']:>3} ch={r['ch']} adv={r['adv']:.1f} "
              f"r={r['ratio']:.4f} n_line1={r['n_chars_line1']:>3} "
              f"compressed={r['compressed']}")
    return rows


def analyze_m2_wrap_budget():
    """D6: m2_wrap_budget — content_w sweep showing slack vs compression."""
    p = os.path.join(DATA_DIR, "m2_wrap_budget.json")
    with open(p, encoding="utf-8") as f:
        data = json.load(f)
    print(f"\n=== D6 m2_wrap_budget ===")
    print(f"  probes: {list(data.keys())}")
    rows = []
    for probe_key, probe in data.items():
        for cfg in probe.get("results", []):
            content_w = cfg.get("content_w")
            natural = cfg.get("natural")
            slack_natural = content_w - natural if (content_w and natural is not None) else None
            advs = cfg.get("advances", [])
            n_line1 = cfg.get("n_chars_line1")
            comp_total = cfg.get("yak_compression_total", 0)
            n_yak_in_line1 = 0
            n_yak_compressed = 0
            for i in range(min(n_line1 or 0, len(advs))):
                a = advs[i]
                if not isinstance(a, list) or len(a) < 2: continue
                ch, adv = a[0], a[1]
                if cls(ch) != "X":
                    n_yak_in_line1 += 1
                    if adv < 12.0 * 0.85:
                        n_yak_compressed += 1
            rows.append({"probe": probe_key, "content_w": content_w,
                          "natural": natural, "slack": slack_natural,
                          "n_chars_line1": n_line1,
                          "n_yak_line1": n_yak_in_line1,
                          "n_yak_compressed": n_yak_compressed,
                          "yak_compression_total_pt": comp_total})
    # Group by probe
    by_probe = defaultdict(list)
    for r in rows:
        by_probe[r["probe"]].append(r)
    for probe, prows in by_probe.items():
        print(f"\n  {probe}:")
        prows = sorted(prows, key=lambda r: -(r["content_w"] or 0))
        for r in prows[:25]:
            cw_s = f"{r['content_w']:>6.1f}" if r['content_w'] is not None else " None "
            nat_s = f"{r['natural']:>6.1f}" if r['natural'] is not None else " None "
            sl_s = f"{r['slack']:>+6.2f}" if r['slack'] is not None else " None "
            l1_s = f"{r['n_chars_line1']:>3}" if r['n_chars_line1'] is not None else " ?"
            print(f"    cw={cw_s} nat={nat_s} "
                  f"slack={sl_s} "
                  f"line1={l1_s} yak={r['n_yak_line1']:>2}"
                  f" comp={r['n_yak_compressed']:>2} total_pt={r['yak_compression_total_pt']}")
    return rows


def main():
    out = {}

    # D1, D2 — full per-char files
    analyze_doc("d77a_p10",
                 os.path.join(DATA_DIR, "d77a_p10_per_char_R33_diag.json"), out)
    analyze_doc("7f272a_p1",
                 os.path.join(DATA_DIR, "7f272a_p1_per_char_R33_diag.json"), out)

    # 3a4f_p74 (oxi-3 native)
    p3 = "pipeline_data/3a4f_p74_per_char_2026-05-02.json"
    if os.path.exists(p3):
        try:
            with open(p3, encoding="utf-8") as f:
                d3 = json.load(f)
            if "paragraphs" in d3:
                analyze_doc("3a4f_p74", p3, out)
            else:
                print(f"\n[skip] 3a4f_p74 has different shape: {list(d3.keys())[:5]}")
        except Exception as e:
            print(f"[skip] 3a4f_p74 read failed: {e}")

    # D4 m2_position
    out["m2_position_sweep"] = analyze_m2_position()

    # D6 m2_wrap_budget
    out["m2_wrap_budget"] = analyze_m2_wrap_budget()

    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2, default=str)
    print(f"\nWrote {OUT}")


if __name__ == "__main__":
    main()
