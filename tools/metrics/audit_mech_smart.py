"""Smarter Mech compression audit: target paragraphs likely to contain
Mech 1 triggers in jc=left context.

Mech 1 fires on:
  Type A → A   (e.g., 「『 ）（ 「（ etc.)
  Type B → A   (e.g., ）（ 」（ 、（)
  Type B → B   (e.g., 」） 、。 「」 ）」)

Strategy: in each candidate doc, find paragraphs with jc=left/none that
contain a Type-A or Type-B yakumono trigger pair. Measure those.
"""
import json, re, sys, time, zipfile
import subprocess
from pathlib import Path
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOCX_DIR = Path(r"C:\Users\ryuji\oxi-1\tools\golden-test\documents\docx")
RESULT_PATH = Path(r"C:\Users\ryuji\oxi-1\pipeline_data\mech_audit_smart.json")

TYPE_A = set("（「『【〔｛〈《［" "‘")
TYPE_B = set("）」』】〕｝〉》］、。，．" "—")
ALL_YAK = TYPE_A | TYPE_B

# Sampling target: top 5 real baseline docs
TARGET_DOCS = [
    "3a4f9fbe1a83_001620506.docx",
    "ed025cbecffb_index-23.docx",
    "d77a58485f16_20240705_resources_data_outline_08.docx",
    "b837808d0555_20240705_resources_data_guideline_02.docx",
    "e3c545fac7a7_LOD_Handbook.docx",
]


def has_mech1_trigger(text: str) -> bool:
    """Does text contain any Type-A→A, B→A, or B→B pair?"""
    for i in range(len(text) - 1):
        a, b = text[i], text[i+1]
        if a in TYPE_A and b in TYPE_A:
            return True
        if a in TYPE_B and b in TYPE_A:
            return True
        if a in TYPE_B and b in TYPE_B:
            return True
    return False


def find_trigger_paras(docx_path):
    """Return list of (para_xml_idx, jc, snippet) for paragraphs with
    jc=left/none AND Mech 1 trigger pair."""
    with zipfile.ZipFile(docx_path) as z:
        try:
            xml = z.read("word/document.xml").decode("utf-8")
        except KeyError:
            return []

    out = []
    for idx, m in enumerate(re.finditer(r'<w:p\b[^>]*>(.*?)</w:p>', xml, re.S), start=1):
        body = m.group(1)
        ppr_m = re.search(r'<w:pPr\b[^>]*>(.*?)</w:pPr>', body, re.S)
        ppr = ppr_m.group(1) if ppr_m else ""
        jc_m = re.search(r'<w:jc w:val="([^"]+)"', ppr)
        jc_val = jc_m.group(1) if jc_m else "(none)"
        if jc_val not in ("(none)", "left", "start"):
            continue
        txt = ''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', body))
        if not has_mech1_trigger(txt):
            continue
        out.append({"para_idx_xml": idx, "jc": jc_val, "snippet": txt[:60]})
    return out


def kill_word_and_restart():
    try:
        subprocess.run(['taskkill','/F','/IM','WINWORD.EXE'], capture_output=True)
    except Exception:
        pass
    time.sleep(3)
    word = w32.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    return word


def measure_para(word, doc, pi):
    """Measure per-char advance for paragraph #pi. Return list of
    (char, advance, font_size) and detected compression info."""
    try:
        para = doc.Paragraphs(pi)
    except Exception:
        return None
    chars = para.Range.Characters
    xs = []
    for ci in range(1, min(chars.Count + 1, 80)):
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
    if not xs:
        return None
    xs_sorted = sorted(xs, key=lambda v: (v[2], v[1]))
    advs = []
    for i in range(len(xs_sorted) - 1):
        a, b = xs_sorted[i], xs_sorted[i+1]
        if abs(a[2] - b[2]) > 0.5:
            continue
        advs.append((a[0], round(b[1] - a[1], 3), a[3]))
    comp_yak = []
    full_yak = []
    for ch, adv, sz in advs:
        if ch in ALL_YAK and sz > 0:
            if adv < sz * 0.7:
                comp_yak.append((ch, round(adv, 2), round(sz, 1)))
            else:
                full_yak.append((ch, round(adv, 2), round(sz, 1)))
    return {
        "n_chars": len(advs),
        "comp_yak": comp_yak,
        "full_yak": full_yak,
        "alignment": para.Alignment,
        "text": (para.Range.Text or "")[:80].replace("\r","\\r"),
    }


def measure_doc(p: Path, target_paras: list, max_paras=5):
    """Measure up to max_paras of target_paras from doc p, restarting
    Word on RPC failure."""
    word = kill_word_and_restart()
    out = []
    chosen = target_paras[:max_paras] if len(target_paras) >= max_paras else target_paras
    try:
        doc = None
        try:
            doc = word.Documents.Open(str(p.resolve()), ReadOnly=True)
        except Exception as e:
            print(f"  open ERR: {e}", file=sys.stderr)
            return out
        try:
            n_total = doc.Paragraphs.Count
            for tp in chosen:
                pi = tp["para_idx_xml"]
                if pi > n_total: continue
                try:
                    r = measure_para(word, doc, pi)
                    if r is None: continue
                    out.append({
                        "para_idx_xml": pi,
                        "jc_xml": tp["jc"],
                        "alignment_word": r["alignment"],
                        "n_chars": r["n_chars"],
                        "n_yak_compressed": len(r["comp_yak"]),
                        "n_yak_total": len(r["comp_yak"]) + len(r["full_yak"]),
                        "compressed_examples": r["comp_yak"][:5],
                        "compression_detected": len(r["comp_yak"]) > 0,
                        "snippet": tp["snippet"][:50],
                    })
                except Exception as e:
                    if "RPC" in str(e) or "サーバー" in str(e):
                        # restart
                        try: doc.Close(SaveChanges=0)
                        except: pass
                        try: word.Quit()
                        except: pass
                        word = kill_word_and_restart()
                        try: doc = word.Documents.Open(str(p.resolve()), ReadOnly=True)
                        except: doc = None
                        if not doc: break
                    else:
                        out.append({"para_idx_xml": pi, "error": str(e)})
        finally:
            if doc:
                try: doc.Close(SaveChanges=0)
                except: pass
    finally:
        try: word.Quit()
        except: pass
    return out


def main():
    print(f"Targeting {len(TARGET_DOCS)} baseline docs", file=sys.stderr)
    results = {}
    for name in TARGET_DOCS:
        p = DOCX_DIR / name
        if not p.exists():
            results[name] = {"error": "file missing"}
            continue
        print(f"\n=== {name} ===", file=sys.stderr)
        targets = find_trigger_paras(p)
        print(f"  {len(targets)} paras with Mech 1 trigger pair + jc=left/none", file=sys.stderr)
        if not targets:
            results[name] = {"trigger_paras": 0, "measurements": []}
            continue
        # Take 5 targets, prefer ones with most yakumono in snippet
        targets.sort(key=lambda x: -sum(1 for c in x["snippet"] if c in ALL_YAK))
        top5 = targets[:5]
        meas = measure_doc(p, top5, max_paras=5)
        results[name] = {
            "trigger_paras_total": len(targets),
            "sampled": [t["para_idx_xml"] for t in top5],
            "measurements": meas,
        }

    RESULT_PATH.parent.mkdir(parents=True, exist_ok=True)
    with open(RESULT_PATH, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)

    print("\n=== Audit Summary ===\n")
    print(f"{'Doc':50s} {'trig_paras':>10} {'measured':>9} {'comp':>5} {'comp_yak':>10}")
    print("-" * 90)
    for n, r in results.items():
        if "error" in r:
            print(f"  {n[:48]:48s}  ERROR: {r['error']}")
            continue
        meas = r.get("measurements", [])
        n_comp = sum(1 for m in meas if m.get("compression_detected"))
        n_meas = sum(1 for m in meas if "compression_detected" in m)
        n_comp_yak = sum(m.get("n_yak_compressed", 0) for m in meas)
        n_yak_tot = sum(m.get("n_yak_total", 0) for m in meas)
        print(f"  {n[:48]:48s} {r.get('trigger_paras_total', 0):>10} {n_meas:>9} {n_comp:>5} "
              f"{n_comp_yak:>5}/{n_yak_tot:<5}")
    print()
    for n, r in results.items():
        if "measurements" not in r: continue
        for m in r["measurements"]:
            if not isinstance(m, dict) or "error" in m: continue
            tag = "✓ COMPRESSED" if m["compression_detected"] else "  no"
            print(f"  {n[:30]:30s} p{m['para_idx_xml']:3d} jc={m['jc_xml']:>6}/Word={m['alignment_word']} {tag}  yak={m['n_yak_compressed']}/{m['n_yak_total']}  ex={m['compressed_examples']}")
            print(f"       {m['snippet']!r}")


if __name__ == "__main__":
    main()
