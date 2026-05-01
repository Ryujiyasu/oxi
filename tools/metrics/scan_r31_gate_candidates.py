"""Scan baseline for docs matching R31 narrow gate conditions:
  (chars-indent paragraph) AND (cross-run yakumono pair).

Cross-tab against effective kern presence.

If most/all R31-gate candidates have kern → R31 narrow path is redundant
with R32 kern gate.
"""
import json
import re
import sys
import zipfile
from pathlib import Path
from collections import Counter

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOCX_DIR = Path("tools/golden-test/documents/docx")
KERN_AUDIT = Path("pipeline_data/kern_audit_2026-05-02.json")
OUT = Path("pipeline_data/r31_gate_candidates.json")

# Yakumono character classes (per spec §4.7 Type A/B/C)
TYPE_A_OPEN = set("（「『【〔｛〈《［" + "“‘")
TYPE_B_CLOSE = set("）」』】〕｝〉》］、。，．" + "”’" + "—")
# CJK trigger: any of A/B/C plus CJK ideograph block
import unicodedata
def is_cjk_ideo(c):
    return '一' <= c <= '鿿' or '぀' <= c <= 'ゟ' or '゠' <= c <= 'ヿ'


def is_yakumono_trigger(c):
    """Trigger for Type A/B compression — A or B yakumono OR CJK ideograph."""
    return c in TYPE_A_OPEN or c in TYPE_B_CLOSE or is_cjk_ideo(c)


def is_yakumono_b(c):
    return c in TYPE_B_CLOSE


def is_yakumono_a(c):
    return c in TYPE_A_OPEN


def has_chars_indent(doc_xml: str) -> bool:
    """Check if any paragraph has *Chars-suffixed indent (leftChars/firstLineChars/etc)."""
    return bool(re.search(r'<w:ind[^>]*(leftChars|firstLineChars|hangingChars|rightChars|startChars|endChars)=', doc_xml))


def has_cross_run_yakumono_pair(doc_xml: str) -> bool:
    """Detect end-of-<w:r> char + start-of-<w:r> char being a yakumono pair.

    Specifically: prev_run last char is B-close + next_run first char is trigger,
    OR prev_run last char is trigger + next_run first char is A-open.
    """
    # Extract per-paragraph run texts. Crude regex (rolls over <w:t> contents):
    # Find <w:p>...</w:p> chunks
    for p_match in re.finditer(r'<w:p\b[^>]*>(.*?)</w:p>', doc_xml, re.DOTALL):
        para_xml = p_match.group(1)
        # Extract text from each <w:r>...</w:r>'s <w:t>...</w:t>
        run_texts = []
        for r_match in re.finditer(r'<w:r\b[^>]*>(.*?)</w:r>', para_xml, re.DOTALL):
            run_xml = r_match.group(1)
            # Concatenate all <w:t>...</w:t> contents in this run
            run_text = ""
            for t_match in re.finditer(r'<w:t\b[^>]*>(.*?)</w:t>', run_xml, re.DOTALL):
                run_text += t_match.group(1)
            if run_text:
                run_texts.append(run_text)
        # Check adjacent pairs
        for i in range(len(run_texts) - 1):
            prev = run_texts[i]
            nxt = run_texts[i+1]
            if not prev or not nxt: continue
            last = prev[-1]
            first = nxt[0]
            # B-close at end + trigger at start of next
            if is_yakumono_b(last) and is_yakumono_trigger(first):
                return True
            # trigger at end + A-open at start of next
            if is_yakumono_trigger(last) and is_yakumono_a(first):
                return True
    return False


def has_compress_punct(settings_xml: str) -> bool:
    """Check w:characterSpacingControl=compressPunctuation* in settings.xml."""
    return bool(re.search(r'<w:characterSpacingControl[^>]*w:val="compressPunctuation', settings_xml))


def main():
    kern = json.loads(KERN_AUDIT.read_text(encoding="utf-8"))
    kern_by_doc = {d["doc_id_short"]: d for d in kern["audit"]}

    results = []
    docx_files = sorted(DOCX_DIR.glob("*.docx"))
    for f in docx_files:
        doc_id = f.stem.split("_")[0] + ("_" + f.stem.split("_", 1)[1] if "_" in f.stem else "")
        doc_id_short = f.stem[:12]  # e.g., "3a4f9fbe1a83"
        try:
            with zipfile.ZipFile(f) as z:
                doc_xml = z.read("word/document.xml").decode("utf-8", errors="replace")
                try:
                    settings_xml = z.read("word/settings.xml").decode("utf-8", errors="replace")
                except KeyError:
                    settings_xml = ""
        except Exception as e:
            results.append({"doc": f.name, "error": str(e)})
            continue
        ci = has_chars_indent(doc_xml)
        cr = has_cross_run_yakumono_pair(doc_xml)
        cp = has_compress_punct(settings_xml)
        kern_info = kern_by_doc.get(doc_id_short, {})
        k_eff = kern_info.get("effective_kern")
        # Determine source from dd vs normal
        if kern_info.get("normal_kern"):
            k_src = "Normal style"
        elif kern_info.get("dd_kern"):
            k_src = "docDefaults"
        else:
            k_src = None
        rec = {
            "doc": f.name,
            "doc_id": doc_id_short,
            "has_chars_indent": ci,
            "has_cross_run_pair": cr,
            "has_compress_punct": cp,
            "r31_gate_fires": ci and cr and cp,
            "effective_kern": k_eff,
            "kern_source": k_src,
        }
        results.append(rec)

    # Cross-tab
    print(f"\nScanned {len(results)} docs.")
    fires = [r for r in results if r.get("r31_gate_fires")]
    print(f"R31 narrow-gate fires (chars-indent + cross-run + cP): {len(fires)}")

    print(f"\n=== R31-fire docs: kern presence cross-tab ===")
    with_kern = [r for r in fires if r["effective_kern"]]
    no_kern = [r for r in fires if not r["effective_kern"]]
    print(f"  WITH effective kern: {len(with_kern)} / {len(fires)}")
    for r in with_kern:
        print(f"    {r['doc_id']:12} kern={r['effective_kern']} ({r['kern_source']})  {r['doc']}")
    print(f"  WITHOUT effective kern: {len(no_kern)} / {len(fires)}")
    for r in no_kern:
        print(f"    {r['doc_id']:12} (no kern)  {r['doc']}")

    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps(results, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"\nSaved -> {OUT}")


if __name__ == "__main__":
    main()
