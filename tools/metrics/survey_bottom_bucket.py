"""依頼 12: R32 後 bottom-bucket survey.

Extract bottom 30 worst-SSIM pages from ssim_baseline.json (proxy for post-R32),
classify each by structural features, cluster, propose next hypotheses.

Features per doc:
- effective kern (from kern_audit_2026-05-02.json)
- jc distribution (jc=both / jc=left / jc=center / jc=right counts)
- list_marker ratio (paragraphs with numPr / total)
- chars-indent ratio (paragraphs with *Chars indent / total)
- table count (number of <w:tbl>)
- floating shape count (<w:drawing> with <wp:anchor>)
- footnote count
- paragraph count (rough)

Output: pipeline_data/bottom_bucket_survey.json + console summary
"""
import json
import re
import sys
import zipfile
from pathlib import Path
from collections import Counter, defaultdict

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOCX_DIR = Path("tools/golden-test/documents/docx")
SSIM = Path("pipeline_data/ssim_baseline.json")
KERN = Path("pipeline_data/kern_audit_2026-05-02.json")
OUT = Path("pipeline_data/bottom_bucket_survey.json")


def short_id(name: str) -> str:
    """Extract 12-char doc_id_short from filename or baseline key."""
    # baseline keys are "<short>_<rest>", filenames are "<short>_<rest>.docx"
    base = name.replace(".docx", "")
    return base.split("_")[0]


def collect_bottom_pages(top_n: int = 30):
    """Return list of (doc_id_short, doc_full_name, page, ssim) sorted asc."""
    data = json.loads(SSIM.read_text(encoding="utf-8"))
    rows = []
    for full_name, pages in data.items():
        sid = short_id(full_name)
        for page, score in pages.items():
            rows.append((sid, full_name, int(page), float(score)))
    rows.sort(key=lambda r: r[3])
    return rows[:top_n]


def analyze_docx(docx_path: Path) -> dict:
    """Return structural feature dict."""
    feat = {
        "n_paras": 0, "n_paras_jc_both": 0, "n_paras_jc_left": 0,
        "n_paras_jc_center": 0, "n_paras_jc_right": 0, "n_paras_jc_distribute": 0,
        "n_paras_numPr": 0, "n_paras_chars_indent": 0,
        "n_tbl": 0, "n_floating_shape": 0, "n_footnotes": 0,
        "compress_punct": False, "doc_grid": None, "compat_mode": None,
    }
    try:
        with zipfile.ZipFile(docx_path) as z:
            doc_xml = z.read("word/document.xml").decode("utf-8", errors="replace")
            try:
                settings_xml = z.read("word/settings.xml").decode("utf-8", errors="replace")
            except KeyError:
                settings_xml = ""
            try:
                fn_xml = z.read("word/footnotes.xml").decode("utf-8", errors="replace")
            except KeyError:
                fn_xml = ""
    except Exception as e:
        return {"error": str(e)}

    # Paragraph-level scan
    feat["n_paras"] = len(re.findall(r"<w:p\b", doc_xml))

    # jc inside pPr — must be on a paragraph (not run)
    jc_matches = re.findall(r'<w:pPr>(?:[^<]|<(?!/w:pPr>))*?<w:jc\s+w:val="([^"]+)"', doc_xml)
    for jc in jc_matches:
        key = f"n_paras_jc_{jc}"
        if key in feat:
            feat[key] = feat.get(key, 0) + 1
        else:
            feat[key] = feat.get(key, 0) + 1

    # numPr (list paragraph)
    feat["n_paras_numPr"] = len(re.findall(r"<w:numPr>", doc_xml))

    # *Chars indent (any of leftChars/firstLineChars/hangingChars/rightChars/startChars/endChars)
    feat["n_paras_chars_indent"] = len(re.findall(r'<w:ind[^>]*(leftChars|firstLineChars|hangingChars|rightChars|startChars|endChars)=', doc_xml))

    # Tables
    feat["n_tbl"] = len(re.findall(r"<w:tbl\b", doc_xml))

    # Floating shapes — <wp:anchor> (anchor = floating, inline = inline)
    feat["n_floating_shape"] = len(re.findall(r"<wp:anchor\b", doc_xml))

    # Footnotes — count actual user footnotes (not separator/continuation)
    if fn_xml:
        # Skip type="separator" and type="continuationSeparator"
        all_fn = re.findall(r'<w:footnote\s+(?:w:type="([^"]*)"\s+)?w:id', fn_xml)
        feat["n_footnotes"] = sum(1 for t in all_fn if t not in ("separator", "continuationSeparator"))

    # Compress punct setting
    feat["compress_punct"] = bool(re.search(r'<w:characterSpacingControl[^>]*w:val="compressPunctuation', settings_xml))

    # docGrid
    grid_match = re.search(r'<w:docGrid\s+w:type="([^"]+)"(?:\s+w:linePitch="([^"]+)")?(?:\s+w:charSpace="([^"]+)")?', doc_xml)
    if grid_match:
        feat["doc_grid"] = {"type": grid_match.group(1), "linePitch": grid_match.group(2), "charSpace": grid_match.group(3)}

    # compat mode
    compat = re.search(r'compatibilityMode"\s+w:val="(\d+)"', settings_xml)
    if compat:
        feat["compat_mode"] = int(compat.group(1))
    return feat


def main():
    bottom = collect_bottom_pages(30)
    kern_data = json.loads(KERN.read_text(encoding="utf-8"))
    kern_by_short = {d["doc_id_short"]: d for d in kern_data["audit"]}

    # Aggregate by doc (multi-page bottom may have same doc)
    by_doc = defaultdict(list)
    for sid, full, page, ssim in bottom:
        by_doc[(sid, full)].append((page, ssim))

    # Find the docx files
    rows = []
    for (sid, full), pages in by_doc.items():
        # Find docx file by short_id prefix
        candidates = list(DOCX_DIR.glob(f"{sid}*.docx"))
        if not candidates:
            print(f"WARN: no docx found for {sid}")
            continue
        docx = candidates[0]
        feat = analyze_docx(docx)
        kern_info = kern_by_short.get(sid, {})
        feat["effective_kern"] = kern_info.get("effective_kern")
        feat["kern_source"] = ("Normal style" if kern_info.get("normal_kern") else
                                "docDefaults" if kern_info.get("dd_kern") else None)
        feat["doc_id_short"] = sid
        feat["doc_full"] = full
        feat["bottom_pages"] = sorted(pages, key=lambda p: p[1])
        feat["min_ssim"] = min(p[1] for p in pages)
        rows.append(feat)

    rows.sort(key=lambda r: r["min_ssim"])

    # Print summary
    print(f"\n=== Bottom {len(bottom)} worst pages, {len(rows)} unique docs ===\n")
    print(f"{'doc_id':12} {'kern':>5} {'jcB':>4} {'jcL':>4} {'jcC':>4} "
          f"{'np':>4} {'cI':>4} {'tb':>3} {'fs':>3} {'fn':>3} {'np tot':>6} "
          f"{'gridT':<10} {'cMode':>5} min_ssim   pages")
    for r in rows:
        kern = "yes" if r["effective_kern"] else "-"
        ksrc = "(N)" if r["kern_source"] == "Normal style" else "(D)" if r["kern_source"] == "docDefaults" else ""
        kern_disp = f"{kern}{ksrc}" if kern == "yes" else "-"
        gtype = (r["doc_grid"] or {}).get("type") or "-"
        cmode = r.get("compat_mode") or "-"
        npp = r.get("n_paras", 0)
        pages_str = " ".join(f"p{p}={s:.3f}" for p, s in r["bottom_pages"][:3])
        print(f"{r['doc_id_short']:12} {kern_disp:>5} "
              f"{r.get('n_paras_jc_both',0):>4} {r.get('n_paras_jc_left',0):>4} "
              f"{r.get('n_paras_jc_center',0):>4} "
              f"{r['n_paras_numPr']:>4} {r['n_paras_chars_indent']:>4} "
              f"{r['n_tbl']:>3} {r['n_floating_shape']:>3} {r['n_footnotes']:>3} "
              f"{npp:>6} "
              f"{gtype:<10} {cmode:>5} {r['min_ssim']:.3f}  {pages_str}")

    # Cluster analysis
    print("\n=== Cluster summary ===")
    n_kern = sum(1 for r in rows if r["effective_kern"])
    n_no_kern = len(rows) - n_kern
    print(f"  kern: {n_kern} with / {n_no_kern} without")

    n_floating = sum(1 for r in rows if r["n_floating_shape"] > 0)
    n_table = sum(1 for r in rows if r["n_tbl"] > 0)
    n_fn = sum(1 for r in rows if r["n_footnotes"] > 0)
    n_chars_ind = sum(1 for r in rows if r["n_paras_chars_indent"] > 0)
    n_numPr = sum(1 for r in rows if r["n_paras_numPr"] > 0)
    print(f"  has floating shape: {n_floating}")
    print(f"  has table: {n_table}")
    print(f"  has footnote: {n_fn}")
    print(f"  has chars-indent: {n_chars_ind}")
    print(f"  has numPr (list): {n_numPr}")

    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps(rows, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"\nSaved -> {OUT}")


if __name__ == "__main__":
    main()
