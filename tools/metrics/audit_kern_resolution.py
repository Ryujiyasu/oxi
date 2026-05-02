"""
Comprehensive kern audit for R32 design.

For each baseline docx, determine the EFFECTIVE kern source per Word's
resolution rule:

  ECMA-376 + Word effective-kern resolution (per OOXML §17.7.5.4 and
  empirical observation):

    Per-run kern resolution chain (highest to lowest priority):
      1. run.rPr.kern (direct on run)
      2. pPr.rPr.kern (paragraph default run rPr)
      3. paragraph style's rPr.kern (via pStyle), then basedOn chain
      4. document default paragraph style (w:default="1") rPr.kern
      5. docDefaults rPrDefault.kern

    Special case: kern w:val="0" explicitly DISABLES kerning (overrides
    inherited values).

  Effective_kern = first non-None value in the chain.
  If all None → no kern.
  If kern w:val=0 found → disabled.
  Else → enabled with that val.

For the audit, we determine for each doc:
  - Has kern in docDefaults rPrDefault? (where R32 currently checks)
  - Has kern in Normal/default style rPr? (the missing 24)
  - Has kern in any other named style?
  - Has kern in any body w:r run?
  - The "effective kern" status that R32 SHOULD see.

Output: pipeline_data/kern_audit.json + a list of "missing" docs (currently
classified as no-kern by R32 but actually have kern via Normal style).
"""
import os
import zipfile
import json
import re
import xml.etree.ElementTree as ET

DOCX_DIR = "C:/Users/ryuji/oxi-main/pipeline_data/golden_per_page"
OUT_JSON = os.path.join(os.path.dirname(__file__), "..", "..", "pipeline_data", "kern_audit.json")

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}


def parse_styles_xml(zip_obj):
    """Return dict with kern presence info from styles.xml."""
    info = {
        "docDefaults_kern": None,           # kern val in rPrDefault, or None
        "default_para_style_id": None,      # style with w:default="1" type="paragraph"
        "default_para_kern": None,          # kern in that default style's rPr
        "normal_style_id": None,            # explicit "Normal" or "a" or styleId of default
        "normal_style_kern": None,          # kern in Normal style
        "named_styles_with_kern": [],       # [(styleId, kern_val), ...]
        "all_style_kern_count": 0,
    }
    try:
        with zip_obj.open("word/styles.xml") as f:
            tree = ET.parse(f)
            root = tree.getroot()

        # docDefaults
        rpr_default = root.find(".//w:docDefaults/w:rPrDefault/w:rPr/w:kern", NS)
        if rpr_default is not None:
            v = rpr_default.get(f"{{{W}}}val")
            info["docDefaults_kern"] = v

        # Find default paragraph style (w:default="1" w:type="paragraph")
        for style in root.findall(".//w:style", NS):
            stype = style.get(f"{{{W}}}type")
            sdefault = style.get(f"{{{W}}}default")
            sid = style.get(f"{{{W}}}styleId")
            if stype == "paragraph" and sdefault == "1":
                info["default_para_style_id"] = sid
                kern_el = style.find("./w:rPr/w:kern", NS)
                if kern_el is not None:
                    info["default_para_kern"] = kern_el.get(f"{{{W}}}val")

        # Find named "Normal" or "a" style explicitly
        for style in root.findall(".//w:style", NS):
            sid = style.get(f"{{{W}}}styleId")
            if sid in ("Normal", "a"):
                info["normal_style_id"] = sid
                kern_el = style.find("./w:rPr/w:kern", NS)
                if kern_el is not None:
                    info["normal_style_kern"] = kern_el.get(f"{{{W}}}val")

        # All named styles with kern
        for style in root.findall(".//w:style", NS):
            sid = style.get(f"{{{W}}}styleId") or "?"
            stype = style.get(f"{{{W}}}type") or "?"
            kern_el = style.find("./w:rPr/w:kern", NS)
            if kern_el is not None:
                v = kern_el.get(f"{{{W}}}val")
                info["named_styles_with_kern"].append({
                    "styleId": sid, "type": stype, "kern": v,
                })
                info["all_style_kern_count"] += 1
    except Exception as e:
        info["error"] = str(e)
    return info


def parse_document_kern(zip_obj):
    """Check document.xml for body-level kern (in r/rPr or in pPr/rPr) and
    collect the set of style IDs referenced by body paragraphs."""
    info = {
        "any_run_kern": False,
        "any_pPr_rPr_kern": False,
        "first_run_kern_val": None,
        "body_used_styles": set(),
    }
    try:
        with zip_obj.open("word/document.xml") as f:
            data = f.read().decode("utf-8")
        for m in re.finditer(r'<w:kern\s+w:val="([^"]+)"', data):
            info["any_run_kern"] = True
            if info["first_run_kern_val"] is None:
                info["first_run_kern_val"] = m.group(1)
        for m in re.finditer(r'<w:pPr>.*?<w:rPr>.*?<w:kern\s+w:val="([^"]+)"', data, re.DOTALL):
            info["any_pPr_rPr_kern"] = True
            break
        # Collect pStyle references from body paragraphs
        for m in re.finditer(r'<w:pStyle\s+w:val="([^"]+)"', data):
            info["body_used_styles"].add(m.group(1))
    except Exception as e:
        info["error"] = str(e)
    return info


def is_effective_kern_present(styles_info, doc_info):
    """Per the resolution rule, what does R32 SHOULD see for body paragraphs?

    For body paragraphs: priority is docDefaults > default-paragraph-style >
    Normal style (id="Normal" or "a"). We don't include named non-default
    paragraph styles (Heading*, Title etc.) because body paragraphs by default
    use the default paragraph style, not those.
    """
    # docDefaults
    if styles_info.get("docDefaults_kern") and styles_info["docDefaults_kern"] != "0":
        return ("docDefaults", styles_info["docDefaults_kern"])
    # Default paragraph style (w:default="1") — this IS body's effective parent
    if styles_info.get("default_para_kern") and styles_info["default_para_kern"] != "0":
        return ("default_para_style", styles_info["default_para_kern"])
    # Explicit Normal style (id Normal or "a") — same as default in most cases
    if styles_info.get("normal_style_kern") and styles_info["normal_style_kern"] != "0":
        return ("Normal_style", styles_info["normal_style_kern"])
    # Body run kern (per-run direct override on body)
    if doc_info.get("any_run_kern") and doc_info.get("first_run_kern_val") not in (None, "0"):
        return ("body_run_direct", doc_info["first_run_kern_val"])
    return (None, None)


def named_style_with_kern_used_in_body(styles_info, doc_info, body_used_styles):
    """Detect body paragraphs that actually use a NAMED style (e.g. heading)
    that has kern in its rPr — even though that style isn't the default."""
    for s in styles_info.get("named_styles_with_kern", []):
        if s.get("kern") in (None, "0"):
            continue
        sid = s.get("styleId")
        if sid in body_used_styles:
            return (sid, s.get("kern"))
    return (None, None)


def r32_currently_detects(styles_info):
    """What R32's current docDefaults-only check sees."""
    v = styles_info.get("docDefaults_kern")
    return v is not None and v != "0"


def main():
    if not os.path.isdir(DOCX_DIR):
        print(f"DOCX_DIR not found: {DOCX_DIR}")
        return

    # The 177-doc baseline lives as per-page splits in golden_per_page/.
    # Each source doc has *_p1.docx, *_p2.docx etc. — styles.xml is identical
    # across pages, so we use _p1 as representative.
    rows = []
    seen_sources = set()
    for fname in sorted(os.listdir(DOCX_DIR)):
        if not fname.endswith(".docx"):
            continue
        # Pick only _p1 representatives
        m = re.match(r"^(.+)_p(\d+)\.docx$", fname)
        if not m:
            continue
        source = m.group(1)
        if int(m.group(2)) != 1:
            continue
        if source in seen_sources:
            continue
        seen_sources.add(source)
        path = os.path.join(DOCX_DIR, fname)
        try:
            with zipfile.ZipFile(path) as z:
                styles_info = parse_styles_xml(z)
                doc_info = parse_document_kern(z)
        except Exception as e:
            rows.append({"path": fname, "error": str(e)})
            continue

        eff_source, eff_val = is_effective_kern_present(styles_info, doc_info)
        # Also check if any heading/named style with kern is actually used in body
        named_used_sid, named_used_val = named_style_with_kern_used_in_body(
            styles_info, doc_info, doc_info.get("body_used_styles", set())
        )
        # If still no effective kern but a named non-default style with kern is used,
        # that paragraph IS getting effective kern via that style.
        if eff_source is None and named_used_sid:
            eff_source = f"named_style:{named_used_sid}"
            eff_val = named_used_val

        r32_sees = r32_currently_detects(styles_info)
        eff_present = eff_source is not None
        misclassified = eff_present and not r32_sees

        rows.append({
            "source_doc": source,
            "path": fname,
            "docDefaults_kern": styles_info.get("docDefaults_kern"),
            "default_para_style_id": styles_info.get("default_para_style_id"),
            "default_para_kern": styles_info.get("default_para_kern"),
            "normal_style_id": styles_info.get("normal_style_id"),
            "normal_style_kern": styles_info.get("normal_style_kern"),
            "named_styles_with_kern_count": styles_info.get("all_style_kern_count"),
            "any_run_kern": doc_info.get("any_run_kern"),
            "first_run_kern_val": doc_info.get("first_run_kern_val"),
            "effective_source": eff_source,
            "effective_val": eff_val,
            "r32_currently_detects": r32_sees,
            "misclassified_by_r32": misclassified,
        })

    # Summary
    n_total = len(rows)
    n_with_eff_kern = sum(1 for r in rows if r.get("effective_source"))
    n_r32_detects = sum(1 for r in rows if r.get("r32_currently_detects"))
    n_misclassified = sum(1 for r in rows if r.get("misclassified_by_r32"))
    by_source = {}
    for r in rows:
        s = r.get("effective_source") or "no_kern"
        by_source[s] = by_source.get(s, 0) + 1

    print(f"Total docs: {n_total}")
    print(f"  With effective kern: {n_with_eff_kern}")
    print(f"  R32 currently detects (docDefaults-only): {n_r32_detects}")
    print(f"  Misclassified by R32 (effective kern via non-docDefaults source): {n_misclassified}")
    print(f"\nBreakdown by effective-kern source:")
    for s, n in sorted(by_source.items(), key=lambda x: -x[1]):
        print(f"  {s:25s}: {n}")

    print(f"\nMisclassified docs (= the missing 24+):")
    miscls = [r for r in rows if r.get("misclassified_by_r32")]
    for r in miscls:
        print(f"  {r['path']:50s} source={r['effective_source']:20s} "
              f"val={r['effective_val']} "
              f"normalStyleId={r.get('normal_style_id')}")

    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump({"summary": {
            "total": n_total,
            "with_effective_kern": n_with_eff_kern,
            "r32_currently_detects": n_r32_detects,
            "misclassified": n_misclassified,
            "by_source": by_source,
        }, "rows": rows}, f, indent=2, ensure_ascii=False)
    print(f"\nSaved to {OUT_JSON}")


if __name__ == "__main__":
    main()
