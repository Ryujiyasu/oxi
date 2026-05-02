"""
Scan pipeline_data/docx for tall-header candidates.

A "tall header" candidate has either:
  (a) multiple paragraphs in any of header1/2/3.xml, OR
  (b) header content with explicit large font size (>=14pt), OR
  (c) topMargin <= headerDistance + 3 * 14pt (potential overflow zone)

Output:
  pipeline_data/tall_header_candidates.json
"""
import os
import zipfile
import json
import re
import xml.etree.ElementTree as ET

DOCX_DIR = os.path.join(os.path.dirname(__file__), "..", "..", "pipeline_data", "docx")
OUT_JSON = os.path.join(os.path.dirname(__file__), "..", "..", "pipeline_data", "tall_header_candidates.json")

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def scan_docx(path):
    info = {
        "path": os.path.basename(path),
        "headers": [],
        "section_props": [],
    }
    try:
        with zipfile.ZipFile(path) as z:
            names = z.namelist()
            # Find header XMLs
            hdr_files = [n for n in names if re.match(r"word/header\d*\.xml$", n)]
            for hf in hdr_files:
                with z.open(hf) as f:
                    tree = ET.parse(f)
                    root = tree.getroot()
                    paras = root.findall(".//w:p", NS)
                    para_info = []
                    for p in paras:
                        runs = p.findall(".//w:r", NS)
                        text = "".join(t.text or "" for t in p.findall(".//w:t", NS))
                        sizes = []
                        for rr in runs:
                            for sz in rr.findall(".//w:sz", NS):
                                v = sz.get(f"{{{NS['w']}}}val")
                                if v:
                                    sizes.append(int(v) / 2.0)
                        para_info.append({
                            "text": text[:50],
                            "sizes_pt": sizes,
                        })
                    info["headers"].append({
                        "file": hf,
                        "para_count": len(paras),
                        "paragraphs": para_info,
                    })
            # Find section properties
            with z.open("word/document.xml") as f:
                tree = ET.parse(f)
                root = tree.getroot()
                for sectPr in root.findall(".//w:sectPr", NS):
                    pgSz = sectPr.find("w:pgSz", NS)
                    pgMar = sectPr.find("w:pgMar", NS)
                    sec = {}
                    if pgSz is not None:
                        sec["pgSz_w"] = pgSz.get(f"{{{NS['w']}}}w")
                        sec["pgSz_h"] = pgSz.get(f"{{{NS['w']}}}h")
                    if pgMar is not None:
                        for k in ["top", "bottom", "left", "right", "header", "footer"]:
                            v = pgMar.get(f"{{{NS['w']}}}{k}")
                            if v:
                                sec[f"pgMar_{k}_tw"] = int(v)
                                sec[f"pgMar_{k}_pt"] = int(v) / 20.0
                    info["section_props"].append(sec)
    except Exception as e:
        info["error"] = str(e)
    return info


def is_tall_header_candidate(info):
    """Heuristic: header has 2+ paragraphs, OR header has large font, OR
    headerDistance + 3*14pt > topMargin."""
    has_multipara_header = any(h["para_count"] >= 2 for h in info["headers"])
    has_large_font = any(
        any(sz >= 14 for p in h["paragraphs"] for sz in p["sizes_pt"])
        for h in info["headers"]
    )
    overflow_zone = False
    for sec in info["section_props"]:
        tm = sec.get("pgMar_top_pt", 72)
        hd = sec.get("pgMar_header_pt", 36)
        if tm - hd < 42:  # < 3 lines of 14pt header
            overflow_zone = True
            break
    return has_multipara_header or has_large_font or overflow_zone


def main():
    if not os.path.isdir(DOCX_DIR):
        print(f"DOCX_DIR not found: {DOCX_DIR}")
        return
    candidates = []
    for fname in sorted(os.listdir(DOCX_DIR)):
        if not fname.endswith(".docx"):
            continue
        path = os.path.join(DOCX_DIR, fname)
        info = scan_docx(path)
        if is_tall_header_candidate(info):
            info["candidate_score"] = (
                sum(h["para_count"] for h in info["headers"]) * 10
                + sum(1 for h in info["headers"]
                      for p in h["paragraphs"]
                      for sz in p["sizes_pt"] if sz >= 14)
            )
            candidates.append(info)

    candidates.sort(key=lambda x: -x.get("candidate_score", 0))

    print(f"Found {len(candidates)} tall-header candidates (top 15):")
    for c in candidates[:15]:
        n_p = sum(h["para_count"] for h in c["headers"])
        sizes = sorted(set(
            sz for h in c["headers"]
            for p in h["paragraphs"]
            for sz in p["sizes_pt"]
        ))
        sec0 = c["section_props"][0] if c["section_props"] else {}
        print(f"  score={c.get('candidate_score', 0):3} {c['path']:40s} "
              f"hdr_paras={n_p:3} sizes={sizes} "
              f"tm={sec0.get('pgMar_top_pt', '?'):>5} hd={sec0.get('pgMar_header_pt', '?'):>5}")

    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump(candidates, f, indent=2, ensure_ascii=False)
    print(f"\nSaved {len(candidates)} candidates to {OUT_JSON}")


if __name__ == "__main__":
    main()
