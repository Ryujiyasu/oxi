"""Scan ALL paragraphs across all SSIM-baseline docs with context tags
(body / table-cell-Nest / header / footer / textbox).

Day 12 SSIM verify revealed that Day 7's static scan was BODY-ONLY,
missing table-cell paragraphs that triggered the Day 8 fix in 7+ docs
and caused compensation-triangle SSIM regression. This scanner walks
document.xml + headerN.xml + footerN.xml and tags each paragraph with
its container chain.

For each paragraph, computes Day 8's narrow trigger condition:
  L > 0 AND FL > 0 AND leading_ws_pt > L_pt + FL_pt

Output: pipeline_data/all_paras_day8_trigger.json
        — list of {doc_id, file, container, pi, sz, lead_ws, L_tw,
          FL_tw, lead_ws_pt, indent_pt, text_preview}

Run: python tools/metrics/scan_all_paras_with_context.py
"""
from __future__ import annotations

import glob
import json
import os
import re
import sys
import zipfile

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
DOCX_DIR = os.path.join(REPO, "tools", "golden-test", "documents", "docx")
OUT = os.path.join(REPO, "pipeline_data", "all_paras_day8_trigger.json")


def parse_paragraphs_with_context(xml: str, source_label: str):
    """Walk paragraphs in an OOXML XML stream. Tag each by container chain.

    Container types:
      - body: top-level <w:p> in <w:body>
      - tbl-N: inside <w:tbl> at nest depth N
      - txbx: inside <w:txbxContent>
      - hdr / ftr: header.xml or footer.xml (caller passes label)

    Returns list of (paragraph_xml_body, container_label).
    """
    out = []
    # Stack of open container types
    stack = []  # e.g. ["body", "tbl-1", "tbl-2"]

    # Use a tag-by-tag scan
    pos = 0
    while pos < len(xml):
        m = re.search(r"<(/?)(w:tbl|w:txbxContent|w:p)\b([^>]*)>", xml[pos:])
        if not m:
            break
        is_close = (m.group(1) == "/")
        tag = m.group(2)
        attrs = m.group(3)
        is_self_closing = attrs.endswith("/")
        m_start = pos + m.start()
        m_end = pos + m.end()
        if tag == "w:tbl":
            if is_close:
                if stack and stack[-1].startswith("tbl-"):
                    stack.pop()
            else:
                depth = sum(1 for s in stack if s.startswith("tbl-")) + 1
                stack.append(f"tbl-{depth}")
        elif tag == "w:txbxContent":
            if is_close:
                if "txbx" in stack:
                    # Pop the most recent txbx marker
                    for i in range(len(stack) - 1, -1, -1):
                        if stack[i] == "txbx":
                            stack.pop(i)
                            break
            else:
                stack.append("txbx")
        elif tag == "w:p":
            if is_close or is_self_closing:
                pass
            else:
                # Find matching </w:p>
                end_m = re.search(r"</w:p>", xml[m_end:])
                if not end_m:
                    pos = m_end
                    continue
                p_body_start = m_start
                p_body_end = m_end + end_m.end()
                p_body = xml[p_body_start:p_body_end]
                # Determine container label
                if not stack:
                    container = source_label  # body / hdr / ftr
                elif stack[-1] == "txbx":
                    container = f"{source_label}-txbx"
                elif stack[-1].startswith("tbl-"):
                    container = f"{source_label}-{stack[-1]}"
                else:
                    container = source_label
                out.append((p_body, container))
                # Skip past the closing </w:p>
                pos = p_body_end
                continue
        pos = m_end

    return out


def compute_para_metrics(p_body: str):
    """Extract (lead_ws, lead_ws_pt, L_tw, FL_tw, sz_pt, text_preview, has_ind)."""
    ts = re.findall(r"<w:t[^>]*>([^<]*)</w:t>", p_body)
    full_text = "".join(ts)
    # First run sz
    first_run = re.search(r"<w:r\b[^>]*>(.*?)</w:r>", p_body, re.DOTALL)
    sz_pt = 10.5
    if first_run:
        sz_m = re.search(r'<w:sz\s+w:val="([^"]+)"', first_run.group(0))
        if sz_m:
            try:
                sz_pt = int(sz_m.group(1)) / 2.0
            except Exception:
                pass
    # Leading whitespace
    lead_ws = 0
    lead_ws_pt = 0.0
    for c in full_text:
        if c == " ":
            lead_ws += 1
            lead_ws_pt += sz_pt * 0.5
        elif c == "　":
            lead_ws += 1
            lead_ws_pt += sz_pt
        else:
            break
    # Indent
    ind_m = re.search(r"<w:ind\b([^/]*)/>", p_body)
    L_tw = 0
    FL_tw = 0
    has_ind = False
    if ind_m:
        has_ind = True
        ind_attrs = dict(re.findall(r'w:(\w+)="([^"]*)"', ind_m.group(1)))
        try:
            L_tw = int(ind_attrs.get("left", "0"))
        except Exception:
            pass
        try:
            FL_tw = int(ind_attrs.get("firstLine", "0"))
        except Exception:
            pass
    return {
        "text": full_text[:40],
        "sz_pt": sz_pt,
        "lead_ws": lead_ws,
        "lead_ws_pt": round(lead_ws_pt, 3),
        "L_tw": L_tw,
        "FL_tw": FL_tw,
        "indent_pt": round((L_tw + FL_tw) / 20.0, 3),
        "has_ind": has_ind,
    }


def scan_docx(docx_path: str):
    """Scan one docx, returning list of trigger paragraphs."""
    doc_id = os.path.basename(docx_path).split("_")[0]
    matches = []
    with zipfile.ZipFile(docx_path) as zf:
        # body
        try:
            body_xml = zf.read("word/document.xml").decode("utf-8")
        except KeyError:
            return matches
        for p_body, container in parse_paragraphs_with_context(body_xml, "body"):
            m = compute_para_metrics(p_body)
            if (
                m["L_tw"] > 0
                and m["FL_tw"] > 0
                and m["lead_ws_pt"] > m["indent_pt"]
                and m["indent_pt"] > 0
            ):
                matches.append({
                    "doc_id": doc_id,
                    "file": "document.xml",
                    "container": container,
                    **m,
                })
        # headers + footers + footnotes + endnotes + comments
        for name in zf.namelist():
            if not name.endswith(".xml"):
                continue
            base = os.path.basename(name)
            if base.startswith("header"):
                label = "hdr"
            elif base.startswith("footer"):
                label = "ftr"
            elif base == "footnotes.xml":
                label = "fnote"
            elif base == "endnotes.xml":
                label = "enote"
            elif base == "comments.xml":
                label = "cmnt"
            else:
                continue
            try:
                hf_xml = zf.read(name).decode("utf-8")
            except Exception:
                continue
            for p_body, container in parse_paragraphs_with_context(hf_xml, label):
                m = compute_para_metrics(p_body)
                if (
                    m["L_tw"] > 0
                    and m["FL_tw"] > 0
                    and m["lead_ws_pt"] > m["indent_pt"]
                    and m["indent_pt"] > 0
                ):
                    matches.append({
                        "doc_id": doc_id,
                        "file": name,
                        "container": container,
                        **m,
                    })
    return matches


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    docx_paths = sorted(glob.glob(os.path.join(DOCX_DIR, "*.docx")))
    print(f"Scanning {len(docx_paths)} docx files...")
    all_matches = []
    by_doc = {}
    for dp in docx_paths:
        try:
            ms = scan_docx(dp)
        except (zipfile.BadZipFile, Exception) as e:
            print(f"  skip {os.path.basename(dp)}: {e}")
            continue
        all_matches.extend(ms)
        if ms:
            by_doc.setdefault(ms[0]["doc_id"], []).extend(ms)

    print(f"\nTotal trigger paragraphs: {len(all_matches)}")
    print(f"\nBy doc:")
    print(f"{'doc':<20} {'count':<6} {'containers':<40} examples")
    print("-" * 130)
    for doc_id, ms in sorted(by_doc.items(), key=lambda kv: -len(kv[1])):
        containers = sorted(set(m["container"] for m in ms))
        ex = "; ".join(f'{m["container"]}|lws={m["lead_ws"]} L={m["L_tw"]} FL={m["FL_tw"]} text={m["text"][:18]!r}' for m in ms[:2])
        print(f"{doc_id:<20} {len(ms):<6} {','.join(containers):<40} {ex}")

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump({"matches": all_matches}, f, ensure_ascii=False, indent=2)
    print(f"\nWrote {OUT}")


if __name__ == "__main__":
    main()
