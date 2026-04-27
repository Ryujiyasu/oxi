"""Measure Word's ruby (furigana) geometry via COM + XML extraction.

For each fixture in pipeline_data/docx/RUBY_V*.docx:
  1. Parse word/document.xml to extract every w:ruby element's:
     - base text, ruby text
     - rubyAlign, hps, hpsRaise, hpsBaseText, lid
  2. Open the .docx in Word.Application (COM, ReadOnly).
  3. For every paragraph:
     - y = Range.Information(6) (vertical position relative to page, top of line box)
     - x = Range.Information(5) (horizontal position relative to page)
     - line_count = paragraph line count
  4. For every character in every paragraph:
     - x, y, font name, font size
     - This reveals whether Word exposes ruby annotation characters via
       Range.Characters iteration, AND if so, what their x/y/size is relative
       to the base.
  5. Compute paragraph-to-paragraph dy to derive line height including any
     ruby-induced expansion.

Writes results to:
  pipeline_data/ruby_geometry_measurements.json

Run from repo root:
  python tools/metrics/measure_ruby_geometry.py
"""
import json
import os
import re
import sys
import time
import zipfile
from typing import Any

import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

FIXTURES = [
    "RUBY_V1_basic",
    "RUBY_V2_align_variants",
    "RUBY_V3_hps_variants",
    "RUBY_V4_lineheight",
    "RUBY_V5_linepitch_240",
    "RUBY_V5_linepitch_280",
    "RUBY_V5_linepitch_312",
    "RUBY_V5_linepitch_360",
    "RUBY_V5_linepitch_400",
    "RUBY_V5_linepitch_480",
    "RUBY_V5b_lines_240",
    "RUBY_V5b_lines_280",
    "RUBY_V5b_lines_312",
    "RUBY_V5b_lines_360",
    "RUBY_V5b_lines_400",
    "RUBY_V5b_lines_480",
    "RUBY_V6_hpsRaise",
    "RUBY_V7_wrap",
    "RUBY_V8_extreme_hps",
    "RUBY_V9_combined_grid",
    "RUBY_V10_base_090dpt",
    "RUBY_V10_base_120dpt",
    "RUBY_V10_base_140dpt",
]

DOCX_DIR = os.path.abspath("pipeline_data/docx")
OUT_PATH = os.path.abspath("pipeline_data/ruby_geometry_measurements.json")


def parse_ruby_xml(docx_path: str) -> list[dict]:
    """Extract w:ruby elements + rubyPr children from word/document.xml."""
    with zipfile.ZipFile(docx_path) as z:
        xml = z.read("word/document.xml").decode("utf-8")
    rubies = []
    # Each <w:ruby>...</w:ruby> block
    for m in re.finditer(r"<w:ruby>(.+?)</w:ruby>", xml, re.DOTALL):
        body = m.group(1)
        # rubyPr children
        align = (re.search(r'<w:rubyAlign w:val="([^"]+)"', body) or [None, None])[1]
        hps = (re.search(r'<w:hps w:val="([^"]+)"', body) or [None, None])[1]
        hps_raise = (re.search(r'<w:hpsRaise w:val="([^"]+)"', body) or [None, None])[1]
        hps_base_text = (re.search(r'<w:hpsBaseText w:val="([^"]+)"', body) or [None, None])[1]
        lid = (re.search(r'<w:lid w:val="([^"]+)"', body) or [None, None])[1]
        # rt and rubyBase text
        rt_match = re.search(r"<w:rt>(.+?)</w:rt>", body, re.DOTALL)
        base_match = re.search(r"<w:rubyBase>(.+?)</w:rubyBase>", body, re.DOTALL)
        rt_text = ""
        base_text = ""
        if rt_match:
            for tm in re.finditer(r"<w:t[^>]*>(.+?)</w:t>", rt_match.group(1)):
                rt_text += tm.group(1)
        if base_match:
            for tm in re.finditer(r"<w:t[^>]*>(.+?)</w:t>", base_match.group(1)):
                base_text += tm.group(1)
        # rt and base sz
        rt_sz = None
        if rt_match:
            sz_m = re.search(r'<w:sz w:val="([^"]+)"', rt_match.group(1))
            if sz_m:
                rt_sz = sz_m.group(1)
        base_sz = None
        if base_match:
            sz_m = re.search(r'<w:sz w:val="([^"]+)"', base_match.group(1))
            if sz_m:
                base_sz = sz_m.group(1)
        rubies.append({
            "base_text": base_text,
            "ruby_text": rt_text,
            "rubyAlign": align,
            "hps": hps,
            "hpsRaise": hps_raise,
            "hpsBaseText": hps_base_text,
            "lid": lid,
            "rt_sz": rt_sz,
            "rubyBase_sz": base_sz,
        })
    return rubies


def measure_via_com(word_app, docx_path: str) -> dict:
    """Open doc and measure paragraph + per-char geometry."""
    abs_path = os.path.abspath(docx_path)
    doc = word_app.Documents.Open(abs_path, ReadOnly=True)
    time.sleep(0.4)

    ps = doc.PageSetup
    info = {
        "page_w_pt": ps.PageWidth,
        "page_h_pt": ps.PageHeight,
        "left_margin_pt": ps.LeftMargin,
        "right_margin_pt": ps.RightMargin,
        "top_margin_pt": ps.TopMargin,
        "bottom_margin_pt": ps.BottomMargin,
        "body_w_pt": ps.PageWidth - ps.LeftMargin - ps.RightMargin,
    }
    paragraphs = []
    para_count = doc.Paragraphs.Count
    for pi in range(1, para_count + 1):
        p = doc.Paragraphs(pi)
        rng = p.Range
        try:
            x = rng.Information(5)  # wdHorizontalPositionRelativeToPage
            y = rng.Information(6)  # wdVerticalPositionRelativeToPage
        except Exception:
            x = y = None
        text = rng.Text or ""
        text_clean = text.replace("\r", "").replace("\x07", "")
        para_data = {
            "para_index": pi,
            "text": text_clean,
            "len": len(text_clean),
            "x_pt": x,
            "y_pt": y,
        }

        # Per-char measurement
        chars_data = []
        chars = rng.Characters
        cnt = chars.Count
        for ci in range(1, cnt + 1):
            try:
                c = chars(ci)
                ch = c.Text
                if ch in ("\r", "\x07"):
                    continue
                cx = c.Information(5)
                cy = c.Information(6)
                font = c.Font.Name
                sz = c.Font.Size
                chars_data.append({
                    "i": ci,
                    "ch": ch,
                    "x_pt": cx,
                    "y_pt": cy,
                    "font": font,
                    "size_pt": sz,
                })
            except Exception:
                continue
        para_data["chars"] = chars_data
        # Group chars by y to detect ruby annotation (smaller font, different y)
        y_groups: dict[float, list[dict]] = {}
        for c in chars_data:
            yk = round(c["y_pt"], 1)
            y_groups.setdefault(yk, []).append(c)
        para_data["y_groups"] = [
            {
                "y_pt": yk,
                "char_count": len(g),
                "sizes_pt": sorted(set(c["size_pt"] for c in g)),
                "fonts": sorted(set(c["font"] for c in g)),
                "text": "".join(c["ch"] for c in g),
            }
            for yk, g in sorted(y_groups.items())
        ]
        paragraphs.append(para_data)

    doc.Close(SaveChanges=False)
    return {"page_info": info, "paragraphs": paragraphs}


def main() -> None:
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    results: dict[str, Any] = {
        "_meta": {
            "tool": "tools/metrics/measure_ruby_geometry.py",
            "purpose": "COM + XML measurement of Word ruby (furigana) geometry",
            "fixtures": FIXTURES,
        },
        "fixtures": {},
    }
    try:
        for fname in FIXTURES:
            print(f"\n=== {fname} ===")
            docx_path = os.path.join(DOCX_DIR, fname + ".docx")
            xml_rubies = parse_ruby_xml(docx_path)
            print(f"  XML w:ruby count: {len(xml_rubies)}")
            for r in xml_rubies:
                print(
                    f"    base={r['base_text']!r} rt={r['ruby_text']!r} "
                    f"align={r['rubyAlign']} hps={r['hps']} hpsRaise={r['hpsRaise']} "
                    f"hpsBase={r['hpsBaseText']} rt_sz={r['rt_sz']}"
                )
            com_data = measure_via_com(word, docx_path)
            results["fixtures"][fname] = {
                "xml_rubies": xml_rubies,
                "com": com_data,
            }
            # Brief print
            print(f"  COM paragraphs: {len(com_data['paragraphs'])}")
            for p in com_data["paragraphs"]:
                ygs = p["y_groups"]
                print(
                    f"    P{p['para_index']} y={p['y_pt']:.2f} text={p['text'][:40]!r} "
                    f"y_groups={len(ygs)} sizes_seen={sorted(set(s for g in ygs for s in g['sizes_pt']))}"
                )
                for g in ygs:
                    print(
                        f"      y={g['y_pt']} chars={g['char_count']} "
                        f"sizes={g['sizes_pt']} text={g['text'][:30]!r}"
                    )
    finally:
        word.Quit()

    os.makedirs(os.path.dirname(OUT_PATH), exist_ok=True)
    with open(OUT_PATH, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f"\nWrote {OUT_PATH}")


if __name__ == "__main__":
    main()
