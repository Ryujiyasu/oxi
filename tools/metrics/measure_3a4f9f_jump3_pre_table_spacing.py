"""Measure 3a4f9f wi=1028-1042 paragraph y positions + line spacing properties.

Targets the +5pt drift introduced between wi=1032 and wi=1033 (Oxi adds 21pt
between two empty body paragraphs while Word adds only 16pt). This is the
cumulative +6pt drift that, combined with Oxi's tighter line-wrap inside
the big non-floating table, causes Sub-jump 3a (wi=1036's "る。" overflow
to a new page).

Per CLAUDE.md R30 fix: Information(6) on a paragraph range returns the
ACTIVE-END position, not the start. Use `doc.Range(rng.Start, rng.Start)`
to get true start-y. Apply this here.

Output: pipeline_data/ra_manual_measurements/3a4f9f_wi1028_1042_pretable_spacing.json
"""
from __future__ import annotations

import json
import os
import sys
import time

import win32com.client


REPO = r"c:\Users\ryuji\oxi-main"
DOCX = os.path.join(REPO, "tools", "golden-test", "documents", "docx",
                    "3a4f9fbe1a83_001620506.docx")
OUT = os.path.join(REPO, "pipeline_data", "ra_manual_measurements",
                   "3a4f9f_wi1028_1042_pretable_spacing.json")

WD_VERTICAL_POSITION_RELATIVE_TO_PAGE = 6
WD_HORIZONTAL_POSITION_RELATIVE_TO_PAGE = 5
WD_ACTIVE_END_PAGE = 3

# Word LineSpacingRule constants
LSR_NAMES = {
    0: "wdLineSpaceSingle",
    1: "wdLineSpace1pt5",
    2: "wdLineSpaceDouble",
    3: "wdLineSpaceAtLeast",
    4: "wdLineSpaceExactly",
    5: "wdLineSpaceMultiple",
}


def main():
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(DOCX, ReadOnly=True)
        time.sleep(2)

        n_paras = doc.Paragraphs.Count
        print(f"Total paragraphs: {n_paras}")
        # Target wi=1028..1042 (1-indexed)
        results = []
        for wi in range(1028, 1043):
            if wi > n_paras:
                break
            para = doc.Paragraphs(wi)
            rng = para.Range
            # R30 fix: collapsed start
            start_rng = doc.Range(rng.Start, rng.Start)
            try:
                start_y = start_rng.Information(WD_VERTICAL_POSITION_RELATIVE_TO_PAGE)
                start_x = start_rng.Information(WD_HORIZONTAL_POSITION_RELATIVE_TO_PAGE)
                page = start_rng.Information(WD_ACTIVE_END_PAGE)
            except Exception as e:
                start_y = None; start_x = None; page = None
                print(f"  wi={wi}: Information() failed: {e}")

            text = (rng.Text or "").rstrip("\r\n").rstrip("\x07")
            text_preview = text[:30] + "..." if len(text) > 30 else text

            fmt = para.Format
            try:
                line_spacing = fmt.LineSpacing
            except Exception:
                line_spacing = None
            try:
                line_spacing_rule = fmt.LineSpacingRule
                lsr_name = LSR_NAMES.get(line_spacing_rule, f"unknown({line_spacing_rule})")
            except Exception:
                line_spacing_rule = None; lsr_name = None
            try:
                space_before = fmt.SpaceBefore
                space_after = fmt.SpaceAfter
            except Exception:
                space_before = None; space_after = None

            try:
                style_name = para.Style.NameLocal
            except Exception:
                style_name = None
            try:
                first_line_indent = fmt.FirstLineIndent
            except Exception:
                first_line_indent = None

            try:
                in_cell = bool(rng.Information(12))  # wdWithInTable
            except Exception:
                in_cell = None

            row = {
                "wi": wi,
                "page": page,
                "start_y": start_y,
                "start_x": start_x,
                "text_preview": text_preview,
                "text_len": len(text),
                "in_table": in_cell,
                "style": style_name,
                "line_spacing": line_spacing,
                "line_spacing_rule": line_spacing_rule,
                "line_spacing_rule_name": lsr_name,
                "space_before": space_before,
                "space_after": space_after,
                "first_line_indent": first_line_indent,
            }
            results.append(row)
            print(f"  wi={wi:>4} pg={page!s:>3} y={start_y!s:>7} x={start_x!s:>7} "
                  f"in_tbl={in_cell!s:>5} style={style_name!s:>20} "
                  f"ls={line_spacing!s:>6} rule={lsr_name!s:>20} "
                  f"sb={space_before!s:>5} sa={space_after!s:>5} "
                  f"text={text_preview!r}")

        # Compute Δy between consecutive paragraphs
        print("\n=== Δy between consecutive paragraphs ===")
        deltas = []
        for i in range(1, len(results)):
            prev = results[i-1]; cur = results[i]
            if prev["start_y"] is None or cur["start_y"] is None:
                continue
            dy = cur["start_y"] - prev["start_y"]
            deltas.append({
                "wi_prev": prev["wi"],
                "wi_cur": cur["wi"],
                "delta_y": round(dy, 3),
                "page_change": cur["page"] - prev["page"],
                "prev_style": prev["style"],
                "cur_style": cur["style"],
            })
            print(f"  wi={prev['wi']}->wi={cur['wi']}: Δy={dy:>+7.2f}  "
                  f"(prev_style={prev['style']!r}, cur_style={cur['style']!r})")

        os.makedirs(os.path.dirname(OUT), exist_ok=True)
        with open(OUT, "w", encoding="utf-8") as f:
            json.dump({"paragraphs": results, "deltas": deltas}, f, ensure_ascii=False, indent=2)
        print(f"\nSaved to {OUT}")

    finally:
        try:
            doc.Close(SaveChanges=False)
        except Exception:
            pass
        word.Quit()


if __name__ == "__main__":
    main()
