"""S98 focused sweep: font_family_ea × font_size × line (auto rule).

Goal: gather cross-font data for the lineRule=auto + line=N formula
identified in S97. Need to derive Word's base computation per
(font_family, font_size) to design a clean fix.
"""
from __future__ import annotations
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from fuzz_focused import run_sweep


def main():
    variants = []
    # Pin: no v_align, no trHeight (simple cells)
    # Sweep: font_family_ea × font_size × line value (lineRule=auto implicit)
    for ff in ["ＭＳ 明朝", "ＭＳ ゴシック", "游明朝", "游ゴシック"]:
        for fs in [18, 21, 22, 24, 28]:  # 9pt, 10.5pt, 11pt, 12pt, 14pt
            for line in [None, 240, 280, 360, 480]:
                label = f"ff={ff[:4]}_fs={fs}_line={line}"
                p_ov = {"font_family_ea": ff, "font_size": fs, "line": line}
                # No lineRule attr → defaults to auto
                variants.append((label, p_ov, None, None))
    run_sweep("sweep_font_line", variants)


if __name__ == "__main__":
    main()
