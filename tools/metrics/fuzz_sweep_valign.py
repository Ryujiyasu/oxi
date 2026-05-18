"""S98 focused sweep: vAlign × trHeight × cell content scenarios.

S97 bisect on fuzz_0060 identified `cell v_align="bottom"` as the
causative attribute for a 136pt cell positioning divergence. Goal:
isolate the exact behavior.
"""
from __future__ import annotations
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from fuzz_focused import run_sweep


def main():
    variants = []
    # Pin: short content, default font, no spacing
    # Sweep: v_align × tr_height × h_rule
    for va in [None, "top", "center", "bottom"]:
        for trh in [None, 437, 658, 1500]:
            for hr in [None, "atLeast", "exact"]:
                if trh is None and hr is not None:
                    continue  # h_rule meaningless without trHeight
                label = f"va={va}_trH={trh}_hR={hr}"
                cell_ov = {"v_align": va}
                row_ov = {"tr_height": trh, "h_rule": hr}
                variants.append((label, {}, cell_ov, row_ov))
    run_sweep("sweep_valign", variants)


if __name__ == "__main__":
    main()
