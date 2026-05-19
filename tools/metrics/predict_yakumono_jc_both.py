"""Session 114 — predict Word's jc=both yakumono compression algorithm
from the 52-variant grid data.

Hypothesis (from S113 analysis):
  1. Word computes natural line width = sum of chars at (fs + 2*cs_pt) widths
     (balanceSBDB doubles cs for fullwidth chars)
  2. Word picks max N chars that fit on L1 (greedy max-fit with '．','，','。'
     compressed to half-width floor)
  3. If natural_line > budget: compression amount = overflow
     - Distributed across '．','，','。' priority chars first
  4. Each char advance snapped to 15tw (0.75pt) grid

Test: for each MS Mincho variant, predict '．' advance and compare to
COM-measured ground truth.
"""
import os
import sys
import io
import json

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
IN_JSON = os.path.normpath(os.path.join(REPO, "tools/metrics/yakumono_grid/results.json"))

# Cell config (constant across grid)
CELL_DXA = 1968
CELL_MAR_DXA = 12  # left + right both
IND_LEFT_TW = 215
IND_RIGHT_TW = 76
HANGING_TW = 192
# L1 budget (with hanging indent applied):
# = cell_width - 2*cellmar - max(ind_left - hanging, 0) - ind_right
def l1_budget_pt():
    cw = CELL_DXA / 20.0
    cm = 2 * CELL_MAR_DXA / 20.0
    eff_left = max(IND_LEFT_TW - HANGING_TW, 0) / 20.0
    eff_right = IND_RIGHT_TW / 20.0
    return cw - cm - eff_left - eff_right


def snap_15tw(pt: float) -> float:
    """Snap to 15tw (0.75pt) grid. Round half-up."""
    tw = pt * 20.0
    snapped = round(tw / 15.0) * 15.0
    return snapped / 20.0


def predict(fs: float, cs_tw: int, n_chars: int):
    """Predict '．' advance given fs, cs (twips), L1 char count.

    Returns (predicted_punct_adv, predicted_kanji_adv, natural_line_width, overflow).
    """
    cs_pt = cs_tw / 20.0
    # Natural char width with balanceSBDB doubling:
    #   advance = font_size + 2 * cs_pt  (for fullwidth chars)
    nat_char_pt = fs + 2 * cs_pt
    # Natural line width = n_chars × nat_char_pt
    nat_line = n_chars * nat_char_pt
    budget = l1_budget_pt()
    overflow = nat_line - budget

    if overflow > 0:
        # Compression: '．' absorbs all overflow (only 1 punct in test content)
        compressed_punct = nat_char_pt - overflow
        # Snap to 15tw
        compressed_punct = snap_15tw(compressed_punct)
        # Floor: don't compress below fs/2 (half-width)
        compressed_punct = max(compressed_punct, fs / 2.0)
        return compressed_punct, nat_char_pt, nat_line, overflow
    else:
        # Underflow: '．' stays at natural (or maybe expands — see analysis)
        return snap_15tw(nat_char_pt), nat_char_pt, nat_line, overflow


def main():
    with open(IN_JSON, "r", encoding="utf-8") as f:
        data = json.load(f)

    print(f"L1 budget = {l1_budget_pt():.3f}pt")
    print()
    print(f"{'variant':<35} {'fs':>4} {'cs':>4} {'N':>3} {'nat':>6} {'budget':>7} {'overflow':>8} | {'p_pred':>7} {'p_actual':>8} {'diff':>6}")
    print("-" * 110)

    abs_errors = []
    for sz in [18, 21, 24, 28]:
        for cs in [-5, -9, -15, -20]:
            name = f"g_msmincho_sz{sz}_cs{cs}_dotF"
            rec = data.get(name)
            if not rec or "error" in rec:
                continue
            fs = sz / 2.0
            n_chars = rec["lines"][0]["n_chars"] if rec.get("lines") else 0
            p_actual = rec.get("dot_advance_pt")
            p_pred, kn, nl, ov = predict(fs, cs, n_chars)
            diff = (p_pred - p_actual) if p_actual is not None else None
            diff_s = f"{diff:+.2f}" if diff is not None else "?"
            abs_errors.append(abs(diff) if diff is not None else 0.0)
            print(f"{name:<35} {fs:>4.1f} {cs:>4} {n_chars:>3} "
                  f"{kn:>6.2f} {l1_budget_pt():>7.2f} {ov:>+8.2f} | "
                  f"{p_pred:>7.2f} {p_actual:>8.2f} {diff_s:>6}")

    print(f"\nMS Mincho — mean abs error: {sum(abs_errors)/len(abs_errors):.3f}pt")
    print(f"MS Mincho — max abs error : {max(abs_errors):.3f}pt")
    print(f"MS Mincho — within 0.25pt: {sum(1 for e in abs_errors if e <= 0.25)}/{len(abs_errors)}")
    print(f"MS Mincho — within 0.75pt: {sum(1 for e in abs_errors if e <= 0.75)}/{len(abs_errors)}")

    # Test cross-font at (fs=10.5, cs=-9, '．') for each font
    print("\n=== Cross-font at (fs=10.5, cs=-9, '．') ===")
    print(f"{'font':>10} | {'N':>3} {'overflow':>8} | {'p_pred':>7} {'p_actual':>8} {'diff':>6}")
    for slug in ["msmincho", "msgothic", "meiryo", "yugothic", "yumincho"]:
        name = f"g_{slug}_sz21_cs-9_dotF"
        rec = data.get(name)
        if not rec or "error" in rec:
            print(f"{slug:>10}: missing")
            continue
        n_chars = rec["lines"][0]["n_chars"]
        p_actual = rec.get("dot_advance_pt")
        p_pred, kn, nl, ov = predict(10.5, -9, n_chars)
        diff = p_pred - p_actual if p_actual else None
        diff_s = f"{diff:+.2f}" if diff is not None else "?"
        print(f"{slug:>10} | {n_chars:>3} {ov:>+8.2f} | {p_pred:>7.2f} {p_actual:>8.2f} {diff_s:>6}")

    # COMMA / KUTEN variants on MS Mincho
    print("\n=== '，' and '。' (MS Mincho) — should match '．' ===")
    print(f"{'variant':<35} {'p_actual':>8} {'p_pred':>7} {'diff':>6}")
    for sz in [18, 21, 24, 28]:
        for cs in [-5, -9, -15, -20]:
            for slug, key in [("commaF", "comma_advance_pt"), ("kuten", "kuten_advance_pt")]:
                name = f"g_msmincho_sz{sz}_cs{cs}_{slug}"
                rec = data.get(name)
                if not rec or "error" in rec:
                    continue
                n_chars = rec["lines"][0]["n_chars"]
                p_actual = rec.get(key)
                p_pred, _, _, _ = predict(sz / 2.0, cs, n_chars)
                diff = p_pred - p_actual if p_actual else None
                diff_s = f"{diff:+.2f}" if diff is not None else "?"
                print(f"{name:<35} {p_actual or 0:>8.2f} {p_pred:>7.2f} {diff_s:>6}")


if __name__ == "__main__":
    main()
