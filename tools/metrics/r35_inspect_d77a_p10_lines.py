"""Inspect d77a_p10 line widths to verify Word's actual fitting behavior.

For each line in d77a_p10:
  Compute observed_width (sum of adv) — what Word actually placed.
  Compute natural_width (sum of sz for yak, adv for others) — pre-Mech 2.
  Compute content_w (x_end - x_start).

Test: does observed_width fit within content_w? (should always)
Test: does natural_width exceed content_w by exactly the Mech 2 compression saving?
"""
import json
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

YAKUMONO_A = set("（「『【〔｛〈《［")
YAKUMONO_B = set("）」』】〕｝〉》］、。，．—")


def cls(ch):
    if ch in YAKUMONO_A: return "A"
    if ch in YAKUMONO_B: return "B"
    return "X"


def main():
    with open("pipeline_data/r35_data/d77a_p10_per_char_R33_diag.json",
              encoding="utf-8") as f:
        d = json.load(f)
    print(f"{'para':>4} {'line':>4} {'n':>3} {'cw':>7} {'obs':>7} {'nat':>7} "
          f"{'slack_obs':>9} {'slack_nat':>9} {'comp_total':>10} {'n_yak':>5} "
          f"{'n_comp':>6}")
    print("-" * 90)
    for p in d["paragraphs"]:
        for li, line in enumerate(p.get("lines", [])):
            chars = line["chars"]
            n = len(chars)
            x_start = line.get("x_start")
            x_end = line.get("x_end")
            cw = (x_end - x_start) if x_start is not None and x_end is not None else None
            obs_sum = sum((c.get("adv") or 0) for c in chars)
            nat_sum = 0
            n_yak = 0
            n_comp = 0
            for c in chars:
                ch = c["ch"]
                sz = c.get("sz") or 12.0
                adv = c.get("adv") or 0
                cl = cls(ch)
                if cl in ("A", "B"):
                    nat_sum += sz
                    n_yak += 1
                    if sz and adv / sz < 0.85:
                        n_comp += 1
                else:
                    nat_sum += adv
            slack_obs = (cw - obs_sum) if cw is not None else None
            slack_nat = (cw - nat_sum) if cw is not None else None
            comp_total = nat_sum - obs_sum
            print(f"{p['para']:>4} {li+1:>4} {n:>3} "
                  f"{cw if cw is None else f'{cw:>7.2f}'} "
                  f"{obs_sum:>7.2f} {nat_sum:>7.2f} "
                  f"{slack_obs if slack_obs is None else f'{slack_obs:>+9.2f}'} "
                  f"{slack_nat if slack_nat is None else f'{slack_nat:>+9.2f}'} "
                  f"{comp_total:>10.2f} {n_yak:>5} {n_comp:>6}")


if __name__ == "__main__":
    main()
