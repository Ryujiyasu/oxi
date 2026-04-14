"""Measure line height for oversized P0 in linesAndChars mode.

Sets LayoutMode=2 and LinesPage via COM, then measures P0 Y positions.
linePitch is controlled by LinesPage (lines per page).
"""
import win32com.client, time, sys, json, os

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False

RESULTS = []

def measure(font, size, lines_page=None):
    """Create doc with linesAndChars grid, measure P0→P1 gap."""
    doc = word.Documents.Add()
    time.sleep(0.2)

    ps = doc.PageSetup
    ps.TopMargin = 56.7
    ps.BottomMargin = 56.7
    ps.LeftMargin = 42.55
    ps.RightMargin = 42.55

    # Set linesAndChars mode
    try:
        ps.LayoutMode = 2  # wdLayoutModeLineGrid + CharGrid
    except Exception:
        pass

    if lines_page:
        try:
            ps.LinesPage = lines_page
        except Exception:
            pass

    rng = doc.Range()
    rng.InsertAfter("ABCDE\r\n")
    rng.InsertAfter("normal\r\n")
    rng.InsertAfter("line3")

    p1 = doc.Paragraphs(1)
    p1.Range.Font.Name = font
    p1.Range.Font.Size = size
    p1.Format.SpaceBefore = 0
    p1.Format.SpaceAfter = 0
    p1.Format.LineSpacingRule = 0  # Single

    p2 = doc.Paragraphs(2)
    p2.Range.Font.Name = font
    p2.Range.Font.Size = 10.5
    p2.Format.SpaceBefore = 0
    p2.Format.SpaceAfter = 0

    p3 = doc.Paragraphs(3)
    p3.Range.Font.Name = font
    p3.Range.Font.Size = 10.5
    p3.Format.SpaceBefore = 0
    p3.Format.SpaceAfter = 0

    time.sleep(0.3)

    # Read actual grid settings
    lm = ps.LayoutMode
    lp = None
    try:
        lp = ps.LinesPage
    except Exception:
        pass

    # Compute actual linePitch from LinesPage
    content_h = ps.PageHeight - ps.TopMargin - ps.BottomMargin
    actual_pitch = content_h / lp if lp and lp > 0 else None

    y1 = p1.Range.Information(6)
    y2 = p2.Range.Information(6)
    y3 = p3.Range.Information(6)

    result = {
        "font": font, "size": size,
        "layout_mode": lm,
        "lines_page": lp,
        "pitch_pt": round(actual_pitch, 4) if actual_pitch else None,
        "top_margin": ps.TopMargin,
        "P0_y": y1, "P1_y": y2, "P2_y": y3,
        "gap01": round(y2 - y1, 2),
        "gap12": round(y3 - y2, 2),
        "P0_offset": round(y1 - ps.TopMargin, 2),
    }
    RESULTS.append(result)

    doc.Close(SaveChanges=False)
    return result


# Part 1: various fonts and sizes with default LinesPage
print("=== Part 1: linesAndChars, default LinesPage ===")
print(f"{'font':<16} {'sz':<5} {'LM':<3} {'LP':<4} {'pitch':<7} {'P0_y':<7} {'P1_y':<7} {'gap01':<7} {'P0_ofs':<7}")

for font in ["ＭＳ ゴシック", "ＭＳ 明朝", "Yu Gothic", "Yu Mincho", "Meiryo"]:
    for size in [10.5, 12, 14, 16, 18, 20, 24, 28]:
        try:
            r = measure(font, size)
            print(f"{font:<16} {size:<5} {r['layout_mode']:<3} {r['lines_page'] or '?':<4} {r['pitch_pt'] or '?':<7} {r['P0_y']:<7.2f} {r['P1_y']:<7.2f} {r['gap01']:<7.2f} {r['P0_offset']:<7.2f}")
        except Exception as e:
            print(f"{font:<16} {size:<5} ERROR: {e}")

# Part 2: ＭＳ ゴシック 20pt with explicit LinesPage (controls pitch)
print("\n=== Part 2: ＭＳ ゴシック 20pt, varying LinesPage ===")
for lp in [20, 25, 30, 35, 40, 44, 50]:
    try:
        r = measure("ＭＳ ゴシック", 20, lines_page=lp)
        print(f"  LP={lp:2d} pitch={r['pitch_pt']:.4f} P0_y={r['P0_y']:.2f} P1_y={r['P1_y']:.2f} gap={r['gap01']:.2f} P0_ofs={r['P0_offset']:.2f}")
    except Exception as e:
        print(f"  LP={lp}: ERROR: {e}")

# Part 3: Match 1ec1 exactly - 44 lines/page gives ~17.85pt pitch
print("\n=== Part 3: Match 1ec1 (LinesPage for ~17.85pt pitch) ===")
# content_h = 841.9 - 56.7 - 56.7 = 728.5
# 728.5 / 17.85 ≈ 40.8 → try 40, 41
for lp in [38, 39, 40, 41, 42]:
    try:
        r = measure("ＭＳ ゴシック", 20, lines_page=lp)
        print(f"  LP={lp:2d} pitch={r['pitch_pt']:.4f} P0_y={r['P0_y']:.2f} P1_y={r['P1_y']:.2f} gap={r['gap01']:.2f}")
    except Exception as e:
        print(f"  LP={lp}: ERROR: {e}")

# Save
out_path = "tools/metrics/output/linesAndChars_oversized.json"
os.makedirs(os.path.dirname(out_path), exist_ok=True)
with open(out_path, "w", encoding="utf-8") as f:
    json.dump(RESULTS, f, indent=2, ensure_ascii=False)
print(f"\nSaved {len(RESULTS)} results to {out_path}")

word.Quit()
