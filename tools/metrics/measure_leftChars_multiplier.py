"""Measure what multiplier Word uses for leftChars indent.

Creates docs with linesAndChars grid + leftChars indent, measures actual
X position to determine if multiplier = fontSize, charPitch, or something else.
"""
import win32com.client, time, sys, json, os

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False

RESULTS = []

def measure_indent(char_space_tw=0, left_chars_val=200, font_size=10.5):
    """Create doc with linesAndChars, set leftChars, measure X position."""
    doc = word.Documents.Add()
    time.sleep(0.2)

    ps = doc.PageSetup
    ps.TopMargin = 56.7
    ps.BottomMargin = 56.7
    ps.LeftMargin = 42.55
    ps.RightMargin = 42.55

    try:
        ps.LayoutMode = 2  # linesAndChars
    except Exception:
        pass

    # Set default font via range (avoid Styles encoding issue)
    rng = doc.Range()
    rng.InsertAfter("ABCDE\r\n")
    rng.InsertAfter("FGHIJ")

    rng2 = doc.Range()
    rng2.Font.Name = "MS Gothic"
    rng2.Font.Size = font_size

    p1 = doc.Paragraphs(1)
    p1.Range.Font.Name = "MS Gothic"
    p1.Range.Font.Size = font_size

    # Set leftChars via CharacterUnitLeftIndent
    try:
        p1.Format.CharacterUnitLeftIndent = left_chars_val / 100.0
    except Exception as e:
        print(f"  CharacterUnitLeftIndent error: {e}")

    p2 = doc.Paragraphs(2)
    p2.Range.Font.Name = "MS Gothic"
    p2.Range.Font.Size = font_size

    time.sleep(0.3)

    # Read actual layout mode and grid
    lm = ps.LayoutMode
    x1 = p1.Range.Information(5)  # wdHorizontalPositionRelativeToPage
    x2 = p2.Range.Information(5)
    y1 = p1.Range.Information(6)

    # Also read computed indent
    left_indent_pt = p1.Format.LeftIndent
    char_indent = p1.Format.CharacterUnitLeftIndent

    content_w = ps.PageWidth - ps.LeftMargin - ps.RightMargin
    left_margin = ps.LeftMargin

    result = {
        "font_size": font_size,
        "char_space_tw": char_space_tw,
        "left_chars_val": left_chars_val,
        "layout_mode": lm,
        "x1": x1,
        "x2": x2,
        "x_diff": round(x1 - x2, 2),
        "left_indent_pt": left_indent_pt,
        "char_unit_indent": char_indent,
        "left_margin": left_margin,
        "content_w": content_w,
    }
    RESULTS.append(result)

    doc.Close(SaveChanges=False)
    return result


print("=== leftChars multiplier measurement ===")
print(f"{'fontSize':<9} {'cs_tw':<7} {'lChars':<7} {'x_p1':<8} {'x_p2':<8} {'xdiff':<7} {'leftIndent_pt':<14} {'charUnit'}")

# Test 1: default font size 10.5, no charSpace
for lc in [100, 200, 300]:
    r = measure_indent(char_space_tw=0, left_chars_val=lc, font_size=10.5)
    print(f"{10.5:<9} {0:<7} {lc:<7} {r['x1']:<8.2f} {r['x2']:<8.2f} {r['x_diff']:<7.2f} {r['left_indent_pt']:<14.2f} {r['char_unit_indent']}")

# Test 2: default font size 12, no charSpace
print()
for lc in [100, 200, 300]:
    r = measure_indent(char_space_tw=0, left_chars_val=lc, font_size=12)
    print(f"{12:<9} {0:<7} {lc:<7} {r['x1']:<8.2f} {r['x2']:<8.2f} {r['x_diff']:<7.2f} {r['left_indent_pt']:<14.2f} {r['char_unit_indent']}")

# Test 3: font size 10.5, various charSpace
print()
for lc in [200]:
    for fs in [10.5, 12, 14]:
        r = measure_indent(char_space_tw=0, left_chars_val=lc, font_size=fs)
        # Compute expected multipliers
        left_margin = r['left_margin']
        actual_indent = r['x1'] - left_margin
        multiplier = actual_indent / (lc / 100.0)
        print(f"  fs={fs} lChars={lc}: x={r['x1']:.2f} margin={left_margin:.2f} indent={actual_indent:.2f} => multiplier={multiplier:.4f}")

# Compute charPitch for comparison
print("\n=== charPitch comparison ===")
for fs in [10.5, 12, 14]:
    # Approximate A4 content width
    content_w = 510.2  # A4 - margins
    raw_pitch = fs  # no charSpace
    chars_line = int(content_w / raw_pitch)
    actual_pitch = content_w / chars_line
    print(f"  fs={fs}: raw_pitch={raw_pitch} chars_line={chars_line} actual_pitch={actual_pitch:.4f}")

# Save
out_path = "tools/metrics/output/leftChars_multiplier.json"
os.makedirs(os.path.dirname(out_path), exist_ok=True)
with open(out_path, "w", encoding="utf-8") as f:
    json.dump(RESULTS, f, indent=2, ensure_ascii=False)
print(f"\nSaved {len(RESULTS)} results to {out_path}")

word.Quit()
