"""
Query Word COM directly for resolved line spacing values.
Opens each test docx and reads the actual LineSpacing property.
"""
import os
import json
import glob
import win32com.client as win32

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(SCRIPT_DIR, "docx_tests_isolated")
MANIFEST = os.path.join(INPUT_DIR, "manifest.json")


def main():
    with open(MANIFEST, encoding="utf-8") as f:
        manifest = json.load(f)

    print("Starting Word...")
    word = win32.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0

    results = []

    # Only check single_nogrid files (base line height)
    entries = [e for e in manifest if e["mode"] == "single_nogrid"]

    for i, entry in enumerate(entries, 1):
        path = os.path.join(INPUT_DIR, entry["filename"])
        print(f"[{i}/{len(entries)}] {entry['filename']} ... ", end="", flush=True)

        try:
            doc = word.Documents.Open(os.path.abspath(path))

            for p_idx in range(1, doc.Paragraphs.Count + 1):
                para = doc.Paragraphs(p_idx)
                pf = para.Format

                # LineSpacing: resolved line spacing in points
                # LineSpacingRule: 0=single, 1=1.5, 2=double, 3=atLeast, 4=exactly, 5=multiple
                ls = pf.LineSpacing
                lsr = pf.LineSpacingRule

                # Also get the font info
                rng = para.Range
                font_name = rng.Font.Name
                font_size = rng.Font.Size

                text = rng.Text[:30].strip()

                if p_idx == 1:  # Only need first paragraph
                    results.append({
                        "filename": entry["filename"],
                        "font_id": entry["font_id"],
                        "size_pt": entry["size_pt"],
                        "mode": entry["mode"],
                        "line_spacing_pt": ls,
                        "line_spacing_rule": lsr,
                        "resolved_font": font_name,
                        "resolved_size": font_size,
                    })
                    print(f"ls={ls:.3f}pt rule={lsr} font={font_name} size={font_size}")
                    break

            doc.Close(0)
        except Exception as e:
            print(f"FAILED: {e}")

    word.Quit()

    # Summary
    print("\n\n=== RESOLVED LINE SPACING (single, no grid) ===")
    print(f"{'Font':<12} {'Size':>5} {'LineSpacing':>12} {'Rule':>5}")
    print("-" * 40)
    for r in results:
        print(f"{r['font_id']:<12} {r['size_pt']:>5} {r['line_spacing_pt']:>12.3f} {r['line_spacing_rule']:>5}")

    out = os.path.join(SCRIPT_DIR, "output", "word_line_spacing.json")
    with open(out, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f"\nSaved to {out}")


if __name__ == "__main__":
    main()
