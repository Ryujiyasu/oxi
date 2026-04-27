"""Probe alternative COM techniques to expose ruby annotation positions.

Approaches to test:
  1. Range.Words iteration — does iterating words (vs characters) reveal ruby?
  2. Range.WordOpenXML — get OOXML for a small range, see ruby structure
  3. Selection.MoveRight with wdCharacter unit — does navigation enter ruby?
  4. Document.GetCrossReferenceItems / Document.Bookmarks for ruby-as-field
  5. Range.Information(13) wdActiveEndAdjustedPageNumber and other less-used
     Information types
  6. Range.End / Range.Start vs ruby placeholder position

Goal: find ANY way to programmatically extract the X/Y of ruby annotation
characters (currently invisible to Range.Characters).
"""
import os
import sys
import time

import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOCX_PATH = os.path.abspath("pipeline_data/docx/RUBY_V2_align_variants.docx")


def main() -> None:
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    doc = word.Documents.Open(DOCX_PATH, ReadOnly=True)
    time.sleep(0.4)

    print("=" * 60)
    print(f"Probing ruby exposure on {os.path.basename(DOCX_PATH)}")
    print(f"Total paragraphs: {doc.Paragraphs.Count}")
    print("=" * 60)

    # First paragraph: "[center]: 特定 [ruby とくてい] の単語にルビ"
    p1 = doc.Paragraphs(1)
    rng = p1.Range

    # ── Approach 1: Range.Words iteration ──
    print("\n[1] Range.Words iteration:")
    words = rng.Words
    print(f"  word count: {words.Count}")
    for wi in range(1, min(words.Count + 1, 15)):
        try:
            w = words(wi)
            wt = w.Text
            wx = w.Information(5)
            wy = w.Information(6)
            print(f"  W{wi}: text={wt!r} x={wx} y={wy}")
        except Exception as e:
            print(f"  W{wi}: ERROR {e}")

    # ── Approach 2: Range.WordOpenXML ──
    print("\n[2] Range.WordOpenXML (first 600 chars):")
    try:
        xml_full = rng.WordOpenXML
        # Print interesting parts
        if "<w:ruby>" in xml_full:
            ruby_start = xml_full.index("<w:ruby>")
            ruby_end = xml_full.index("</w:ruby>", ruby_start) + len("</w:ruby>")
            print(f"  ruby element found at offset {ruby_start}:")
            print(f"  {xml_full[ruby_start:ruby_end]}")
        else:
            print("  no <w:ruby> in XML")
    except Exception as e:
        print(f"  ERROR: {e}")

    # ── Approach 3: Selection MoveRight character-by-character ──
    print("\n[3] Selection MoveRight character-by-character:")
    rng.Select()
    sel = word.Selection
    sel.HomeKey(Unit=5)  # wdLine
    for step in range(20):
        try:
            x = sel.Information(5)
            y = sel.Information(6)
            txt = sel.Text or ""
            print(f"  step={step} text={txt!r} x={x} y={y}")
            sel.MoveRight(Unit=1, Count=1, Extend=0)  # wdCharacter
        except Exception as e:
            print(f"  step={step}: ERROR {e}")
            break

    # ── Approach 4: Range duplicate + collapsing to find sub-positions ──
    print("\n[4] Range.Duplicate + per-position SetRange:")
    text = rng.Text
    print(f"  full text: {text!r}")
    print(f"  Range.Start={rng.Start} End={rng.End}")
    # Try to set range to specific char positions and check x/y
    for offset in range(0, min(rng.End - rng.Start, 16)):
        try:
            r2 = rng.Duplicate
            r2.SetRange(Start=rng.Start + offset, End=rng.Start + offset + 1)
            t = r2.Text
            x = r2.Information(5)
            y = r2.Information(6)
            print(f"  off={offset} text={t!r} x={x} y={y}")
        except Exception as e:
            print(f"  off={offset}: ERROR {e}")
            break

    # ── Approach 5: Range.Comments / Footnotes / Endnotes / Fields ──
    print("\n[5] Range collections (Fields, Footnotes, etc.):")
    for attr in ["Fields", "Footnotes", "Endnotes", "FormFields", "ContentControls",
                 "Hyperlinks", "Bookmarks"]:
        try:
            coll = getattr(rng, attr)
            print(f"  rng.{attr}: count={coll.Count}")
        except Exception as e:
            print(f"  rng.{attr}: ERROR {e}")

    # ── Approach 6: Iterate doc.StoryRanges ──
    print("\n[6] doc.StoryRanges:")
    try:
        for sr in doc.StoryRanges:
            try:
                stype = sr.StoryType
                print(f"  story type={stype} text={sr.Text[:40]!r}")
            except Exception as e:
                print(f"  story: ERROR {e}")
    except Exception as e:
        print(f"  StoryRanges: ERROR {e}")

    doc.Close(SaveChanges=False)
    word.Quit()


if __name__ == "__main__":
    main()
