"""Measure leftChars multiplier with different Normal style font sizes."""
import win32com.client, time, sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

word = win32com.client.gencache.EnsureDispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False

def measure(normal_size, left_chars):
    doc = word.Documents.Add()
    time.sleep(0.2)

    # Change Normal style font size FIRST
    # Set default font size via rPrDefault
    rng = doc.Range()
    rng.Font.Size = normal_size

    ps = doc.PageSetup
    ps.TopMargin = 56.7
    ps.BottomMargin = 56.7
    ps.LeftMargin = 42.55
    ps.RightMargin = 42.55
    try:
        ps.LayoutMode = 2
    except Exception:
        pass

    rng = doc.Range()
    rng.InsertAfter("ABCDE\r\n")
    rng.InsertAfter("FGHIJ")

    p1 = doc.Paragraphs(1)
    try:
        p1.Format.CharacterUnitLeftIndent = left_chars / 100.0
    except Exception as e:
        print(f"  Error: {e}")
        doc.Close(SaveChanges=False)
        return None

    time.sleep(0.2)

    indent_pt = p1.Format.LeftIndent
    normal_font_size = normal_size

    doc.Close(SaveChanges=False)
    return indent_pt, normal_font_size


print(f"{'NormalSize':<11} {'lChars':<7} {'LeftIndent':<11} {'multiplier'}")
for normal_sz in [8, 9, 10, 10.5, 11, 12, 14, 16, 20]:
    for lc in [100, 200]:
        result = measure(normal_sz, lc)
        if result:
            indent, actual_ns = result
            mult = indent / (lc / 100.0)
            print(f"{normal_sz:<11} {lc:<7} {indent:<11.2f} {mult:.4f}")

word.Quit()
