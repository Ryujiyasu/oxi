"""Check what font Word actually uses when Arial Unicode MS is missing."""
import win32com.client, time, os, sys
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False

path = os.path.abspath("tools/golden-test/documents/docx/0e7af1ae8f21_20230331_resources_open_data_contract_sample_00.docx")
doc = word.Documents.Open(path, ReadOnly=True)
time.sleep(0.5)

# Sample first 10 chars of various paragraphs
for pi in [1, 5, 10, 50, 100, 200, 300]:
    try:
        p = doc.Paragraphs(pi)
        chars = p.Range.Characters
        if chars.Count == 0:
            continue
        c1 = chars(1)
        text = p.Range.Text[:30]
        print(f"P{pi}: text={text!r}")
        print(f"  Font.Name (XML spec)={c1.Font.Name!r} size={c1.Font.Size}")
        # Word's actual rendered font may differ; check via different attrs
        try:
            print(f"  Font.NameAscii={c1.Font.NameAscii!r}")
            print(f"  Font.NameFarEast={c1.Font.NameFarEast!r}")
            print(f"  Font.NameOther={c1.Font.NameOther!r}")
        except Exception as e:
            print(f"  err: {e}")
    except Exception as e:
        print(f"P{pi}: skip ({e})")

# Also check FontSubstitutes mechanism
print("\nFontSubstitutes for 'Arial Unicode MS':")
try:
    # Word doesn't expose this directly via COM. Try registry or fallback heuristic.
    # Just check if Word recognizes the font name
    fns = list(word.FontNames)
    if "Arial Unicode MS" in fns:
        print("  Available")
    else:
        print("  NOT available — Word will substitute")
except Exception as e:
    print(f"  err: {e}")

doc.Close(SaveChanges=False)
word.Quit()
