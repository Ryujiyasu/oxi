"""Measure LayoutMode for multiple documents to check no-type docGrid behavior."""
import win32com.client
import time, os, glob

def measure():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    time.sleep(1)

    docx_dir = os.path.abspath("tools/golden-test/documents/docx")
    docs = sorted(glob.glob(os.path.join(docx_dir, "*.docx")))

    results = {"0": 0, "1": 0, "2": 0, "3": 0}
    no_type_modes = {}

    for f in docs:
        name = os.path.splitext(os.path.basename(f))[0]
        try:
            doc = word.Documents.Open(f, ReadOnly=True, Visible=False)
            time.sleep(0.3)
            sec = doc.Sections(1)
            lm = sec.PageSetup.LayoutMode
            results[str(lm)] = results.get(str(lm), 0) + 1

            # Check docGrid type from XML
            import zipfile, xml.etree.ElementTree as ET
            z = zipfile.ZipFile(f)
            doc_xml = z.read('word/document.xml')
            ns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
            root = ET.fromstring(doc_xml)
            grids = root.findall(f'.//{ns}docGrid')
            has_type = False
            for g in grids:
                t = g.get(f'{ns}type', g.get('type', ''))
                if t:
                    has_type = True

            if not has_type:
                key = f"no_type_lm{lm}"
                no_type_modes[key] = no_type_modes.get(key, 0) + 1
                if lm != 0:  # Interesting: no-type with non-zero LayoutMode
                    print(f"  [!] {name}: no-type docGrid but LayoutMode={lm}")

            doc.Close(0)
        except Exception as e:
            print(f"  [ERR] {name}: {e}")
            try:
                word.Documents.Close(0)
            except:
                pass

    print(f"\nLayoutMode distribution:")
    for k, v in sorted(results.items()):
        print(f"  LayoutMode={k}: {v} docs")

    print(f"\nNo-type docGrid LayoutMode distribution:")
    for k, v in sorted(no_type_modes.items()):
        print(f"  {k}: {v} docs")

    word.Quit()

if __name__ == "__main__":
    measure()
