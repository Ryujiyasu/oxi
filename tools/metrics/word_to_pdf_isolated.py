"""Convert isolated-test docx files to PDF via Word COM."""
import os
import sys
import glob
import win32com.client as win32

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(SCRIPT_DIR, "docx_tests_isolated")
OUTPUT_DIR = os.path.join(SCRIPT_DIR, "output", "pdfs_isolated")

def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    files = sorted(glob.glob(os.path.join(INPUT_DIR, "*.docx")))
    if not files:
        print("No docx files found")
        sys.exit(1)

    print("Starting Word...")
    word = win32.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0

    ok = 0
    for i, path in enumerate(files, 1):
        name = os.path.basename(path)
        pdf = os.path.join(OUTPUT_DIR, name.replace(".docx", ".pdf"))
        print(f"[{i}/{len(files)}] {name} ... ", end="", flush=True)
        try:
            doc = word.Documents.Open(os.path.abspath(path))
            doc.SaveAs2(os.path.abspath(pdf), FileFormat=17)
            doc.Close(0)
            print("OK")
            ok += 1
        except Exception as e:
            print(f"FAILED: {e}")

    word.Quit()
    print(f"\n{ok}/{len(files)} PDFs written")

if __name__ == "__main__":
    main()
