# -*- coding: utf-8 -*-
"""Which face does Word actually paint a given piece of text in?

`Font.NameFarEast` answers with the theme TOKEN ("+本文のフォント - 日本語")
whenever the run inherits from the theme, so COM cannot say which face that
token resolved to. The exported PDF can: it names the embedded subset per span.

    python _word_pdf_faces.py <docx> [substring ...]

With no substring it lists every face on page 1 with a sample of its text.
"""
import os
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")


def main():
    import fitz
    import win32com.client as w

    src = os.path.abspath(sys.argv[1])
    wants = sys.argv[2:]
    out = os.path.join(os.environ.get("TEMP", "."), "_word_pdf_faces.pdf")
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    try:
        d = app.Documents.Open(src, ReadOnly=True, AddToRecentFiles=False)
        try:
            d.ExportAsFixedFormat(out, 17)
        finally:
            d.Close(False)
    finally:
        app.Quit()

    doc = fitz.open(out)
    if wants:
        for pi in range(doc.page_count):
            for b in doc[pi].get_text("dict")["blocks"]:
                for ln in b.get("lines", []):
                    line = "".join(s["text"] for s in ln["spans"])
                    if not any(wnt in line for wnt in wants):
                        continue
                    for s in ln["spans"]:
                        if s["text"].strip():
                            print(f"  p{pi+1} {s['font']:<26} {s['size']:5.2f}  {s['text'][:34]!r}")
        return
    seen = {}
    for b in doc[0].get_text("dict")["blocks"]:
        for ln in b.get("lines", []):
            for s in ln["spans"]:
                if s["text"].strip():
                    seen.setdefault((s["font"], round(s["size"], 2)), s["text"][:24])
    for (font, size), sample in sorted(seen.items()):
        print(f"  {font:<26} {size:5.2f}  {sample!r}")


if __name__ == "__main__":
    main()
