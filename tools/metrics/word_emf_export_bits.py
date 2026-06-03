"""Export Word pages as EMF via Range.EnhMetaFileBits (NO clipboard).

CopyAsPicture fails ("no selection") on ranges spanning tables (b35 etc.).
Range.EnhMetaFileBits returns the EMF bytes directly as a memoryview — works on
tables. Usage: python word_emf_export_bits.py <docx> <out_prefix> [page|all]
Produces out_prefix_pN.emf. cp932-safe (no Japanese in code; ASCII output)."""
import win32com.client
import sys
import os
import pythoncom


def export(docx, prefix, which='all'):
    docx = os.path.abspath(docx)
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(docx, ReadOnly=True)
        total = doc.ComputeStatistics(2)  # wdStatisticPages
        print("Pages:", total)
        pages = range(1, total + 1) if which == 'all' else [int(which)]
        for pg in pages:
            try:
                rng = doc.GoTo(1, 2, pg)  # wdGoToPage, wdGoToAbsolute
                if pg < total:
                    nxt = doc.GoTo(1, 2, pg + 1)
                    rng.End = nxt.Start
                else:
                    rng.End = doc.Content.End
                bits = rng.EnhMetaFileBits  # memoryview of EMF bytes (no clipboard)
                data = bytes(bits)
                out = "%s_p%d.emf" % (prefix, pg)
                with open(out, 'wb') as f:
                    f.write(data)
                print("  Saved %s (%d bytes)" % (out, len(data)))
            except Exception as e:
                print("  Page %d FAILED: %s" % (pg, str(e)[:90]))
        doc.Close(False)
    finally:
        word.Quit()


if __name__ == '__main__':
    if len(sys.argv) < 3:
        print("Usage: word_emf_export_bits.py <docx> <out_prefix> [page|all]")
        sys.exit(1)
    export(sys.argv[1], sys.argv[2], sys.argv[3] if len(sys.argv) > 3 else 'all')
