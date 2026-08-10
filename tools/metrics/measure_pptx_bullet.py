"""Measure bullet rendering in Word PDF (PowerPoint COM oracle for spec: bullet).

Exports the repro deck to PDF and prints per-paragraph COM info + a marker for
the fitz-based measurement to key on.
"""
import sys
import os

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

import win32com.client
import pythoncom


def main():
    pptx_path = os.path.abspath(sys.argv[1] if len(sys.argv) > 1 else "bullet.pptx")
    out_pdf = os.path.abspath(sys.argv[2] if len(sys.argv) > 2 else "bullet.pdf")

    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("PowerPoint.Application")
    try:
        pres = app.Presentations.Open(pptx_path, ReadOnly=True, WithWindow=False)
        try:
            print("slides=%d" % pres.Slides.Count)
            print("width=%s height=%s" % (pres.PageSetup.SlideWidth, pres.PageSetup.SlideHeight))
            for si in range(1, pres.Slides.Count + 1):
                sl = pres.Slides(si)
                print("--- slide %d shapes=%d ---" % (si, sl.Shapes.Count))
                for sh in sl.Shapes:
                    print("shape[%s] type=%s left=%s top=%s w=%s h=%s" % (
                        sh.Name, sh.Type, sh.Left, sh.Top, sh.Width, sh.Height))
                    if sh.HasTextFrame:
                        tf = sh.TextFrame
                        tr = tf.TextRange
                        print("  n_paras=%d text=%r" % (tr.Paragraphs().Count, tr.Text[:60]))
                        for pi in range(1, tr.Paragraphs().Count + 1):
                            p = tr.Paragraphs(pi)
                            fmt = p.ParagraphFormat
                            print("    para%d lvl=%s bullet_visible=%s align=%s " % (
                                pi, p.IndentLevel, fmt.Bullet.Visible, fmt.Alignment))
                            print("      bullet_char=%r indent=%s space_before=%s space_after=%s" % (
                                fmt.Bullet.Character, p.IndentLevel, fmt.SpaceBefore, fmt.SpaceAfter))
            pres.SaveAs(out_pdf, 32)  # ppSaveAsPDF
            print("pdf=%s" % out_pdf)
        finally:
            pres.Close()
    finally:
        app.Quit()
    pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
