"""Export .docx to PDF with WPS Office (Writer) over COM, one document at a time -- the
WPS column of the Chinese benchmark (2026-11-09, WPS Office 中国版 via winget
Kingsoft.WPSOffice.CN). WPS's Writer object model follows Word's; the export is tried as
ExportAsFixedFormat, then SaveAs2 / SaveAs with wdFormatPDF (17), then ExportPdf.

Usage: python wps_export_pdf.py <truth_pdf_dir> <docx_root> <out_dir>
  for every <type>__<sha16>.pdf in truth_pdf_dir, converts <docx_root>/<type>/<sha16>.docx
  into <out_dir>/<type>__<sha16>.pdf (skips existing). Log: <out_dir>/wps.log
"""
import sys, os, glob, time

sys.stdout.reconfigure(encoding='utf-8')
truth, root, out = sys.argv[1], sys.argv[2], sys.argv[3]
os.makedirs(out, exist_ok=True)
import win32com.client

app = None
for prog in ('KWps.Application', 'Kwps.Application', 'WPS.Application'):
    try:
        app = win32com.client.DispatchEx(prog)
        print('COM', prog)
        break
    except Exception as e:
        print('no', prog, e)
if app is None:
    sys.exit('WPS COM not available')
try:
    app.Visible = False
except Exception:
    pass
try:
    app.DisplayAlerts = 0
except Exception:
    pass


def export(doc, dst):
    for how in ('efx', 'saveas2', 'saveas', 'exportpdf'):
        try:
            if how == 'efx':
                doc.ExportAsFixedFormat(dst, 17)
            elif how == 'saveas2':
                doc.SaveAs2(dst, 17)
            elif how == 'saveas':
                doc.SaveAs(dst, 17)
            else:
                doc.ExportPdf(dst)
            if os.path.exists(dst) and os.path.getsize(dst) > 0:
                return how
        except Exception:
            continue
    return None


log = open(os.path.join(out, 'wps.log'), 'a', encoding='utf8')
pdfs = sorted(glob.glob(os.path.join(truth, '*.pdf')))
try:
    for i, p in enumerate(pdfs, 1):
        stem = os.path.basename(p)[:-4]
        t, h = stem.split('__', 1)
        dst = os.path.abspath(os.path.join(out, stem + '.pdf'))
        if os.path.exists(dst):
            continue
        t0 = time.time()
        try:
            doc = app.Documents.Open(os.path.abspath(os.path.join(root, t, h + '.docx')), False, True)
            how = export(doc, dst)
            doc.Close(False)
            msg = f'[WPS {i:2d}/{len(pdfs)}] {stem} {"OK " + how if how else "FAIL export"} {time.time() - t0:.1f}s'
        except Exception as e:
            msg = f'[WPS {i:2d}/{len(pdfs)}] {stem} FAIL {e}'
        print(msg); log.write(msg + '\n'); log.flush()
finally:
    try:
        app.Quit()
    except Exception:
        pass
