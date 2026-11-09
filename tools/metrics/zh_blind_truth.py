"""Word truth for a frozen Chinese blind set, exported TWICE (FROZEN_BLIND_ZH_A.md): the
first pass lets Word fetch any cloud font it needs (a font downloading during the export
leaves a substitute in that PDF); the second pass is the truth.

Usage: python zh_blind_truth.py A
Reads pipeline_data/zh_benchmark/_final_zhBlind<A>50.json; writes
pipeline_data/zh_benchmark/ssim_zhBlind<A>50/{word_pdf_run1,word_pdf}/<type>__<sha16>.pdf,
word.log, cloudfonts_{before,after}.txt, _word_version.json
"""
import sys, os, json, time
sys.stdout.reconfigure(encoding='utf-8')
REPO = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..'))
L = sys.argv[1]
ids = json.load(open(os.path.join(REPO, 'pipeline_data', 'zh_benchmark', f'_final_zhBlind{L}50.json'), encoding='utf8'))
out = os.path.join(REPO, 'pipeline_data', 'zh_benchmark', f'ssim_zhBlind{L}50')
cloud = os.path.join(os.environ['LOCALAPPDATA'], 'Microsoft', 'FontCache', '4', 'CloudFonts')
os.makedirs(out, exist_ok=True)
open(os.path.join(out, 'cloudfonts_before.txt'), 'w', encoding='utf8').write('\n'.join(sorted(os.listdir(cloud))))
import win32com.client
log = open(os.path.join(out, 'word.log'), 'a', encoding='utf8')
for run, sub in ((1, 'word_pdf_run1'), (2, 'word_pdf')):
    os.makedirs(os.path.join(out, sub), exist_ok=True)
    w = win32com.client.DispatchEx('Word.Application')
    w.Visible = False
    w.DisplayAlerts = 0
    if run == 2:
        json.dump({'version': w.Version, 'build': w.Build}, open(os.path.join(out, '_word_version.json'), 'w'))
    try:
        for i, d in enumerate(ids, 1):
            t, h = d.split('/')
            dst = os.path.join(out, sub, f'{t}__{h}.pdf')
            if os.path.exists(dst):
                continue
            t0 = time.time()
            try:
                doc = w.Documents.Open(os.path.join(REPO, 'pipeline_data', 'docx_corpus', 'zh', t, h + '.docx'),
                                       ReadOnly=True, AddToRecentFiles=False)
                doc.ExportAsFixedFormat(dst, 17)
                doc.Close(False)
                msg = f'[run{run} {i:2d}/{len(ids)}] {t}__{h} OK {time.time() - t0:.1f}s'
            except Exception as e:
                msg = f'[run{run} {i:2d}/{len(ids)}] {t}__{h} FAIL {e}'
            print(msg); log.write(msg + '\n'); log.flush()
    finally:
        w.Quit()
open(os.path.join(out, 'cloudfonts_after.txt'), 'w', encoding='utf8').write('\n'.join(sorted(os.listdir(cloud))))
