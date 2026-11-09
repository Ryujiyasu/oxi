"""Re-measure the gate's Word pagination truth for the documents that ask for faces this
machine gained (2026-11-09: the Windows Chinese supplemental fonts -- SimHei 黑体, KaiTi 楷体,
FangSong 仿宋, DFKai-SB 標楷體, DengXian Light, MingLiU 細明體 now system-wide). Their old
truth was Word's substitution; ja/policies__0d1dcfaf (標楷體) went 1.0 -> 0.31 in every build.

Same measurement and staging as refresh_gate_truth.py (one Word process per document, no
concurrent COM), restricted to the affected documents.

    python tools/metrics/refresh_truth_for_faces.py list
    OXI_TRUTH_STAGE=E:/oxi-scratch/truth_refresh_20261109 python tools/metrics/refresh_truth_for_faces.py measure
    OXI_TRUTH_STAGE=E:/oxi-scratch/truth_refresh_20261109 python tools/metrics/refresh_gate_truth.py report|install
"""
import os, re, sys, zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
import refresh_gate_truth as R

FACES = {'SimHei', '黑体', 'KaiTi', '楷体', 'FangSong', '仿宋', 'DFKai-SB', '標楷體', '标楷体',
         'DengXian Light', '等线 Light', 'MingLiU', '細明體', '细明体', 'MingLiU_HKSCS'}


def uses_new_face(docx):
    try:
        z = zipfile.ZipFile(docx)
        s = ''.join(z.read(n).decode('utf8', 'replace') for n in z.namelist()
                    if n in ('word/document.xml', 'word/styles.xml')
                    or re.match(r'word/(header|footer)\d*\.xml', n))
    except Exception:
        return False
    names = set(re.findall(r'w:(?:eastAsia|ascii|hAnsi)="([^"]+)"', s))
    # theme script lists (the default Office theme names 等线 Light for Hans) are not counted:
    # they reach a run only through a Chinese-language theme font reference
    return bool(names & FACES)


_all = R.work_list
R.work_list = lambda: [(t, d) for t, d in _all() if uses_new_face(d)]

if __name__ == '__main__':
    if sys.argv[1] == 'list':
        for t, d in R.work_list():
            print(os.path.basename(t), d)
        print(len(R.work_list()))
    elif sys.argv[1] == 'measure':
        R.measure()
