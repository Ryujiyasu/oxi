"""Verify jc=both trigger across multiple docs.

Expected:
- Docs with jc=both on Normal: yakumono compressed
- 0e7a (jc=unspecified): yakumono NOT compressed

Pick first paragraph containing ・ in each, measure advance.
"""
import os, sys, time, subprocess
import win32com.client
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOCS = [
    ("d77a", "tools/golden-test/documents/docx/d77a58485f16_20240705_resources_data_outline_08.docx", "both"),
    ("e3c545", "tools/golden-test/documents/docx/e3c545fac7a7_LOD_Handbook.docx", "both"),
    ("b35", "tools/golden-test/documents/docx/b35123fe8efc_tokumei_08_01.docx", "both"),
    ("b837", "tools/golden-test/documents/docx/b837808d0555_20240705_resources_data_guideline_02.docx", "both"),
    ("0e7a", "tools/golden-test/documents/docx/0e7af1ae8f21_20230331_resources_open_data_contract_sample_00.docx", "unspec"),
]


def measure_one(stem, path, jc):
    try:
        subprocess.run(["taskkill", "/F", "/IM", "WINWORD.EXE"], capture_output=True, timeout=5)
    except Exception: pass
    time.sleep(0.5)
    word = win32com.client.DispatchEx("Word.Application")
    try:
        try: word.Visible = False
        except: pass
        try: word.DisplayAlerts = False
        except: pass
        doc = word.Documents.Open(os.path.abspath(path), ReadOnly=True); time.sleep(0.3)
        doc.Repaginate()
        # Find first para with ・
        nparas = doc.Paragraphs.Count
        for pi in range(1, min(nparas + 1, 300)):
            try:
                p = doc.Paragraphs(pi)
                txt = p.Range.Text.replace("\r","").replace("\x07","")
                if "・" in txt[:5]:
                    pr = p.Range
                    # Find ・ position
                    for ci in range(1, min(pr.Characters.Count + 1, 10)):
                        ch = pr.Characters(ci)
                        if ch.Text == "・":
                            x1 = ch.Information(5)
                            y1 = ch.Information(6)
                            if ci + 1 <= pr.Characters.Count:
                                ch2 = pr.Characters(ci + 1)
                                x2 = ch2.Information(5)
                                y2 = ch2.Information(6)
                                if y1 == y2:
                                    fs = ch.Font.Size or 10.5
                                    adv = x2 - x1
                                    print(f"[{stem:>6} jc={jc:>6}] para_idx={pi} fs={fs} ・adv={adv:.2f}pt (fontSize={fs}, ratio={adv/fs:.2%})")
                                    return adv
                            break
                    break
            except Exception:
                continue
        doc.Close(False)
    except Exception as e:
        print(f"[{stem}] err: {e}")
    finally:
        try: word.Quit()
        except: pass


for stem, path, jc in DOCS:
    measure_one(stem, path, jc)
