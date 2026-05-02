"""Verify master's V4 finding: `あ、」い` MS Mincho 12pt — master measured
both 、 and 」 at 12pt (NO compression). Our 4-font sweep showed 5.5pt
compression for B→B adjacency. Need to reproduce and resolve.
"""
import win32com.client
import os
import time
import zipfile
import re
import shutil
import sys
import tempfile

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath("pipeline_data/v4_repro")


def patch_setting(path, val):
    tmp = tempfile.mkdtemp(prefix="patch_")
    try:
        with zipfile.ZipFile(path) as z:
            z.extractall(tmp)
        sp = os.path.join(tmp, "word", "settings.xml")
        with open(sp, encoding="utf-8") as f:
            s = f.read()
        ns = re.sub(
            r'<w:characterSpacingControl[^/]*/>',
            f'<w:characterSpacingControl w:val="{val}"/>', s)
        if ns == s:
            ns = s.replace(
                "</w:settings>",
                f'<w:characterSpacingControl w:val="{val}"/></w:settings>')
        with open(sp, "w", encoding="utf-8") as f:
            f.write(ns)
        os.remove(path)
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
            for r, _, fs in os.walk(tmp):
                for fn in fs:
                    full = os.path.join(r, fn)
                    arc = os.path.relpath(full, tmp).replace("\\", "/")
                    z.write(full, arc)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


PROBES = [
    # Master's V1-V7 minimal repros
    ("V1_a_pun_i",       "あ、い"),         # CJK-、-CJK
    ("V3_a_pun_b_latin", "a、b"),           # Latin-、-Latin
    ("V4_a_pun_close_i", "あ、」い"),       # CJK-、-」-CJK pair
    ("V5_a_pun_pun_i",   "あ、、い"),       # CJK-、-、-CJK double
    ("V6_a_pun",         "あ、"),           # CJK-、 line-end
    # Our 4-font probe equivalents at MS Mincho 12pt
    ("ours_close_punct2", "漢」、漢"),
    ("ours_close_open",   "漢」（漢"),
    ("ours_punct_open",   "漢、（漢"),
]

FONT = "ＭＳ 明朝"
SIZE = 12.0


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    try:
        for setting in ["doNotCompress", "compressPunctuation"]:
            print(f"\n=== Setting: {setting} ===")
            for label, text in PROBES:
                fname = f"{label}_{setting}.docx"
                path = os.path.join(OUT_DIR, fname)
                try:
                    d = word.Documents.Add()
                    time.sleep(0.2)
                    rng = d.Range()
                    rng.InsertAfter(text)
                    rng = d.Range()
                    rng.Font.Name = FONT
                    rng.Font.Size = SIZE
                    d.Paragraphs(1).Alignment = 0
                    d.SaveAs2(path, FileFormat=12)
                    d.Close(SaveChanges=False)
                    patch_setting(path, setting)
                    d = word.Documents.Open(path, ReadOnly=True)
                    time.sleep(0.2)
                    chars = d.Range().Characters
                    xs = []
                    for ci in range(1, chars.Count + 1):
                        try:
                            c = chars(ci)
                            t = c.Text
                            if t in ("\r", "\x07"):
                                continue
                            xs.append((t, float(c.Information(5))))
                        except Exception:
                            continue
                    d.Close(SaveChanges=False)
                    advs = [(xs[i][0], round(xs[i + 1][1] - xs[i][1], 4))
                            for i in range(len(xs) - 1)]
                    print(f"  [{label:22s}] {text}: {advs}", flush=True)
                except Exception as e:
                    print(f"  [{label}] ERR: {e}", flush=True)
    finally:
        try:
            word.Quit()
        except Exception:
            pass


if __name__ == "__main__":
    main()
