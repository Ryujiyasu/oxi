# -*- coding: utf-8 -*-
"""Incremental Oxi-vs-LibreOffice bug-finder: appends per-page results to a
JSONL file as it goes, so partial output survives a timeout. Renders Oxi at
110 DPI (faster; SSIM ranking is robust to DPI). Skips docs whose result is
already in the JSONL (resumable)."""
import os, subprocess, tempfile, sys, json
from pathlib import Path
sys.path.insert(0, r"c:\Users\ryuji\oxi-main")
from pipeline.config import WORD_PNG_DIR
from pipeline.ssim_calculator import _load_rgb, _resize_to_match
from skimage.metrics import structural_similarity as ssim
sys.stdout.reconfigure(encoding="utf-8")
REPO = r"c:\Users\ryuji\oxi-main"
DW = os.path.join(REPO, "tools", "oxi-dwrite-renderer", "target", "release", "oxi-dwrite-renderer.exe")
DOCS = os.path.join(REPO, "tools", "golden-test", "documents", "docx")
LIBRA = os.path.join(REPO, "pipeline_data", "libra_png")
DPI = 110
OUT = sys.argv[1] if len(sys.argv) > 1 else r"c:\tmp\bugfind.jsonl"

def ssim2(a, b):
    w = _load_rgb(a); o = _resize_to_match(_load_rgb(b), w)
    return ssim(w, o, full=False, channel_axis=2, data_range=255)

def find_docx(base):
    e = Path(DOCS) / (base + ".docx")
    if e.exists(): return str(e)
    c = sorted(p for p in Path(DOCS).glob(base.split("_")[0] + "*.docx") if not p.name.startswith("~$"))
    return str(c[0]) if c else None

done = set()
if os.path.exists(OUT):
    for ln in open(OUT, encoding="utf-8"):
        try: done.add(json.loads(ln)["doc"])
        except Exception: pass

bases = sorted(d.name for d in Path(WORD_PNG_DIR).iterdir() if d.is_dir()
               and (Path(LIBRA) / d.name).is_dir())
fout = open(OUT, "a", encoding="utf-8")
with tempfile.TemporaryDirectory() as tmp:
    for i, base in enumerate(bases):
        if base in done:
            continue
        docx = find_docx(base)
        if not docx:
            continue
        od = Path(tmp) / base
        od.mkdir(parents=True, exist_ok=True)
        try:
            subprocess.run([DW, docx, str(od / "p"), str(DPI)], capture_output=True, timeout=180)
        except Exception:
            continue
        wdir = Path(WORD_PNG_DIR) / base; ldir = Path(LIBRA) / base
        p = 1
        while True:
            wp = wdir / f"page_{p:04d}.png"; lp = ldir / f"page_{p:04d}.png"; op = od / f"p_p{p}.png"
            if not wp.exists(): break
            if op.exists() and lp.exists():
                try:
                    so = ssim2(str(wp), str(op)); sl = ssim2(str(wp), str(lp))
                    fout.write(json.dumps({"doc": base, "pg": p, "oxi": round(so,4),
                                           "libra": round(sl,4), "delta": round(sl-so,4)}) + "\n")
                    fout.flush()
                except Exception:
                    pass
            p += 1
        sys.stderr.write(f"  {i}/{len(bases)} {base}\n"); sys.stderr.flush()
print("DONE")
