# -*- coding: utf-8 -*-
"""Incremental, resumable SSIM A/B — same maths as tools/metrics/ssim_ab.py.

  python _ssim_ab_inc.py OXI_S1189 <baselist.txt> [start] [count]

Appends one line per base to C:/tmp/<FLAG>_ssim.log ("base  pages  net(B-A)"),
so a kill keeps every doc already measured. A = flag SET, B = flag UNSET
(exactly ssim_ab's convention: for an OPT-IN flag, A is the NEW behaviour and
net(B-A) must be read with the sign inverted).
"""
import os, subprocess, sys, tempfile
from pathlib import Path
_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO))
from pipeline.config import WORD_PNG_DIR, RENDER_DPI
from pipeline.ssim_calculator import _load_rgb, _resize_to_match
from skimage.metrics import structural_similarity as ssim
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DW = os.environ.get("OXI_DWRITE_EXE") or str(
    _REPO / "tools" / "oxi-dwrite-renderer" / "target" / "release" / "oxi-dwrite-renderer.exe")
DOCS = _REPO / "tools" / "golden-test" / "documents" / "docx"

FLAG = sys.argv[1]
BASELIST = sys.argv[2]
START = int(sys.argv[3]) if len(sys.argv) > 3 else 0
COUNT = int(sys.argv[4]) if len(sys.argv) > 4 else 10**9
LOG = f"C:/tmp/{FLAG}_ssim.log"

def find(base):
    e = DOCS / (base + ".docx")
    if e.exists():
        return str(e)
    c = sorted(p for p in DOCS.glob(base.split("_")[0] + "*.docx")
               if not p.name.startswith("~$"))
    return str(c[0]) if c else None

def render(docx, on, outdir):
    env = dict(os.environ)
    if on:
        env[FLAG] = "1"
    else:
        env.pop(FLAG, None)
    Path(outdir).mkdir(parents=True, exist_ok=True)
    subprocess.run([DW, docx, str(Path(outdir) / "p"), str(RENDER_DPI)],
                   capture_output=True, timeout=600, env=env)
    ps, i = [], 1
    while (Path(outdir) / f"p_p{i}.png").exists():
        ps.append(str(Path(outdir) / f"p_p{i}.png")); i += 1
    return ps

def ssim2(w, o):
    a = _load_rgb(w); b = _resize_to_match(_load_rgb(o), a)
    return ssim(a, b, full=False, channel_axis=2, data_range=255)

done = set()
if os.path.exists(LOG):
    for ln in open(LOG, encoding="utf-8"):
        done.add(ln.split("\t")[0])

bases = [b.strip() for b in open(BASELIST, encoding="utf-8") if b.strip()]
todo = [b for b in bases[START:START + COUNT] if b not in done]
with open(LOG, "a", encoding="utf-8") as log:
    for base in todo:
        d = find(base)
        if not d:
            log.write(f"{base}\t-\tNO_DOCX\n"); log.flush(); continue
        with tempfile.TemporaryDirectory(prefix="ssimab_") as tmp:
            pa = render(d, True, Path(tmp) / "A")
            pb = render(d, False, Path(tmp) / "B")
            same = (len(pa) == len(pb)) and all(
                open(x, "rb").read() == open(y, "rb").read() for x, y in zip(pa, pb))
            if same:
                log.write(f"{base}\t{len(pa)}\tIDENTICAL\n"); log.flush()
                print(f"{base}\tIDENTICAL"); continue
            wdir = Path(WORD_PNG_DIR) / base
            net, npg, i = 0.0, 0, 1
            while True:
                wp = wdir / f"page_{i:04d}.png"
                ap = Path(tmp) / "A" / f"p_p{i}.png"
                bp = Path(tmp) / "B" / f"p_p{i}.png"
                if not wp.exists() or not ap.exists() or not bp.exists():
                    break
                try:
                    net += ssim2(str(wp), str(bp)) - ssim2(str(wp), str(ap))
                    npg += 1
                except Exception:
                    pass
                i += 1
            if npg == 0:
                log.write(f"{base}\t{len(pa)}\tCHANGED_NO_REF\n")
            else:
                log.write(f"{base}\t{npg}\t{net:+.6f}\n")
            log.flush()
            print(f"{base}\tpages={npg}\tnet(B-A)={net:+.6f}")
