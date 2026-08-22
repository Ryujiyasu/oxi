# -*- coding: utf-8 -*-
"""Incremental, resumable SSIM A/B for a COMBINATION of flags.

  python _ssim_ab_envs.py OXI_S1192,OXI_S1195 <baselist.txt> [start] [count]

Same maths and same log format as `_ssim_ab_inc.py` (which takes ONE flag).
A = flags SET, B = flags UNSET, so for OPT-IN flags A is the NEW behaviour and
`net(B-A)` reads with the sign inverted: an improvement prints NEGATIVE.
A *_DISABLE flag is inverted (set in the B arm), matching `_ab_envs.py`.
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

FLAGS = [f for f in sys.argv[1].split(",") if f]
BASELIST = sys.argv[2]
START = int(sys.argv[3]) if len(sys.argv) > 3 else 0
COUNT = int(sys.argv[4]) if len(sys.argv) > 4 else 10 ** 9
LOG = "C:/tmp/%s_ssim.log" % "_".join(FLAGS)


def find(base):
    e = DOCS / (base + ".docx")
    if e.exists():
        return str(e)
    c = sorted(p for p in DOCS.glob(base.split("_")[0] + "*.docx")
               if not p.name.startswith("~$"))
    return str(c[0]) if c else None


def render(docx, on, outdir):
    env = dict(os.environ)
    for flag in FLAGS:
        name, _, value = flag.partition("=")
        invert = name.endswith("_DISABLE")
        env.pop(name, None)
        if on != invert:
            env[name] = value or "1"
    Path(outdir).mkdir(parents=True, exist_ok=True)
    subprocess.run([DW, docx, str(Path(outdir) / "p"), str(RENDER_DPI)],
                   capture_output=True, timeout=600, env=env)
    ps, i = [], 1
    while (Path(outdir) / ("p_p%d.png" % i)).exists():
        ps.append(str(Path(outdir) / ("p_p%d.png" % i))); i += 1
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
            log.write("%s\t-\tNO_DOCX\n" % base); log.flush(); continue
        with tempfile.TemporaryDirectory(prefix="ssimab_") as tmp:
            pa = render(d, True, Path(tmp) / "A")
            pb = render(d, False, Path(tmp) / "B")
            same = (len(pa) == len(pb)) and all(
                open(x, "rb").read() == open(y, "rb").read() for x, y in zip(pa, pb))
            if same:
                log.write("%s\t%d\tIDENTICAL\n" % (base, len(pa))); log.flush()
                print("%s\tIDENTICAL" % base, flush=True); continue
            wdir = Path(WORD_PNG_DIR) / base
            net, npg, i = 0.0, 0, 1
            while True:
                wp = wdir / ("page_%04d.png" % i)
                ap = Path(tmp) / "A" / ("p_p%d.png" % i)
                bp = Path(tmp) / "B" / ("p_p%d.png" % i)
                if not wp.exists() or not ap.exists() or not bp.exists():
                    break
                try:
                    net += ssim2(str(wp), str(bp)) - ssim2(str(wp), str(ap))
                    npg += 1
                except Exception:  # noqa: BLE001
                    pass
                i += 1
            if npg == 0:
                log.write("%s\t%d\tCHANGED_NO_REF\n" % (base, len(pa)))
            else:
                log.write("%s\t%d\t%+.6f\n" % (base, npg, net))
            log.flush()
            print("%s\tpages=%d\tnet(B-A)=%+.6f" % (base, npg, net), flush=True)
