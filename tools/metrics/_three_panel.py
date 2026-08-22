# -*- coding: utf-8 -*-
"""Word | Oxi-A | Oxi-B three-panel for one page of one doc.

  python _three_panel.py <base> <page> <out.png> [FLAGS_B] [FLAGS_A] [y0 y1]

FLAGS are comma-separated env names (A defaults to none = shipped default).
Renders both arms with the DWrite renderer, crops [y0..y1] in points when
given, and writes the labelled strip.
"""
import os, subprocess, sys, tempfile
from pathlib import Path
from PIL import Image, ImageDraw

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO))
from pipeline.config import WORD_PNG_DIR, RENDER_DPI  # noqa: E402

DW = REPO / "tools" / "oxi-dwrite-renderer" / "target" / "release" / "oxi-dwrite-renderer.exe"
DOCS = REPO / "tools" / "golden-test" / "documents" / "docx"

base = sys.argv[1]
page = int(sys.argv[2])
out = sys.argv[3]
flags_b = [f for f in (sys.argv[4] if len(sys.argv) > 4 else "").split(",") if f]
flags_a = [f for f in (sys.argv[5] if len(sys.argv) > 5 else "").split(",") if f]
y0 = float(sys.argv[6]) if len(sys.argv) > 6 else None
y1 = float(sys.argv[7]) if len(sys.argv) > 7 else None

docx = DOCS / (base + ".docx")
if not docx.exists():
    docx = sorted(DOCS.glob(base.split("_")[0] + "*.docx"))[0]


def render(flags, outdir):
    env = dict(os.environ)
    for f in ("OXI_S1192", "OXI_S1195", "OXI_S1192G"):
        env.pop(f, None)
    for f in flags:
        env[f] = "1"
    Path(outdir).mkdir(parents=True, exist_ok=True)
    subprocess.run([str(DW), str(docx), str(Path(outdir) / "p"), str(RENDER_DPI)],
                   capture_output=True, timeout=900, env=env)
    return Path(outdir) / ("p_p%d.png" % page)


def crop(im):
    if y0 is None:
        return im
    s = im.height / 842.0 if im.height > 1000 else 1.0
    return im.crop((0, int(y0 * s), im.width, int(y1 * s)))


with tempfile.TemporaryDirectory(prefix="tp_") as tmp:
    pa = render(flags_a, Path(tmp) / "A")
    pb = render(flags_b, Path(tmp) / "B")
    wp = Path(WORD_PNG_DIR) / base / ("page_%04d.png" % page)
    ims = [crop(Image.open(p).convert("RGB")) for p in (wp, pa, pb)]
    h = max(i.height for i in ims)
    ims = [i.resize((int(i.width * h / i.height), h), Image.LANCZOS) for i in ims]
    pad, top = 8, 26
    W = sum(i.width for i in ims) + pad * (len(ims) + 1)
    canvas = Image.new("RGB", (W, h + top + pad), "white")
    d = ImageDraw.Draw(canvas)
    labels = ["Word", "Oxi  " + ("+".join(flags_a) if flags_a else "default"),
              "Oxi  " + ("+".join(flags_b) if flags_b else "default")]
    x = pad
    for im, lab in zip(ims, labels):
        canvas.paste(im, (x, top))
        d.text((x + 4, 6), lab, fill="black")
        x += im.width + pad
    canvas.save(out)
    print("wrote", out, canvas.size)
