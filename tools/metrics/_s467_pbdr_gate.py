"""S467 drift-free OFF-vs-ON SSIM gate for the pBdr full-border fix.
Renders each affected doc with DWrite twice (ON=default, OFF=OXI_S467_PBDR_DISABLE=1)
to temp dirs, SSIM each page vs the cached Word PNG, and reports ON-vs-OFF delta.
Avoids the ratchet-only committed-baseline drift (S466 lesson)."""
import os, io, sys, glob, subprocess, json
import numpy as np
from PIL import Image
from skimage.metrics import structural_similarity as ssim

REPO = r"C:\Users\ryuji\oxi-main"
DWRITE = os.path.join(REPO, "tools", "oxi-dwrite-renderer", "target", "release", "oxi-dwrite-renderer.exe")
DOCX = os.path.join(REPO, "tools", "golden-test", "documents", "docx")
WORD = os.path.join(REPO, "pipeline_data", "word_png")
TMP = r"C:\Users\ryuji\AppData\Local\Temp\s467gate"
os.makedirs(TMP, exist_ok=True)


def render(docx_path, prefix, on):
    # on=False: clean baseline (no env). on=True: combined S467 fix
    # (VSNAP visual cumulative 0.75 grid-snap + PBDR full-border).
    env = dict(os.environ)
    env.pop("OXI_S467_VSNAP", None)
    env.pop("OXI_S467_PBDR_ENABLE", None)
    if on:
        env["OXI_S467_VSNAP"] = "1"
        env["OXI_S467_PBDR_ENABLE"] = "1"
    subprocess.run([DWRITE, docx_path, prefix, "150"], stdout=subprocess.DEVNULL,
                   stderr=subprocess.DEVNULL, env=env)


def ssim_page(word_png, oxi_png):
    if not (os.path.exists(word_png) and os.path.exists(oxi_png)):
        return None
    a = np.array(Image.open(word_png).convert("L"))
    b = Image.open(oxi_png).convert("L")
    if b.size != (a.shape[1], a.shape[0]):
        b = b.resize((a.shape[1], a.shape[0]), Image.LANCZOS)
    return float(ssim(a, np.array(b), data_range=255))


def main():
    pat = sys.argv[1] if len(sys.argv) > 1 else "gen2"
    docs = sorted(glob.glob(os.path.join(DOCX, "*%s*.docx" % pat)))
    rows = []
    for dp in docs:
        doc_id = os.path.splitext(os.path.basename(dp))[0]
        wdir = os.path.join(WORD, doc_id)
        if not os.path.isdir(wdir):
            continue
        won = os.path.join(TMP, "on_" + doc_id)
        woff = os.path.join(TMP, "off_" + doc_id)
        render(dp, won, on=True)
        render(dp, woff, on=False)
        for wpng in sorted(glob.glob(os.path.join(wdir, "page_*.png"))):
            pg = int(os.path.basename(wpng)[5:9])
            on = ssim_page(wpng, "%s_p%d.png" % (won, pg))
            off = ssim_page(wpng, "%s_p%d.png" % (woff, pg))
            if on is None or off is None:
                continue
            rows.append((doc_id, pg, off, on, on - off))
    rows.sort(key=lambda r: r[4])
    nchg = [r for r in rows if abs(r[4]) > 0.0005]
    print("%-44s %4s %8s %8s %8s" % ("doc", "pg", "OFF", "ON", "delta"))
    for doc, pg, off, on, d in rows:
        if abs(d) > 0.0005:
            print("%-44s %4d %8.4f %8.4f %+8.4f" % (doc[:44], pg, off, on, d))
    if rows:
        net = sum(r[4] for r in rows)
        up = sum(1 for r in rows if r[4] > 0.0005)
        dn = sum(1 for r in rows if r[4] < -0.0005)
        print("-" * 78)
        print("pages=%d  net delta=%+.4f  mean delta=%+.5f  improved=%d regressed=%d same=%d"
              % (len(rows), net, net / len(rows), up, dn, len(rows) - up - dn))


if __name__ == "__main__":
    main()
