"""S468 decisive experiment: drift-free OFF-vs-ON SSIM gate for VSNAP ALONE
(OXI_S467_VSNAP only, NOT combined with PBDR). Full corpus.

S467 gated VSNAP+PBDR COMBINED -> net-neutral, attributed CJK regressions to
"raw height mismatch". S468 proved raw heights MATCH (grid quantization is the
only diff) and that VSNAP only fires on is_multiple_spacing paragraphs. This
isolates VSNAP's TRUE per-doc effect, drift-free (fresh OFF and ON renders
each compared to the cached Word PNG -> immune to committed-baseline staleness).

Usage: python _s468_vsnap_gate.py [glob-substr]   (default = all docs w/ Word PNG)
"""
import os, sys, glob, subprocess
import numpy as np
from PIL import Image
from skimage.metrics import structural_similarity as ssim

REPO = r"C:\Users\ryuji\oxi-main"
DWRITE = os.path.join(REPO, "tools", "oxi-dwrite-renderer", "target", "release", "oxi-dwrite-renderer.exe")
DOCX = os.path.join(REPO, "tools", "golden-test", "documents", "docx")
WORD = os.path.join(REPO, "pipeline_data", "word_png")
TMP = r"C:\Users\ryuji\AppData\Local\Temp\s468vgate"
os.makedirs(TMP, exist_ok=True)


def render(docx_path, prefix, on):
    env = dict(os.environ)
    env.pop("OXI_S467_VSNAP", None)
    env.pop("OXI_S467_PBDR_ENABLE", None)
    if on:
        env["OXI_S467_VSNAP"] = "1"   # VSNAP ONLY (no PBDR)
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
    pat = sys.argv[1] if len(sys.argv) > 1 else ""
    docs = sorted(glob.glob(os.path.join(DOCX, "*%s*.docx" % pat)))
    rows = []
    per_doc = {}
    prog = os.path.join(REPO, "tools", "metrics", "_s468_vsnap_gate.progress")
    pf = open(prog, "w", encoding="utf-8")
    ndoc = 0
    for dp in docs:
        doc_id = os.path.splitext(os.path.basename(dp))[0]
        wdir = os.path.join(WORD, doc_id)
        if not os.path.isdir(wdir):
            continue
        won = os.path.join(TMP, "on_" + doc_id)
        woff = os.path.join(TMP, "off_" + doc_id)
        render(dp, won, on=True)
        render(dp, woff, on=False)
        ds = []
        for wpng in sorted(glob.glob(os.path.join(wdir, "page_*.png"))):
            pg = int(os.path.basename(wpng)[5:9])
            on = ssim_page(wpng, "%s_p%d.png" % (won, pg))
            off = ssim_page(wpng, "%s_p%d.png" % (woff, pg))
            if on is None or off is None:
                continue
            rows.append((doc_id, pg, off, on, on - off))
            per_doc.setdefault(doc_id, []).append(on - off)
            ds.append(on - off)
        ndoc += 1
        if ds:
            pf.write("%-46s net=%+.4f (%d pg)  [running net=%+.4f over %d docs]\n"
                     % (doc_id[:46], sum(ds), len(ds), sum(r[4] for r in rows), ndoc))
            pf.flush()
    rows.sort(key=lambda r: r[4])
    out = []
    out.append("%-44s %4s %8s %8s %8s" % ("doc", "pg", "OFF", "ON", "delta"))
    for doc, pg, off, on, d in rows:
        if abs(d) > 0.0005:
            out.append("%-44s %4d %8.4f %8.4f %+8.4f" % (doc[:44], pg, off, on, d))
    if rows:
        net = sum(r[4] for r in rows)
        up = sum(1 for r in rows if r[4] > 0.0005)
        dn = sum(1 for r in rows if r[4] < -0.0005)
        out.append("-" * 78)
        out.append("PAGES=%d  net=%+.4f  mean=%+.6f  improved=%d regressed=%d same=%d"
                   % (len(rows), net, net / len(rows), up, dn, len(rows) - up - dn))
        # per-doc net
        out.append("--- per-doc net (sorted) ---")
        dn_sorted = sorted(per_doc.items(), key=lambda kv: sum(kv[1]))
        for doc, ds in dn_sorted:
            s = sum(ds)
            if abs(s) > 0.001:
                out.append("%-46s net=%+.4f  (%d pg)" % (doc[:46], s, len(ds)))
    txt = "\n".join(out)
    open(os.path.join(REPO, "tools", "metrics", "_s468_vsnap_gate.out"), "w", encoding="utf-8").write(txt)
    print(txt.encode("ascii", "replace").decode())


if __name__ == "__main__":
    main()
