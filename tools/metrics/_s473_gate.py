"""S469 full-corpus drift-free gate for the textbox CJK exact/atLeast glyph
correction (OXI_S473_LOCOMP, default 1.9). OFF = env=0, ON = default.
Fresh DWrite renders both modes vs cached Word PNG (immune to baseline
staleness). Reports per-page + per-doc deltas + bottom-N movement."""
import os, sys, glob, subprocess
import numpy as np
from PIL import Image
from skimage.metrics import structural_similarity as ssim

REPO = r"C:\Users\ryuji\oxi-main"
DWRITE = os.path.join(REPO, "tools", "oxi-dwrite-renderer", "target", "release", "oxi-dwrite-renderer.exe")
DOCX = os.path.join(REPO, "tools", "golden-test", "documents", "docx")
WORD = os.path.join(REPO, "pipeline_data", "word_png")
TMP = r"C:\Users\ryuji\AppData\Local\Temp\s469gate"
os.makedirs(TMP, exist_ok=True)


def render(docx_path, prefix, on):
    env = dict(os.environ)
    if on:
        env["OXI_S473_LOCOMP"] = "1"  # ON
    else:
        env.pop("OXI_S473_LOCOMP", None)  # OFF=baseline
    subprocess.run([DWRITE, docx_path, prefix, "150"], stdout=subprocess.DEVNULL,
                   stderr=subprocess.DEVNULL, env=env)


def ssim_page(w, o):
    if not (os.path.exists(w) and os.path.exists(o)):
        return None
    a = np.array(Image.open(w).convert("L"))
    b = Image.open(o).convert("L")
    if b.size != (a.shape[1], a.shape[0]):
        b = b.resize((a.shape[1], a.shape[0]), Image.LANCZOS)
    return float(ssim(a, np.array(b), data_range=255))


def main():
    pat = sys.argv[1] if len(sys.argv) > 1 else ""
    docs = sorted(glob.glob(os.path.join(DOCX, "*%s*.docx" % pat)))
    rows = []; per_doc = {}
    prog = open(os.path.join(REPO, "tools", "metrics", "_s473_gate.progress"), "w", encoding="utf-8")
    for dp in docs:
        doc_id = os.path.splitext(os.path.basename(dp))[0]
        wdir = os.path.join(WORD, doc_id)
        if not os.path.isdir(wdir):
            continue
        won = os.path.join(TMP, "on_" + doc_id); woff = os.path.join(TMP, "off_" + doc_id)
        render(dp, won, True); render(dp, woff, False)
        ds = []
        for wp in sorted(glob.glob(os.path.join(wdir, "page_*.png"))):
            pg = int(os.path.basename(wp)[5:9])
            on = ssim_page(wp, "%s_p%d.png" % (won, pg))
            off = ssim_page(wp, "%s_p%d.png" % (woff, pg))
            if on is None or off is None:
                continue
            rows.append((doc_id, pg, off, on, on - off)); ds.append(on - off)
            per_doc.setdefault(doc_id, []).append(on - off)
        if ds and abs(sum(ds)) > 0.0005:
            prog.write("%-46s net=%+.4f (%dpg)\n" % (doc_id[:46], sum(ds), len(ds))); prog.flush()
    rows.sort(key=lambda r: r[4])
    out = ["%-44s %4s %8s %8s %8s" % ("doc", "pg", "OFF", "ON", "delta")]
    for doc, pg, off, on, dd in rows:
        if abs(dd) > 0.0005:
            out.append("%-44s %4d %8.4f %8.4f %+8.4f" % (doc[:44], pg, off, on, dd))
    if rows:
        net = sum(r[4] for r in rows)
        up = sum(1 for r in rows if r[4] > 0.0005); dn = sum(1 for r in rows if r[4] < -0.0005)
        out.append("-" * 78)
        out.append("PAGES=%d net=%+.4f mean=%+.6f improved=%d regressed=%d same=%d"
                   % (len(rows), net, net / len(rows), up, dn, len(rows) - up - dn))
        # bottom-N (by OFF ssim) movement
        by_off = sorted(rows, key=lambda r: r[2])
        for N in (3, 5, 10):
            offsum = sum(r[2] for r in by_off[:N]); onsum = sum(r[3] for r in by_off[:N])
            out.append("bottom-%d (by OFF): OFF=%.4f ON=%.4f delta=%+.4f" % (N, offsum, onsum, onsum - offsum))
        out.append("--- per-doc net (movers) ---")
        for doc, dl in sorted(per_doc.items(), key=lambda kv: sum(kv[1])):
            if abs(sum(dl)) > 0.001:
                out.append("%-46s net=%+.4f (%dpg)" % (doc[:46], sum(dl), len(dl)))
    txt = "\n".join(out)
    open(os.path.join(REPO, "tools", "metrics", "_s473_gate.out"), "w", encoding="utf-8").write(txt)
    print(txt.encode("ascii", "replace").decode())


if __name__ == "__main__":
    main()
