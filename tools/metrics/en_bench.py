# -*- coding: utf-8 -*-
"""EN discovery benchmark pipeline (docx-corpus, frozen rule — see
pipeline_data/en_benchmark/FROZEN_SELECTION_RULE.md).

Selects the first-5-per-type SHA-256-ascending docx-corpus/en docs that pass
Oxi-harness + Word-render, then renders Word (ground truth) / Oxi / LibreOffice
at 150 DPI and reports SSIM(Oxi,Word), SSIM(LO,Word), and the Oxi-vs-LO delta.

Kept SEPARATE from the frozen v0.8 benchmark (235 JP + 6 real_en).

Phases (each resumable; artifacts cached):
  python en_bench.py select    # freeze the 50-doc list -> _selected.json
  python en_bench.py word       # Word ground-truth PNGs
  python en_bench.py oxi        # Oxi PNGs
  python en_bench.py lo         # LibreOffice PNGs
  python en_bench.py ssim       # compute + report
  python en_bench.py all        # select..ssim (word/lo need COM/soffice)
"""
import os, sys, json, subprocess, glob
from pathlib import Path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.stdout.reconfigure(encoding="utf-8")

REPO = Path(r"c:\Users\ryuji\oxi-main")
CORPUS = REPO / "pipeline_data" / "docx_corpus" / "en"
BENCH = REPO / "pipeline_data" / "en_benchmark"
WORD_PNG = BENCH / "word_png"
OXI_PNG = BENCH / "oxi_png"
LO_PNG = BENCH / "lo_png"
HARNESS = REPO / "pipeline_data" / "docx_corpus" / "_harness.json"
SELECTED = BENCH / "_selected.json"
RESULT = BENCH / "_result.json"
DW = REPO / "tools" / "oxi-dwrite-renderer" / "target" / "release" / "oxi-dwrite-renderer.exe"
SOFFICE = Path(r"C:\Program Files\LibreOffice\program\soffice.exe")
DPI = 150
TYPES = ["legal", "forms", "reports", "policies", "educational",
         "correspondence", "technical", "administrative", "creative", "reference"]
N_PER_TYPE = 5


def doc_id(path):
    # <type>__<sha16>
    p = Path(path)
    return f"{p.parent.name}__{p.stem}"


def pdf_to_pngs(pdf, outdir, dpi=DPI):
    import fitz
    outdir = Path(outdir); outdir.mkdir(parents=True, exist_ok=True)
    d = fitz.open(pdf)
    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)
    n = 0
    for i in range(len(d)):
        pix = d[i].get_pixmap(matrix=mat, alpha=False)
        pix.save(str(outdir / f"page_{i+1:04d}.png"))
        n += 1
    d.close()
    return n


def select():
    # candidates per type = SHA-ascending (filename order) docs that pass Oxi harness
    harness = {r["doc"].replace("\\", "/"): r for r in json.load(open(HARNESS, encoding="utf-8"))}
    per_type = {t: [] for t in TYPES}
    for f in sorted(glob.glob(str(CORPUS / "*" / "*.docx"))):
        rel = os.path.relpath(f, CORPUS.parent).replace("\\", "/")  # en/<type>/<sha>.docx
        t = Path(f).parent.name
        h = harness.get(rel)
        ok = h and h.get("status") == "ok" and h.get("pages")
        per_type.setdefault(t, [])
        per_type[t].append({"path": f, "sha": Path(f).stem, "oxi_ok": bool(ok),
                             "oxi_pages": h.get("pages") if h else None})
    # candidate order = SHA ascending (already sorted by glob); filter oxi_ok
    sel = {}
    for t in TYPES:
        cands = [c for c in sorted(per_type.get(t, []), key=lambda x: x["sha"]) if c["oxi_ok"]]
        sel[t] = cands  # keep the full ordered candidate list; word phase takes first 5 that Word-render
    json.dump(sel, open(SELECTED, "w", encoding="utf-8"), indent=1)
    tot = sum(min(N_PER_TYPE, len(v)) for v in sel.values())
    print("candidates per type (oxi-ok):",
          {t: len(sel[t]) for t in TYPES})
    print(f"target selection = {tot} (first {N_PER_TYPE}/type that also Word-render)")


def word():
    import win32com.client
    sel = json.load(open(SELECTED, encoding="utf-8"))
    word_app = win32com.client.Dispatch("Word.Application")
    word_app.Visible = False
    word_app.DisplayAlerts = 0
    final = {}
    tmp_pdf = str(BENCH / "_tmp.pdf")
    for t in TYPES:
        kept = []
        for c in sel.get(t, []):
            if len(kept) >= N_PER_TYPE:
                break
            did = doc_id(c["path"])
            outdir = WORD_PNG / did
            if (outdir / "page_0001.png").exists():
                kept.append(c); continue
            try:
                if os.path.exists(tmp_pdf):
                    os.remove(tmp_pdf)
                d = word_app.Documents.Open(str(Path(c["path"]).resolve()),
                                            ReadOnly=True, AddToRecentFiles=False)
                d.ExportAsFixedFormat(tmp_pdf, 17)
                d.Close(False)
                npg = pdf_to_pngs(tmp_pdf, outdir)
                kept.append(c)
                print(f"  {did}: word {npg}pg")
            except Exception as e:
                print(f"  {did}: WORD FAIL {str(e)[:60]} -> skip (take next)")
                try: d.Close(False)
                except Exception: pass
        final[t] = kept
    word_app.Quit()
    json.dump(final, open(BENCH / "_final.json", "w", encoding="utf-8"), indent=1)
    tot = sum(len(v) for v in final.values())
    print(f"FINAL selection Word-rendered: {tot} docs " +
          str({t: len(final[t]) for t in TYPES}))


def oxi():
    final = json.load(open(BENCH / "_final.json", encoding="utf-8"))
    n = 0
    for t in TYPES:
        for c in final[t]:
            did = doc_id(c["path"])
            outdir = OXI_PNG / did
            if (outdir / "p_p1.png").exists():
                n += 1; continue
            outdir.mkdir(parents=True, exist_ok=True)
            subprocess.run([str(DW), str(Path(c["path"]).resolve()),
                            str(outdir / "p"), str(DPI)], capture_output=True, timeout=300)
            n += 1
    print(f"oxi rendered {n} docs")


def lo():
    final = json.load(open(BENCH / "_final.json", encoding="utf-8"))
    tmp = BENCH / "_lo_tmp"
    tmp.mkdir(parents=True, exist_ok=True)
    n = 0
    for t in TYPES:
        for c in final[t]:
            did = doc_id(c["path"])
            outdir = LO_PNG / did
            if (outdir / "page_0001.png").exists():
                n += 1; continue
            for old in tmp.glob("*.pdf"):
                old.unlink()
            src = Path(c["path"]).resolve()
            r = subprocess.run([str(SOFFICE), "--headless", "--convert-to", "pdf",
                                "--outdir", str(tmp), str(src)],
                               capture_output=True, text=True, timeout=180)
            pdf = tmp / (src.stem + ".pdf")
            if pdf.exists():
                pdf_to_pngs(str(pdf), outdir)
                n += 1
            else:
                print(f"  {did}: LO FAIL {(r.stderr or '')[:60]}")
    print(f"lo rendered {n} docs")


def _ssim(a_png, b_png):
    from pipeline.ssim_calculator import _load_rgb, _resize_to_match
    from skimage.metrics import structural_similarity as ssim
    w = _load_rgb(str(a_png)); o = _resize_to_match(_load_rgb(str(b_png)), w)
    return ssim(w, o, full=False, channel_axis=2, data_range=255)


def do_ssim():
    sys.path.insert(0, str(REPO))
    final = json.load(open(BENCH / "_final.json", encoding="utf-8"))
    rows = []
    for t in TYPES:
        for c in final[t]:
            did = doc_id(c["path"])
            wdir = WORD_PNG / did
            i = 1; ox = []; lo_ = []
            while True:
                wp = wdir / f"page_{i:04d}.png"
                if not wp.exists(): break
                op = OXI_PNG / did / f"p_p{i}.png"
                lp = LO_PNG / did / f"page_{i:04d}.png"
                if op.exists():
                    try: ox.append(_ssim(wp, op))
                    except Exception: pass
                if lp.exists():
                    try: lo_.append(_ssim(wp, lp))
                    except Exception: pass
                i += 1
            oxm = sum(ox)/len(ox) if ox else None
            lom = sum(lo_)/len(lo_) if lo_ else None
            rows.append({"doc": did, "type": t, "pages": i-1,
                         "oxi": oxm, "lo": lom,
                         "delta": (oxm-lom) if (oxm is not None and lom is not None) else None})
    json.dump(rows, open(RESULT, "w", encoding="utf-8"), indent=1)
    ox_docs = [r["oxi"] for r in rows if r["oxi"] is not None]
    lo_docs = [r["lo"] for r in rows if r["lo"] is not None]
    both = [r for r in rows if r["delta"] is not None]
    print(f"\n=== EN discovery benchmark ({len(rows)} docs) ===")
    if ox_docs:
        print(f"Oxi  per-doc: mean={sum(ox_docs)/len(ox_docs):.4f}  floor={min(ox_docs):.4f}")
    if lo_docs:
        print(f"LO   per-doc: mean={sum(lo_docs)/len(lo_docs):.4f}  floor={min(lo_docs):.4f}")
    if both:
        wins = sum(1 for r in both if r["delta"] > 0.0005)
        ties = sum(1 for r in both if abs(r["delta"]) <= 0.0005)
        print(f"Oxi vs LO (both rendered, {len(both)} docs): Oxi wins {wins}, ties {ties}, LO wins {len(both)-wins-ties}")
        print(f"  mean delta (Oxi-LO) = {sum(r['delta'] for r in both)/len(both):+.4f}")
    print("\nbottom-8 Oxi docs:")
    for r in sorted(rows, key=lambda x: x["oxi"] if x["oxi"] is not None else 9)[:8]:
        d = f"{r['delta']:+.3f}" if r["delta"] is not None else "  -  "
        print(f"  {r['oxi'] if r['oxi'] else 0:.4f} (LO {r['lo'] if r['lo'] else 0:.4f}, Δ{d})  {r['doc']}")


if __name__ == "__main__":
    mode = sys.argv[1] if len(sys.argv) > 1 else "all"
    if mode in ("select", "all"): select()
    if mode in ("word", "all"): word()
    if mode in ("oxi", "all"): oxi()
    if mode in ("lo", "all"): lo()
    if mode in ("ssim", "all"): do_ssim()
