# -*- coding: utf-8 -*-
"""Line census for ANOTHER engine's PDF against the same Word truth.

line_census.py scores Oxi (glyph dump) against Word's PDF on page 1.  This scores
any engine that can turn a .docx into a PDF -- ibisOffice (`--export pdf`),
LibreOffice (`soffice --headless --convert-to pdf`), CraftAI's Office -- with the
same Word truth, the same line extraction (line_census.word_lines) and the same
pairing (line_census.pair), and writes the same summary.json shape, so
scratchpad census_cmp.py / line_census report compare engines directly.

    CENSUS_BLIND=blindG_20260929 python tools/metrics/engine_census.py convert ibis <tag>
    CENSUS_BLIND=blindG_20260929 python tools/metrics/engine_census.py score <tag>

Converted PDFs go to pipeline_data/engine_census/<tag>/pdf/<key>.pdf; the score
goes to pipeline_data/line_census/<tag>/summary.json.  Page counts (engine vs Word)
are added per document as "pages".
"""
import json, os, shutil, subprocess, sys, tempfile, time
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.stdout.reconfigure(encoding="utf-8")
import line_census as L  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "engine_census"
IBIS = os.environ.get("IBIS_EXE", r"C:\Users\ryuji\tools\ibisoffice-0.8.1\ibisOffice\ibisOffice.exe")
SOFFICE = os.environ.get("SOFFICE_EXE", r"C:\Program Files\LibreOffice\program\soffice.exe")


def convert_one(engine, docx, pdf, timeout=300):
    t = tempfile.mkdtemp()
    src = os.path.join(t, Path(docx).name); shutil.copy(docx, src)
    if engine == "ibis":
        # syntax checked against the binary's --help before the first run (see the archive note)
        cmd = [IBIS, "--export", "pdf", src, "--output", pdf]
    elif engine == "soffice":
        cmd = [SOFFICE, "--headless", "--convert-to", "pdf", "--outdir", t, src]
    else:
        raise SystemExit(f"unknown engine {engine}")
    t0 = time.time()
    r = subprocess.run(cmd, capture_output=True, timeout=timeout)
    if engine == "soffice":
        out = os.path.join(t, Path(docx).stem + ".pdf")
        if os.path.exists(out):
            shutil.move(out, pdf)
    return r.returncode, time.time() - t0, (r.stderr or b"")[-300:].decode("utf-8", "replace")


def convert(engine, tag):
    od = OUT / tag / "pdf"; od.mkdir(parents=True, exist_ok=True)
    log = {}
    for key, p in L.docs():
        pdf = od / f"{key}.pdf"
        if pdf.exists():
            continue
        try:
            rc, dt, err = convert_one(engine, str(p), str(pdf))
        except subprocess.TimeoutExpired:
            rc, dt, err = -1, 300.0, "timeout"
        log[key] = {"rc": rc, "sec": round(dt, 1), "ok": pdf.exists(), "err": err if not pdf.exists() else ""}
        print(f"  {key:44s} rc {rc} {dt:6.1f}s {'ok' if pdf.exists() else 'FAILED ' + err[:80]}", flush=True)
    prev = OUT / tag / "convert_log.json"
    old = json.loads(prev.read_text(encoding="utf-8")) if prev.exists() else {}
    old.update(log)
    prev.write_text(json.dumps(old, ensure_ascii=False, indent=1), encoding="utf-8")


def pages(pdf):
    import fitz
    return len(fitz.open(str(pdf)))


def score(tag):
    (L.OUT / tag).mkdir(parents=True, exist_ok=True)
    summary = {}
    for key, p in L.docs():
        wpdf = L.OUT / "pdf" / f"{key}.pdf"; epdf = OUT / tag / "pdf" / f"{key}.pdf"
        if not wpdf.exists() or not epdf.exists():
            summary[key] = {"error": "no word pdf" if not wpdf.exists() else "no engine pdf"}; continue
        try:
            W = L.word_lines(wpdf); E = L.word_lines(epdf)
        except Exception as e:
            summary[key] = {"error": str(e)[:120]}; continue
        rows = L.pair(W, E)
        dys = [r["dy"] for r in rows if r["dy"] is not None]
        bad = [r for r in rows if r["dy"] is not None and abs(r["dy"]) > 2.0]
        first = next((r for r in rows if r["dy"] is not None and abs(r["dy"]) > 2.0), None)
        summary[key] = {"n_word": len(W), "n_oxi": len(E), "paired": len(dys), "unpaired": len(rows) - len(dys),
                        "n_dy_gt2": len(bad), "max_dy": max((abs(d) for d in dys), default=0.0),
                        "first_bad": first, "rows": rows, "pages": [pages(wpdf), pages(epdf)]}
    json.dump(summary, open(L.OUT / tag / "summary.json", "w", encoding="utf-8"), ensure_ascii=False, indent=0)
    L.report(tag, 15)
    ok = [v for v in summary.values() if "error" not in v]
    same = sum(1 for v in ok if v["pages"][0] == v["pages"][1])
    print(f"page count equal to Word: {same}/{len(ok)}  (errors {len(summary) - len(ok)})")


if __name__ == "__main__":
    cmd = sys.argv[1]
    if cmd == "convert":
        convert(sys.argv[2], sys.argv[3])
    elif cmd == "score":
        score(sys.argv[2])
