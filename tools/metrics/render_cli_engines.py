# -*- coding: utf-8 -*-
"""Render a frozen blind rotation with the command-line engines that export a PDF
on their own: ibisOffice (`ibisOffice --export pdf --out-dir D FILE`) and
WordCraft (`wordcraft-cli convert IN OUT.pdf`).

Both write a vector PDF per document, which `score_engine_column.py` scores
against the Word truth exactly as it scores LibreOffice / ONLYOFFICE / Polaris.

    python tools/metrics/render_cli_engines.py en I ibis
    python tools/metrics/render_cli_engines.py ja I wordcraft

Binaries are the portable release zips under C:/Users/ryuji/tools/ (SHA-256
checked against the release's own checksum file when downloaded); override with
OXI_IBIS_EXE / OXI_WORDCRAFT_EXE.
"""
from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
TOOLS = Path(r"C:\Users\ryuji\tools")
EXES = {
    "ibis": Path(os.environ.get("OXI_IBIS_EXE", TOOLS / "ibisoffice-0.8.4" / "ibisOffice" / "ibisOffice.exe")),
    "wordcraft": Path(os.environ.get("OXI_WORDCRAFT_EXE", TOOLS / "wordcraft-0.10.0" / "wordcraft-0.10.0-windows-x64-portable" / "wordcraft-cli.exe")),
}
FOLDER = {"ibis": "ibis_pdf", "wordcraft": "wordcraft_pdf"}
TIMEOUT = 300


def docs(lang: str, letter: str) -> list[tuple[str, Path]]:
    stem = "_final_blind" if lang == "en" else "_final_jablind"
    sel = json.loads((REPO / "pipeline_data" / f"{lang}_benchmark" / f"{stem}{letter}50.json").read_text(encoding="utf-8"))
    out = []
    for kind, lst in sel.items():
        for e in lst:
            name = e["path"].replace(chr(92), "/").split("/")[-1]
            path = REPO / "pipeline_data" / "docx_corpus" / lang / kind / name
            out.append((f"{kind}__{Path(name).stem}", path))
    return out


def render(engine: str, doc_id: str, src: Path, dest: Path) -> str:
    exe = EXES[engine]
    start = time.time()
    try:
        if engine == "wordcraft":
            run = subprocess.run([str(exe), "convert", str(src), str(dest)],
                                 capture_output=True, timeout=TIMEOUT)
        else:
            tmp = Path(tempfile.mkdtemp(prefix="ibis_"))
            run = subprocess.run([str(exe), "--export", "pdf", "--out-dir", str(tmp), str(src)],
                                 capture_output=True, timeout=TIMEOUT)
            made = list(tmp.glob("*.pdf"))
            if made:
                shutil.move(str(made[0]), dest)
            shutil.rmtree(tmp, ignore_errors=True)
    except subprocess.TimeoutExpired:
        return f"TIMEOUT {TIMEOUT}s"
    if dest.is_file() and dest.stat().st_size > 0:
        return f"OK {time.time() - start:.1f}s"
    tail = (run.stderr or run.stdout or b"").decode("utf-8", "replace").strip().splitlines()[-1:]
    return f"FAIL rc={run.returncode} {tail}"


def main() -> int:
    if len(sys.argv) < 4 or sys.argv[1] not in ("en", "ja") or sys.argv[3] not in EXES:
        print(f"usage: {Path(sys.argv[0]).name} [en|ja] <letter> [{'|'.join(EXES)}]")
        return 2
    lang, letter, engine = sys.argv[1], sys.argv[2].upper(), sys.argv[3]
    if not EXES[engine].is_file():
        print(f"{engine}: no binary at {EXES[engine]}")
        return 1
    out = REPO / "pipeline_data" / f"{lang}_benchmark" / f"ssim_blind{letter}50" / FOLDER[engine]
    out.mkdir(parents=True, exist_ok=True)
    todo = [(d, p) for d, p in docs(lang, letter) if not (out / f"{d}.pdf").is_file()]
    print(f"{lang} blind-{letter} {engine}: {len(todo)} to render -> {out}", flush=True)
    ok = 0
    with ThreadPoolExecutor(max_workers=4) as pool:
        futs = {pool.submit(render, engine, d, p, out / f"{d}.pdf"): d for d, p in todo}
        for n, f in enumerate(as_completed(futs), 1):
            res = f.result()
            ok += res.startswith("OK")
            print(f"[{engine} {n:2}/{len(todo)}] {futs[f]} {res}", flush=True)
    print(f"{engine}: {ok}/{len(todo)} rendered", flush=True)
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    raise SystemExit(main())
