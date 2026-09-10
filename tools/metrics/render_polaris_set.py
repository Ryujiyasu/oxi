# -*- coding: utf-8 -*-
"""Export a whole frozen set to PDF through Polaris Office, unattended.

Polaris has no command line and no usable COM: the only way in is its own
window, driven by `_polaris_export.py`. That makes this the slowest engine in
the comparison by a wide margin and the only one that takes the screen while it
runs — nothing else can use the mouse or keyboard until it finishes.

    python tools/metrics/render_polaris_set.py en D
    python tools/metrics/render_polaris_set.py ja D

Each document gets its own process and its own timeout, so one that hangs the
window costs one document rather than the run. Finished PDFs are kept, so the
command can be stopped and started again and picks up where it left off.
"""
from __future__ import annotations

import json
import subprocess
import sys
import time
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
EXPORT = REPO / "tools" / "metrics" / "_polaris_export.py"
# Polaris opens, lays out and writes a PDF through its own UI; a long document
# takes a while, and one that has gone wrong takes forever.
EACH = 240


def main() -> int:
    if len(sys.argv) < 3 or sys.argv[1] not in ("en", "ja"):
        print(f"usage: {Path(sys.argv[0]).name} [en|ja] <rotation letter>")
        return 2
    lang, letter = sys.argv[1], sys.argv[2].upper()
    bench = REPO / "pipeline_data" / f"{lang}_benchmark"
    stem = "_final_blind" if lang == "en" else "_final_jablind"
    selection = bench / f"{stem}{letter}50.json"
    out = bench / f"ssim_blind{letter}50" / "polaris_pdf"
    out.mkdir(parents=True, exist_ok=True)

    data = json.loads(selection.read_text(encoding="utf-8"))
    docs = [(f"{kind}__{Path(e['path']).stem}", Path(e["path"]))
            for kind, entries in data.items() for e in entries]
    print(f"{lang} blind-{letter}: {len(docs)} documents through Polaris")

    done = failed = 0
    for n, (doc_id, path) in enumerate(docs, 1):
        dest = out / f"{doc_id}.pdf"
        if dest.is_file() and dest.stat().st_size > 1024:
            print(f"[{n:2}/{len(docs)}] {doc_id} already done", flush=True)
            done += 1
            continue
        began = time.time()
        try:
            run = subprocess.run(
                [sys.executable, str(EXPORT), str(path.resolve()), str(dest)],
                capture_output=True, text=True, timeout=EACH,
                encoding="utf-8", errors="replace")
            ok = dest.is_file() and dest.stat().st_size > 1024
            note = "" if ok else (run.stdout or run.stderr or "")[-80:].replace("\n", " ")
        except subprocess.TimeoutExpired:
            ok, note = False, f"gave up after {EACH}s"
        if ok:
            done += 1
        else:
            failed += 1
            dest.unlink(missing_ok=True)
        print(f"[{n:2}/{len(docs)}] {doc_id} {'OK' if ok else 'FAILED ' + note}"
              f" {time.time() - began:.0f}s", flush=True)

    print(f"\n{done} exported, {failed} would not export")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    raise SystemExit(main())
