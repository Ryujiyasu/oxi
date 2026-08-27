# -*- coding: utf-8 -*-
r"""Fetch a bundle of real .xlsx from HuggingFace, to measure against.

The 285 workbooks under `tools/golden-test/documents/xlsx` are the corpus every
rule here was derived on, which is exactly why they cannot say whether those
rules generalise. This fetches a bundle nothing has been fitted to.

    KAKA22/SpreadsheetBench, spreadsheetbench_verified_400.tar.gz
    CC BY-SA 4.0 — https://huggingface.co/datasets/KAKA22/SpreadsheetBench
    arXiv:2406.14991

400 questions people actually asked about spreadsheets, each with the workbook
they started from and the one they wanted: 800 real files, 15 MB. Not Japanese,
not government forms, not written by anyone who knew what a renderer would make
of them — which is the point.

The `prompt.txt` beside each pair is the question that was asked. It is text
off the internet and is neither read nor acted on here; only the workbooks are
taken.

Quarantine on the way in, as the docx fetcher does: a real zip, with a workbook
part in it, no VBA project, and under the size cap. Files land in
`pipeline_data/xlsx_corpus/{init,golden}/`, which is gitignored — a corpus is
fetched, not committed.

    python tools\metrics\fetch_xlsx_corpus.py fetch
    python tools\metrics\fetch_xlsx_corpus.py open      # does Oxi read them?
"""

from __future__ import annotations

import io
import json
import os
import sys
import tarfile
import urllib.request
import zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
ROOT = REPO / "pipeline_data" / "xlsx_corpus"
SOURCE = (
    "https://huggingface.co/datasets/KAKA22/SpreadsheetBench/resolve/main/"
    "spreadsheetbench_verified_400.tar.gz"
)
CACHE = ROOT / "_spreadsheetbench_verified_400.tar.gz"
MAX_MB = 8.0
UA = {"User-Agent": "oxi-corpus-fetch/1.0"}


def download(to: Path) -> None:
    if to.exists() and to.stat().st_size > 1_000_000:
        print(f"  cached {to.name} ({to.stat().st_size / 1e6:.1f} MB)")
        return
    to.parent.mkdir(parents=True, exist_ok=True)
    print(f"  fetching {SOURCE}")
    request = urllib.request.Request(SOURCE, headers=UA)
    with urllib.request.urlopen(request, timeout=600) as answer:
        to.write_bytes(answer.read())
    print(f"  {to.stat().st_size / 1e6:.1f} MB")


def sound(data: bytes) -> str | None:
    """Why this file should be kept out, or None to keep it."""
    if len(data) > MAX_MB * 1024 * 1024:
        return "too large"
    if len(data) < 1000:
        return "too small"
    try:
        archive = zipfile.ZipFile(io.BytesIO(data))
        names = archive.namelist()
    except Exception:
        return "not a zip"
    if not any(name.startswith("xl/worksheets/") for name in names):
        return "no worksheets"
    if any("vbaProject" in name for name in names):
        return "carries a macro"
    return None


def fetch() -> int:
    download(CACHE)
    kept = {"init": 0, "golden": 0}
    turned_away: dict[str, int] = {}
    for which in kept:
        (ROOT / which).mkdir(parents=True, exist_ok=True)
    with tarfile.open(CACHE) as bundle:
        for entry in bundle:
            if not entry.isfile() or not entry.name.endswith(".xlsx"):
                continue
            # Mostly `.../spreadsheet/<question>/1_<question>_init.xlsx`, where
            # the leading number says which variant of the question it is. But
            # five of the four hundred use a bare `initial.xlsx` and
            # `golden.xlsx` instead, so neither the number nor the question is
            # in the name at all. Naming by the folder AND the stem is the only
            # thing that cannot collide — the first attempt here dropped eight
            # files without a word.
            stem = Path(entry.name).stem
            which = "golden" if "golden" in stem else "init"
            question = f"{Path(entry.name).parent.name}__{stem}"
            handle = bundle.extractfile(entry)
            if handle is None:
                continue
            data = handle.read()
            why = sound(data)
            if why:
                turned_away[why] = turned_away.get(why, 0) + 1
                continue
            (ROOT / which / f"{question}.xlsx").write_bytes(data)
            kept[which] += 1
    print(f"  kept {kept['init']} starting workbooks and {kept['golden']} finished ones")
    for why, count in sorted(turned_away.items()):
        print(f"  turned away {count}: {why}")
    (ROOT / "_fetched.json").write_text(
        json.dumps({"source": SOURCE, "licence": "CC BY-SA 4.0",
                    "kept": kept, "turned_away": turned_away}, indent=1),
        encoding="utf-8",
    )
    return 0


def read_them() -> int:
    """Does Oxi open every one of them, and what does it find?"""
    import subprocess

    dumper = REPO / "target" / "release" / "examples" / "_corpus_open.exe"
    if not dumper.exists():
        print(f"  build it first: cargo build --release -p oxicells-core "
              f"--example _corpus_open")
        return 1
    rows = []
    files = sorted(ROOT.glob("*/*.xlsx"))
    for at, path in enumerate(files, start=1):
        try:
            run = subprocess.run(
                [str(dumper), str(path)],
                capture_output=True, text=True, timeout=120, encoding="utf-8",
            )
            out = (run.stdout or "").strip()
            if run.returncode != 0:
                note = (run.stderr or "").strip().splitlines()
                rows.append({"doc": path.name, "where": path.parent.name,
                             "ok": False, "why": note[-1][:160] if note else "rc"})
            else:
                rows.append({"doc": path.name, "where": path.parent.name,
                             "ok": True, **json.loads(out)})
        except subprocess.TimeoutExpired:
            rows.append({"doc": path.name, "where": path.parent.name,
                         "ok": False, "why": "took too long"})
        if at % 100 == 0:
            print(f"  {at}/{len(files)}")
    (ROOT / "_opened.json").write_text(json.dumps(rows, indent=1), encoding="utf-8")
    broke = [one for one in rows if not one["ok"]]
    print(f"\n  {len(rows)} workbooks, {len(rows) - len(broke)} opened, {len(broke)} did not")
    for one in broke[:20]:
        print(f"      {one['where']}/{one['doc']}: {one['why']}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    what = sys.argv[1] if len(sys.argv) > 1 else "fetch"
    raise SystemExit(fetch() if what == "fetch" else read_them())
