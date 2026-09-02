# -*- coding: utf-8 -*-
"""Add decks to the DEV corpus from the same dataset the blind corpus came from.

The blind 50 are a random sample of `noxneural/pptx_collection_templates`
(`_download_pptx_sample.py`). This takes a DIFFERENT, non-overlapping sample and
files it under `dev/pptx` as `dNN__`.

★Why this is cheap now: `pptx_line_audit_com.py` asks PowerPoint itself, so a new
deck needs NO truth PDF and no render to be useful -- it can be black-box tested
against the application the moment it lands. Every earlier corpus expansion had
to pay for a PowerPoint PDF export per deck first.

Decks already in either corpus are excluded by basename, and the sample is
seeded so a re-run adds the same decks rather than a new random set.

    python tools/metrics/fetch_pptx_dev_decks.py --count 24
    python tools/metrics/fetch_pptx_dev_decks.py --count 24 --dry-run
"""
from __future__ import annotations

import argparse
import json
import os
import random
import re
import sys
from pathlib import Path

import requests

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
BENCH = REPO / "pipeline_data" / "pptx_benchmark"
DEV = BENCH / "dev" / "pptx"
DS = "noxneural/pptx_collection_templates"
SEED = 20260902


def sanitize(name: str) -> str:
    name = re.sub(r'[\\/:*?"<>|]', "_", name)
    name = re.sub(r"\s+", "_", name.strip())
    return name[:80] or "slide"


def listing() -> list[str]:
    """Every .pptx path in the dataset, following the API's pagination."""
    out: list[str] = []
    url = f"https://huggingface.co/api/datasets/{DS}/tree/main"
    params = {"recursive": "true", "expand": "false"}
    while True:
        r = requests.get(url, params=params, timeout=180)
        r.raise_for_status()
        items = r.json()
        if not items:
            break
        out += [i["path"] for i in items
                if isinstance(i, dict) and i.get("path", "").lower().endswith(".pptx")]
        link = r.headers.get("Link", "")
        m = re.search(r'<([^>]+)>;\s*rel="next"', link)
        if not m:
            break
        url, params = m.group(1), None
    return out


def taken() -> set[str]:
    """Basenames already in either corpus, however they were named locally."""
    seen = set()
    sp = BENCH / "sample_paths.json"
    if sp.exists():
        for p in json.loads(sp.read_text(encoding="utf-8")):
            seen.add(os.path.basename(p).lower())
    for d in (BENCH / "pptx", DEV):
        for f in d.glob("*.pptx"):
            # local names carry an `NN__` prefix and a `.pptx` suffix
            base = re.sub(r"^d?\d+__", "", f.name)
            seen.add(base.lower())
            seen.add(base.replace(".pptx.pptx", ".pptx").lower())
    return seen


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--count", type=int, default=24)
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    all_paths = listing()
    used = taken()
    fresh = [p for p in all_paths if os.path.basename(p).lower() not in used]
    print(f"{len(all_paths)} decks in the dataset, {len(used)} already taken, "
          f"{len(fresh)} available")

    rng = random.Random(SEED)
    picked = rng.sample(fresh, min(args.count, len(fresh)))
    start = 1 + max([int(m.group(1)) for f in DEV.glob("*.pptx")
                     if (m := re.match(r"d(\d+)__", f.name))] or [0])
    print(f"adding {len(picked)} decks starting at d{start:02d}\n")

    DEV.mkdir(parents=True, exist_ok=True)
    added = []
    for i, path in enumerate(picked, start=start):
        local = f"d{i:02d}__{sanitize(os.path.basename(path))}"
        dest = DEV / local
        if args.dry_run:
            print(f"  d{i:02d}  {os.path.basename(path)[:64]}")
            added.append({"idx": i, "hf_path": path, "local": local})
            continue
        url = f"https://huggingface.co/datasets/{DS}/resolve/main/{requests.utils.quote(path)}"
        try:
            with requests.get(url, timeout=600, stream=True) as r:
                r.raise_for_status()
                head = next(r.iter_content(8))
                # A dataset entry that is not a zip is not a pptx, whatever it
                # is named; writing it would fail later and blame the parser.
                if head[:2] != b"PK":
                    print(f"  d{i:02d}  SKIP (not a zip): {path[:50]}")
                    continue
                with dest.open("wb") as fh:
                    fh.write(head)
                    for chunk in r.iter_content(1 << 20):
                        if chunk:
                            fh.write(chunk)
            print(f"  d{i:02d}  {dest.stat().st_size:>9,}b  {local[:58]}")
            added.append({"idx": i, "hf_path": path, "local": local})
        except Exception as e:  # noqa: BLE001
            print(f"  d{i:02d}  ERROR {path[:40]}: {e!r}")

    man = BENCH / "dev" / "manifest_added.json"
    if not args.dry_run and added:
        prev = json.loads(man.read_text(encoding="utf-8")) if man.exists() else []
        man.write_text(json.dumps(prev + added, ensure_ascii=False, indent=1),
                       encoding="utf-8")
        print(f"\nwrote {man}")
    print(f"{len(added)} decks added")


if __name__ == "__main__":
    main()
