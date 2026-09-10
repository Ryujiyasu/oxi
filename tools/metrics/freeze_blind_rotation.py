# -*- coding: utf-8 -*-
"""Freeze the next blind rotation: the next five documents per type that
nothing has seen.

The rule is in `pipeline_data/<lang>_benchmark/FROZEN_SELECTION_RULE.md` and was
declared before this ran. It is mechanical on purpose — the manifest arrives in
SHA-256 ascending order, this walks it from the top, skips every document any
earlier set already holds, and takes the next five per type that survive the
validity quarantine. Nothing about a document's content, size or difficulty
enters the choice, because a benchmark you can steer is not a benchmark.

    python tools/metrics/freeze_blind_rotation.py en D
    python tools/metrics/freeze_blind_rotation.py ja D

Quarantine (validity only, never quality): a real zip, holds
`word/document.xml`, carries no `vbaProject.bin`, at most 4 MB. A document that
fails yields its slot to the next one down the manifest.
"""
from __future__ import annotations

import io
import json
import sys
import urllib.request
import zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
CORPUS = REPO / "pipeline_data" / "docx_corpus"
TYPES = ["legal", "forms", "reports", "policies", "educational",
         "correspondence", "technical", "administrative", "creative", "reference"]
PER_TYPE = 5
MAX_MB = 4.0
UA = {"User-Agent": "oxi-corpus-fetch/1.0 (+https://github.com/Ryujiyasu/oxi)"}

# Every selection file that already claims documents, per language. A rotation
# takes what none of these hold.
TAKEN = {
    "en": ["_final.json", "_final_next50.json", "_final_blind50.json",
           "_final_blindB50.json", "_final_blindC50.json"],
    "ja": ["_final_jablind50.json", "_final_jablindB50.json", "_final_jablindC50.json"],
}


def already(lang: str) -> set[str]:
    """The sixteen-character ids every earlier set holds."""
    bench = REPO / "pipeline_data" / f"{lang}_benchmark"
    held: set[str] = set()
    for name in TAKEN[lang]:
        path = bench / name
        if not path.is_file():
            print(f"  (no {name} — nothing to exclude from it)")
            continue
        data = json.loads(path.read_text(encoding="utf-8"))
        rows = data.values() if isinstance(data, dict) else [data]
        for group in rows:
            for entry in group:
                where = entry["path"] if isinstance(entry, dict) else entry
                held.add(Path(where).stem)
    # The validation set for JA is ranks 1-5, which live on disk rather than in
    # a selection file: whatever is already downloaded has been seen.
    for kind in TYPES:
        folder = CORPUS / lang / kind
        if folder.is_dir():
            held.update(p.stem for p in folder.glob("*.docx"))
    return held


def sound(raw: bytes) -> bool:
    """Validity, not quality: this is the only test a document has to pass."""
    if len(raw) > MAX_MB * 1024 * 1024:
        return False
    try:
        zf = zipfile.ZipFile(io.BytesIO(raw))
        names = zf.namelist()
    except Exception:
        return False
    return "word/document.xml" in names and not any(
        n.endswith("vbaProject.bin") for n in names)


def take(lang: str, kind: str, skip: set[str]) -> list[Path]:
    url = f"https://api.docxcorp.us/manifest?type={kind}&lang={lang}"
    manifest = urllib.request.urlopen(
        urllib.request.Request(url, headers=UA), timeout=120).read().decode().split()
    folder = CORPUS / lang / kind
    folder.mkdir(parents=True, exist_ok=True)
    got: list[Path] = []
    looked = 0
    for link in manifest:
        if len(got) == PER_TYPE:
            break
        sha = Path(link).stem
        short = sha[:16]
        if short in skip:
            continue
        looked += 1
        try:
            raw = urllib.request.urlopen(
                urllib.request.Request(link, headers=UA), timeout=120).read()
        except Exception as error:
            print(f"    {short}: could not be fetched ({str(error)[:40]})")
            continue
        if not sound(raw):
            print(f"    {short}: fails the quarantine, slot passes on")
            continue
        at = folder / f"{short}.docx"
        at.write_bytes(raw)
        got.append(at)
    print(f"  {kind:15} {len(got)} taken, {looked} looked at past the exclusions")
    return got


def main() -> int:
    if len(sys.argv) < 3 or sys.argv[1] not in TAKEN:
        print(f"usage: {Path(sys.argv[0]).name} [en|ja] <rotation letter>")
        return 2
    lang, letter = sys.argv[1], sys.argv[2].upper()
    bench = REPO / "pipeline_data" / f"{lang}_benchmark"
    stem = "_final_blind" if lang == "en" else "_final_jablind"
    out = bench / f"{stem}{letter}50.json"
    if out.exists():
        print(f"{out.name} already exists — a frozen set is never re-cut")
        return 1

    skip = already(lang)
    print(f"{lang}: {len(skip)} documents are already spoken for\n")
    chosen: dict[str, list[dict]] = {}
    for kind in TYPES:
        chosen[kind] = [{"path": str(p)} for p in take(lang, kind, skip)]
        skip.update(Path(e["path"]).stem for e in chosen[kind])

    total = sum(len(v) for v in chosen.values())
    out.write_text(json.dumps(chosen, indent=1, ensure_ascii=False), encoding="utf-8")
    print(f"\nfrozen: {out}  ({total} documents)")
    if total != PER_TYPE * len(TYPES):
        print(f"  note: {PER_TYPE * len(TYPES)} were asked for; report the actual N")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
