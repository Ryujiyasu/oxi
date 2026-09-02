# -*- coding: utf-8 -*-
"""A/B one flag over the pptx corpus in LAYOUT, not pixels.

`pptx_flag_ab.py` answers "did SSIM move", which costs a full 150-DPI render
per deck per arm (~75s on the blind set) and needs a truth PDF, so the 64-deck
dev corpus is out of its reach entirely. But `--dump-layout` returns BEFORE any
PNG is drawn, and it already carries the quantity a break-changing flag moves:
`line_x_offsets`, one entry per line the engine set.

So this asks the cheaper, sharper question first -- **which decks does this
flag touch, and does it change where paragraphs BREAK or only where their lines
sit** -- and leaves the expensive oracles (PowerPoint COM in
`pptx_line_audit_com.py`, the truth PDF in `pptx_flag_ab.py`) for the decks it
names. A deck whose two dumps are identical is proof the flag does not reach
it, on the same binary, which is the arm-A half of every ledger entry.

It cannot say WHO IS RIGHT. A line-count change here is a question for
PowerPoint, not an improvement.

    python tools/metrics/pptx_dump_ab.py OXI_FDBREAK_ENABLE --decks dev
    python tools/metrics/pptx_dump_ab.py OXI_MASTERUNIT_DISABLE --decks 31,35
"""
from __future__ import annotations

import argparse
import json
import os
import re
import subprocess
import sys
import tempfile
import time
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
ROOT = REPO / "pipeline_data" / "pptx_benchmark"
EXE = REPO / "tools" / "oxi-pptx-renderer" / "target" / "release" / "oxi-pptx-renderer.exe"


def wait_for_powerpoint_to_exit(limit: float = 60.0) -> None:
    """A render started while PowerPoint still holds the embedded fonts
    resolves them against that process (`pptx_com_render_must_not_overlap`)."""
    deadline = time.time() + limit
    while time.time() < deadline:
        r = subprocess.run(["tasklist", "/FI", "IMAGENAME eq POWERPNT.EXE", "/NH"],
                           capture_output=True, text=True, check=False)
        if "POWERPNT" not in (r.stdout or ""):
            return
        time.sleep(0.5)


def dump(src: Path, flag: str, on: bool) -> dict | None:
    """The engine's layout for `src` with the feature on or off.

    `on` selects the FEATURE, not the variable, exactly as `pptx_flag_ab.py`
    does it: an `_ENABLE` flag is set to turn the feature on, a `_DISABLE` flag
    is set to turn it off. The variable is removed from the inherited
    environment first, so a value already exported cannot decide the arm.
    """
    env = dict(os.environ)
    env.pop(flag, None)
    if (not on) if flag.endswith("_DISABLE") else on:
        env[flag] = "1"
    with tempfile.TemporaryDirectory() as td:
        out = Path(td) / "layout.json"
        subprocess.run(
            [str(EXE), str(src), str(Path(td) / "slide"), "150", f"--dump-layout={out}"],
            capture_output=True, env=env, timeout=3600, check=False)
        if not out.exists():
            return None
        return json.loads(out.read_text(encoding="utf-8"))


def paragraphs(d: dict) -> dict[tuple, tuple]:
    """Every paragraph the dump holds, keyed so the two arms line up.

    The key is (slide, shape geometry, paragraph index) rather than the text:
    a flag that changes a break does not change what the paragraph SAYS, and
    two shapes on a slide can hold the same words (`pptx_line_audit_com.py`
    deck 33 s7). Geometry plus index is stable across the arms because neither
    is something a measuring flag moves.
    """
    out: dict[tuple, tuple] = {}
    for si, slide in enumerate(d.get("slides", []), start=1):
        for sh in slide.get("shapes", []):
            content = sh.get("content") or {}
            for pi, p in enumerate(content.get("paragraphs") or []):
                xs = p.get("line_x_offsets") or []
                key = (si, round(sh.get("x", 0.0), 2), round(sh.get("y", 0.0), 2),
                       round(sh.get("w", 0.0), 2), pi)
                text = "".join(r.get("text", "") for r in p.get("runs", []))
                out[key] = (len(xs), tuple(round(v, 3) for v in xs), text)
    return out


def deck_paths(spec: str) -> list[tuple[str, Path]]:
    """`dev`, `blind`, `all`, or a comma list of `31` / `d31` names."""
    dev = {m.group(1): f for f in sorted((ROOT / "dev" / "pptx").glob("*.pptx"))
           if (m := re.match(r"(d\d+)__", f.name))}
    manifest = json.loads((ROOT / "manifest.json").read_text(encoding="utf-8"))
    blind = {f"{i['idx']:02d}": ROOT / "pptx" / i["local"] for i in manifest}
    blind = {k: v for k, v in blind.items() if v.exists()}
    if spec == "dev":
        return sorted(dev.items())
    if spec == "blind":
        return sorted(blind.items())
    if spec == "all":
        return sorted(dev.items()) + sorted(blind.items())
    picked: list[tuple[str, Path]] = []
    for name in (s.strip() for s in spec.split(",") if s.strip()):
        if name.lower().startswith("d") and name.lower() in dev:
            picked.append((name.lower(), dev[name.lower()]))
        elif f"{int(name):02d}" in blind:
            key = f"{int(name):02d}"
            picked.append((key, blind[key]))
        else:
            print(f"{name}: no such deck", flush=True)
    return picked


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("flag")
    ap.add_argument("--decks", default="dev")
    args = ap.parse_args()

    decks = deck_paths(args.decks)
    if not decks:
        sys.exit("no decks selected")
    print(f"{args.flag}: OFF vs ON over {len(decks)} decks\n", flush=True)

    touched: list[str] = []
    tot_count = tot_shift = tot_paras = 0
    for name, src in decks:
        t0 = time.time()
        wait_for_powerpoint_to_exit()
        a = dump(src, args.flag, on=False)
        b = dump(src, args.flag, on=True)
        if a is None or b is None:
            print(f"{name}: render failed", flush=True)
            continue
        pa, pb = paragraphs(a), paragraphs(b)
        tot_paras += len(pa)
        shared = pa.keys() & pb.keys()
        count = [(k, pa[k], pb[k]) for k in shared if pa[k][0] != pb[k][0]]
        shift = [k for k in shared if pa[k][0] == pb[k][0] and pa[k][1] != pb[k][1]]
        gone = (pa.keys() ^ pb.keys())
        if not count and not shift and not gone:
            print(f"{name}: untouched ({len(pa)} paragraphs) [{time.time()-t0:.0f}s]",
                  flush=True)
            continue
        touched.append(name)
        tot_count += len(count)
        tot_shift += len(shift)
        more = sum(1 for _, x, y in count if y[0] > x[0])
        print(f"{name}: {len(count)} paragraphs BREAK differently "
              f"({more} gain a line, {len(count)-more} lose one), "
              f"{len(shift)} shift within the line, of {len(pa)} "
              f"[{time.time()-t0:.0f}s]", flush=True)
        if gone:
            print(f"      {len(gone)} paragraphs exist in only one arm "
                  f"-- the flag moves SHAPES, not just measurement", flush=True)
        for k, x, y in count[:6]:
            print(f"      s{k[0]:<3} {x[0]} -> {y[0]} lines  {x[2][:44]!r}", flush=True)

    print(f"\n{len(touched)}/{len(decks)} decks touched, {tot_paras} paragraphs read: "
          f"{tot_count} break differently, {tot_shift} shift within the line")
    if touched:
        print("touched: " + " ".join(touched))


if __name__ == "__main__":
    main()
