# -*- coding: utf-8 -*-
"""Page-count census over the JP Phase-1 sets with ONE opt-out flag toggled.

Answers "is this ship net-positive for JP pagination?" without the full
paragraph match: page_count_delta (pcd) is the Phase-1 gate's first-order
signal and the user's stated top priority (2026-08-21).

    python _jp_p1_flag_census.py OXI_S1237_DISABLE [set ...]

Arm A = flag set (ship OFF), arm B = default (ship ON). Renders BOTH arms
with the same binary so no cached artifact can leak in -- the sets'
own `phase_oxi` skips any doc whose JSON exists, which silently scores a
previous binary (2026-09-01 incident).
"""
import json
import os
import subprocess
import sys
import tempfile
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
REPO = Path(__file__).resolve().parents[2]
GDI = os.environ.get("OXI_GDI_EXE") or str(
    REPO / "tools" / "oxi-gdi-renderer" / "target" / "release" / "oxi-gdi-renderer.exe")
BENCH = REPO / "pipeline_data" / "ja_benchmark"
SETS = {"blind50": ("_final_jablind50.json", "p1_blind50"),
        "blindB50": ("_final_jablindB50.json", "p1_blindB50")}

FLAG = sys.argv[1]
want = sys.argv[2:] or list(SETS)


def docs(setname):
    manifest, outdir = SETS[setname]
    data = json.loads((BENCH / manifest).read_text(encoding="utf-8"))
    for _t, lst in data.items():
        for c in lst:
            p = Path(c["path"])
            yield f"{p.parent.name}__{p.stem}", str(p.resolve()), BENCH / outdir


def pages(docx, flag_on):
    env = dict(os.environ)
    if flag_on:
        env[FLAG] = "1"
    else:
        env.pop(FLAG, None)
    with tempfile.TemporaryDirectory(prefix="jp1_") as t:
        dj = os.path.join(t, "l.json")
        r = subprocess.run([GDI, docx, os.path.join(t, "p"), "--dump-layout=" + dj],
                           capture_output=True, env=env, timeout=300)
        if r.returncode != 0 or not os.path.exists(dj):
            return None
        with open(dj, encoding="utf-8") as f:
            return len(json.load(f)["pages"])


rows = []
for setname in want:
    for did, path, outdir in docs(setname):
        wf = outdir / "word" / f"{did}.json"
        if not wf.exists():
            continue
        wn = json.loads(wf.read_text(encoding="utf-8"))["n_pages"]
        off = pages(path, True)    # ship disabled
        on = pages(path, False)    # default
        if off is None or on is None:
            print("  RENDER FAIL %s" % did)
            continue
        rows.append((setname, did, wn, on, off))
        if on != off:
            print("  %-11s %-34s word=%-3d ON=%-3d OFF=%-3d  pcd %+d -> %+d %s"
                  % (setname, did, wn, on, off, on - wn, off - wn,
                     "BETTER" if abs(off - wn) < abs(on - wn)
                     else ("WORSE" if abs(off - wn) > abs(on - wn) else "same")))

n = len(rows)
ok_on = sum(1 for r in rows if r[3] == r[2])
ok_off = sum(1 for r in rows if r[4] == r[2])
moved = [r for r in rows if r[3] != r[4]]
better = sum(1 for r in moved if abs(r[4] - r[2]) < abs(r[3] - r[2]))
worse = sum(1 for r in moved if abs(r[4] - r[2]) > abs(r[3] - r[2]))
print("\n%s over %d docs: moved %d | turning it OFF: better %d / worse %d"
      % (FLAG, n, len(moved), better, worse))
print("pcd==0   ship ON (default): %d/%d      ship OFF: %d/%d" % (ok_on, n, ok_off, n))
print("sum|pcd| ship ON: %d   ship OFF: %d"
      % (sum(abs(r[3] - r[2]) for r in rows), sum(abs(r[4] - r[2]) for r in rows)))
