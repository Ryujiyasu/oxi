# -*- coding: utf-8 -*-
"""Iteration gate on the documents a change CAN touch.

  python subset_gate.py "<census expr>" --exe NEW.exe [--base BASE.exe] [--identity N]

For every document whose feature_census row satisfies the expression and that
has Word truth (golden pagination_word/, en/ja benchmark p1_*/word/), run the
pagination diff with NEW (and BASE, to name PASS->FAIL flips). Then take N
random documents that do NOT satisfy the expression and check that NEW and
BASE produce byte-identical layout dumps -- a wrong predicate shows up there.
Full gates remain the commit rule; this is the fast loop.
"""
import os, sys, json, glob, random, subprocess, tempfile, shutil
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "tools" / "metrics"))
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
import feature_census as FC  # noqa: E402
import pagination_diff as PD  # noqa: E402

TRUTH_DIRS = [REPO / "pipeline_data" / "pagination_word"] + \
    sorted(REPO.glob("pipeline_data/en_benchmark/p1_*/word")) + sorted(REPO.glob("pipeline_data/ja_benchmark/p1_*/word"))


def truth_for(did):
    stem = did.split("/", 1)[1]
    cands = [stem]
    if did.startswith("golden/"):
        # golden truth is keyed by the 12-hex prefix of the docx name
        cands += [stem.split("_")[0], stem[:12]]
    for d in TRUTH_DIRS:
        for c in cands:
            f = d / f"{c}.json"
            if f.exists():
                return f
    return None


def measure(exe, path):
    tmp = tempfile.mkdtemp(prefix="sg_")
    try:
        dump = os.path.join(tmp, "l.json")
        r = subprocess.run([exe, path, os.path.join(tmp, "p"), "110", "--dump-layout=" + dump], capture_output=True)
        if not os.path.exists(dump):
            return None, None
        import measure_pagination_oxi as MO
        raw = open(dump, encoding="utf-8").read()
        return MO.extract_from_dump(json.loads(raw)) if hasattr(MO, "extract_from_dump") else None, raw
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def measure_via_module(exe, path):
    os.environ["OXI_GDI_EXE"] = exe
    import importlib, measure_pagination_oxi as MO
    importlib.reload(MO)
    return MO.measure_doc(path)


def main():
    expr = sys.argv[1]
    a = sys.argv[2:]
    exe = a[a.index("--exe") + 1]
    base = a[a.index("--base") + 1] if "--base" in a else None
    n_id = int(a[a.index("--identity") + 1]) if "--identity" in a else 0
    rows = FC.load()
    hits = FC.query(expr)
    print(f"predicate: {expr}\nmatched {len(hits)} docs")
    res = []
    for did in hits:
        t = truth_for(did)
        if t is None:
            res.append((did, None, None)); continue
        word = json.loads(t.read_text(encoding="utf-8"))
        path = rows[did]["path"]
        try:
            new = PD.diff_doc(did, word, measure_via_module(exe, path))["pass"]
        except Exception as e:
            new = f"err {str(e)[:40]}"
        old = None
        if base:
            try:
                old = PD.diff_doc(did, word, measure_via_module(base, path))["pass"]
            except Exception as e:
                old = f"err {str(e)[:40]}"
        res.append((did, old, new))
    measured = [r for r in res if r[2] is not None]
    print(f"with truth: {len(measured)}   NEW pass {sum(1 for r in measured if r[2] is True)}/{len(measured)}"
          + (f"   BASE pass {sum(1 for r in measured if r[1] is True)}/{len(measured)}" if base else ""))
    for did, old, new in measured:
        tag = ""
        if base and old is True and new is not True: tag = "  <<< PASS->FAIL"
        if base and old is not True and new is True: tag = "  >>> FAIL->PASS"
        if tag or new is not True:
            print(f"  {did}: base={old} new={new}{tag}")
    for did, _, _ in res:
        if truth_for(did) is None:
            print(f"  (no truth) {did}")
    if n_id and base:
        rest = [d for d in rows if d not in set(hits) and "error" not in rows[d]]
        random.seed(0)
        bad = 0
        for did in random.sample(rest, min(n_id, len(rest))):
            _, r_new = measure(exe, rows[did]["path"]); _, r_old = measure(base, rows[did]["path"])
            same = (r_new == r_old)
            bad += (not same)
            print(f"  identity {did}: {'same' if same else 'DIFFERS <<<'}")
        print(f"identity sample: {n_id - bad}/{n_id} byte-identical")


if __name__ == "__main__":
    main()
