# -*- coding: utf-8 -*-
"""Phase 1 across the frozen JA sets — the EN side's `_p1_status.py`, for ja.

The EN benchmark has carried a Phase-1 driver for all six of its sets since it
was built. The JA side never got one, so Word's per-paragraph truth exists for
jablind50 and jablindB50 only and the other two sets have never been through
the gate at all. That is not a low score, it is an unmeasured one.

    python _ja_p1.py                 # every set that has Word truth
    python _ja_p1.py jablindD50      # one set
    python _ja_p1.py --measure-word jablindC50 jablindD50

`--measure-word` fills the truth in (Word COM, about a second per page) and is
the reason this file exists; after that the plain form reports the gate.
"""
import json
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "tools" / "metrics"))
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

import measure_pagination_oxi as MO  # noqa: E402
import pagination_diff as PD  # noqa: E402

BENCH = REPO / "pipeline_data" / "ja_benchmark"
SETS = {
    "jablind50": ("_final_jablind50.json", "p1_blind50"),
    "jablindB50": ("_final_jablindB50.json", "p1_blindB50"),
    "jablindC50": ("_final_jablindC50.json", "p1_blindC50"),
    "jablindD50": ("_final_jablindD50.json", "p1_blindD50"),
}


def selected(final_name: str):
    """(doc id, path) for one frozen set, named the way the EN side names them."""
    final = json.load(open(BENCH / final_name, encoding="utf-8"))
    out = []
    for _topic, entries in final.items():
        for entry in entries:
            p = Path(entry["path"])
            out.append((f"{p.parent.name}__{p.stem}", str(p.resolve())))
    return out


def measure_word(names):
    """Fill in Word's per-paragraph truth for the named sets."""
    import measure_pagination_word as MW
    for name in names:
        final_name, p1 = SETS[name]
        out = BENCH / p1 / "word"
        out.mkdir(parents=True, exist_ok=True)
        docs = selected(final_name)
        print(f"{name}: {len(docs)} documents -> {out}")
        import win32com.client
        # DispatchEx, not Dispatch: quitting after the first set leaves a dying
        # instance behind, and Dispatch attaches to it — the second set then
        # cannot even set Visible. A fresh instance per set sidesteps that.
        app = win32com.client.DispatchEx("Word.Application")
        for attr, value in (("Visible", False), ("DisplayAlerts", False)):
            try:
                setattr(app, attr, value)
            except Exception:  # noqa: BLE001
                pass
        try:
            for n, (doc_id, path) in enumerate(docs, 1):
                at = out / f"{doc_id}.json"
                if at.is_file():
                    continue
                try:
                    got = MW.measure_doc(app, path)
                except Exception as exc:  # noqa: BLE001
                    print(f"  [{n:3}/{len(docs)}] {doc_id}: {exc}", flush=True)
                    continue
                at.write_text(json.dumps(got, ensure_ascii=False), encoding="utf-8")
                print(f"  [{n:3}/{len(docs)}] {doc_id}: "
                      f"{len(got.get('paragraphs', []))} paras, {got.get('pages')} pages",
                      flush=True)
        finally:
            app.Quit()


def report(only=None):
    total_pass = total = 0
    for name, (final_name, p1) in SETS.items():
        if only and name != only:
            continue
        truth_dir = BENCH / p1 / "word"
        if not truth_dir.is_dir():
            print(f"{name}: no Word truth yet (run --measure-word {name})")
            continue
        npass = ntot = 0
        fails = []
        for doc_id, path in selected(final_name):
            truth = truth_dir / f"{doc_id}.json"
            if not truth.is_file():
                continue
            word = json.loads(truth.read_text(encoding="utf-8"))
            try:
                oxi = MO.measure_doc(path)
            except Exception as exc:  # noqa: BLE001
                fails.append((doc_id, f"render: {exc}"))
                ntot += 1
                continue
            got = PD.diff_doc(doc_id, word, oxi)
            ntot += 1
            if got["pass"]:
                npass += 1
            else:
                fails.append((doc_id, f"score={got['score']:.4f} "
                                      f"pcd={got['page_count_delta']:+d} "
                                      f"{got['delta_histogram']}"))
        total_pass += npass
        total += ntot
        rate = 100.0 * npass / ntot if ntot else 0.0
        print(f"{name}: {npass}/{ntot} = {rate:.2f}%")
        for doc_id, why in fails[:8]:
            print(f"    {doc_id}: {why}")
        if len(fails) > 8:
            print(f"    ... and {len(fails) - 8} more")
    if total:
        print(f"\nJA Phase 1: {total_pass}/{total} = {100.0 * total_pass / total:.2f}%")


def main() -> int:
    args = sys.argv[1:]
    if args and args[0] == "--measure-word":
        names = [a for a in args[1:] if a in SETS] or list(SETS)
        measure_word(names)
        return 0
    report(args[0] if args else None)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
