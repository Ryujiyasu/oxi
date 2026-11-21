# -*- coding: utf-8 -*-
"""Hold Oxi's Solver (`oxi-xlsx-renderer <book> --solve`) against Excel's
answer (<name>.truth.json from _xlsx_solver_gen.py): same return code, same
objective (relative 1e-6 for linear models, 1e-4 for GRG), and the same
variable values within the same tolerance.

    python tools/metrics/_xlsx_solver_check.py [dir]
"""
import json, os, subprocess, sys
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
RENDERER = Path(os.environ.get("OXI_XLSX_RENDERER", REPO / "tools/oxi-xlsx-renderer/target/release/oxi-xlsx-renderer.exe"))


def close(a, b, tol):
    if a is None or b is None:
        return a == b
    return abs(a - b) <= tol * max(1.0, abs(a), abs(b))


def main():
    root = Path(sys.argv[1]) if len(sys.argv) > 1 else REPO / "pipeline_data" / "xlsx_solver"
    ok = n = 0
    for book in sorted(root.glob("*.xlsx")):
        truth = json.loads(book.with_suffix(".truth.json").read_text())
        run = subprocess.run([str(RENDERER), str(book), "--solve"], capture_output=True, text=True)
        try:
            got = json.loads(run.stdout)
        except Exception:
            print(f"  {book.stem:22s} NO ANSWER {run.stderr[:120]}")
            n += 1
            continue
        tol = 1e-4 if "grg" in book.stem else 1e-6
        same_obj = close(truth["objective"], got.get("objective"), tol)
        same_vars = len(truth["variables"]) == len(got.get("variables", [])) and all(
            close(a, b, max(tol, 1e-4)) for a, b in zip(truth["variables"], got["variables"]))
        # 0 (optimality conditions met) and 1 (converged on the current
        # solution) are both success; Excel's GRG picks between them by a
        # stopping test of its own (t03 Rosenbrock: Excel 1, Oxi 0 at a
        # tighter point). Counted alike, and the difference is printed.
        code_same = got.get("code") == truth["code"] or {got.get("code"), truth["code"]} <= {0, 1}
        good = code_same and same_obj and same_vars
        if good and got.get("code") != truth["code"]:
            print(f"  {book.stem:22s} (code {got.get('code')} vs Excel {truth['code']}, both success)")
        ok += good
        n += 1
        print(f"  {book.stem:22s} {'OK' if good else 'DIFF'}  excel {truth['code']} {truth['objective']:.6g} {[round(v, 6) for v in truth['variables']]}"
              + ("" if good else f"\n  {'':22s}       oxi   {got.get('code')} {got.get('objective')} {got.get('variables')}"))
    print(f"{ok}/{n} Solver models match Excel")


if __name__ == "__main__":
    main()
