# -*- coding: utf-8 -*-
"""Build Solver truth with Excel: each model is laid out on a sheet, set up
through Solver.xlam (SolverOk / SolverAdd), saved unsolved, then solved; the
answer (return code, variable cells, objective) goes to <name>.truth.json.

    python tools/metrics/_xlsx_solver_gen.py
"""
import json
from pathlib import Path
import win32com.client

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "xlsx_solver"
LE, EQ, GE, INT, BIN = 1, 2, 3, 4, 5
MAX, MIN, VALUE = 1, 2, 3
GRG, SIMPLEX = 1, 2

# name: (cells {A1: value or formula}, objective, goal, value, variables, constraints [(lhs, rel, rhs)], engine, nonneg)
MODELS = {
    "s01_lp_max": ({"A1": 0, "A2": 0, "A3": "=3*A1+2*A2", "A4": "=A1+A2", "A5": "=A1+3*A2"},
                   "$A$3", MAX, 0, "$A$1:$A$2", [("$A$4", LE, "4"), ("$A$5", LE, "6")], SIMPLEX, True),
    "s02_lp_min_ge": ({"A1": 0, "A2": 0, "A3": "=2*A1+3*A2", "A4": "=A1+A2", "A5": "=2*A1+A2"},
                      "$A$3", MIN, 0, "$A$1:$A$2", [("$A$4", GE, "10"), ("$A$5", GE, "12")], SIMPLEX, True),
    "s03_lp_eq_cells": ({"A1": 0, "A2": 0, "A3": 0, "B1": 100, "B2": 40,
                         "A4": "=4*A1+6*A2+5*A3", "A5": "=A1+A2+A3", "A6": "=A1-A2"},
                        "$A$4", MAX, 0, "$A$1:$A$3", [("$A$5", EQ, "=$B$1"), ("$A$6", LE, "=$B$2"), ("$A$2", LE, "30")], SIMPLEX, True),
    "s04_integer": ({"A1": 0, "A2": 0, "A3": "=5*A1+4*A2", "A4": "=6*A1+4*A2", "A5": "=A1+2*A2"},
                    "$A$3", MAX, 0, "$A$1:$A$2", [("$A$4", LE, "24"), ("$A$5", LE, "6"), ("$A$1:$A$2", INT, "integer")], SIMPLEX, True),
    "s05_knapsack_bin": ({"A1": 0, "A2": 0, "A3": 0, "A4": 0, "A5": "=10*A1+13*A2+7*A3+8*A4", "A6": "=4*A1+6*A2+3*A3+5*A4"},
                         "$A$5", MAX, 0, "$A$1:$A$4", [("$A$6", LE, "10"), ("$A$1:$A$4", BIN, "binary")], SIMPLEX, True),
    "s06_grg_quadratic": ({"A1": 0, "A2": 0, "A3": "=(A1-3)^2+(A2-2)^2", "A4": "=A1+A2"},
                          "$A$3", MIN, 0, "$A$1:$A$2", [("$A$4", LE, "4")], GRG, True),
    "s07_grg_value_of": ({"A1": 1, "A2": "=A1^3-2*A1"},
                         "$A$2", VALUE, 10, "$A$1", [], GRG, True),
    "s08_grg_max_area": ({"A1": 1, "A2": 1, "A3": "=A1*A2", "A4": "=2*A1+2*A2"},
                         "$A$3", MAX, 0, "$A$1:$A$2", [("$A$4", EQ, "20")], GRG, True),
    "s09_lp_free_vars": ({"A1": 0, "A2": 0, "A3": "=A1+A2", "A4": "=A1-A2", "A5": "=A1+2*A2"},
                         "$A$3", MIN, 0, "$A$1:$A$2", [("$A$4", GE, "-2"), ("$A$5", GE, "-10"), ("$A$4", LE, "3"), ("$A$5", LE, "8")], SIMPLEX, False),
    "s10_transport": ({**{f"{c}{r}": 0 for c in "ABC" for r in (1, 2)},
                       "D1": "=SUM(A1:C1)", "D2": "=SUM(A2:C2)", "A3": "=A1+A2", "B3": "=B1+B2", "C3": "=C1+C2",
                       "E1": "=SUMPRODUCT(A1:C2,F1:H2)", "F1": 4, "G1": 6, "H1": 9, "F2": 5, "G2": 3, "H2": 8},
                      "$E$1", MIN, 0, "$A$1:$C$2", [("$D$1:$D$2", LE, "50"), ("$A$3:$C$3", GE, "30")], SIMPLEX, True),
}


def second():
    import random
    rnd = random.Random(1121)
    models = {
        "t01_infeasible": ({"A1": 0, "A2": 0, "A3": "=A1+A2", "A4": "=A1+A2"},
                           "$A$3", MAX, 0, "$A$1:$A$2", [("$A$4", LE, "2"), ("$A$4", GE, "5")], SIMPLEX, True),
        "t02_unbounded": ({"A1": 0, "A2": 0, "A3": "=A1+A2", "A4": "=A1-A2"},
                          "$A$3", MAX, 0, "$A$1:$A$2", [("$A$4", LE, "2")], SIMPLEX, True),
        "t03_grg_rosenbrock": ({"A1": -1, "A2": 1, "A3": "=100*(A2-A1^2)^2+(1-A1)^2"},
                               "$A$3", MIN, 0, "$A$1:$A$2", [], GRG, False),
        "t04_int_equality": ({"A1": 0, "A2": 0, "A3": 0, "A4": "=3*A1+2*A2+4*A3", "A5": "=A1+A2+A3", "A6": "=2*A1+A3"},
                             "$A$4", MIN, 0, "$A$1:$A$3", [("$A$5", EQ, "7"), ("$A$6", GE, "5"), ("$A$1:$A$3", INT, "integer")], SIMPLEX, True),
        "t05_grg_circle": ({"A1": 1, "A2": 1, "A3": "=A1+2*A2", "A4": "=A1^2+A2^2"},
                           "$A$3", MAX, 0, "$A$1:$A$2", [("$A$4", LE, "5")], GRG, False),
    }
    for k in range(6):
        n, m = rnd.randint(2, 4), rnd.randint(2, 4)
        cells = {}
        letters = "ABCD"
        for j in range(n):
            cells[f"A{j + 1}"] = 0
        obj = "=" + "+".join(f"{rnd.randint(1, 9)}*A{j + 1}" for j in range(n))
        cells["C1"] = obj
        cons = []
        for i in range(m):
            cells[f"C{i + 2}"] = "=" + "+".join(f"{rnd.randint(1, 9)}*A{j + 1}" for j in range(n))
            cons.append((f"$C${i + 2}", LE, str(rnd.randint(10, 60))))
        models[f"t{6 + k:02d}_random_lp"] = (cells, "$C$1", MAX, 0, f"$A$1:$A${n}", cons, SIMPLEX, True)
    return models


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    xl = win32com.client.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
    try:
        addin = [a for a in xl.AddIns if "SOLVER" in a.FullName.upper()][0]
        addin.Installed = True
        xl.Workbooks.Open(addin.FullName)
        import sys
        models = second() if "--second" in sys.argv else MODELS
        for name, (cells, obj, goal, val, adj, cons, eng, nonneg) in models.items():
            wb = xl.Workbooks.Add(); sh = wb.Worksheets(1); sh.Activate()
            for a1, v in cells.items():
                if isinstance(v, str) and v.startswith("="):
                    sh.Range(a1).Formula = v
                else:
                    sh.Range(a1).Value = v
            xl.Run("Solver.xlam!SolverReset")
            xl.Run("Solver.xlam!SolverOk", obj, goal, val, adj, eng)
            for lhs, rel, rhs in cons:
                if rel in (INT, BIN):
                    xl.Run("Solver.xlam!SolverAdd", lhs, rel)
                else:
                    xl.Run("Solver.xlam!SolverAdd", lhs, rel, rhs)
            # "Make unconstrained variables non-negative" is the sheet name
            # solver_neg: 1 yes, 2 no.
            sh.Names.Add("solver_neg", "=1" if nonneg else "=2")
            path = OUT / f"{name}.xlsx"
            if path.exists():
                path.unlink()
            wb.SaveAs(str(path), 51)
            code = xl.Run("Solver.xlam!SolverSolve", True)
            xl.Run("Solver.xlam!SolverFinish", 1)
            adj_cells = [c.Value for c in sh.Range(adj).Cells]
            truth = {"code": code, "variables": adj_cells, "objective": sh.Range(obj).Value}
            (OUT / f"{name}.truth.json").write_text(json.dumps(truth))
            wb.Close(False)
            print(name, truth)
    finally:
        xl.Quit()


if __name__ == "__main__":
    main()
