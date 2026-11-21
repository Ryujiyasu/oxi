# -*- coding: utf-8 -*-
"""What Solver's VBA functions leave in a workbook, for the VBA host to match.

Through COM and Application.Run on Solver.xlam (never through the VBA project,
whose trust setting stays as it is): after SolverReset, SolverOk, SolverAdd,
SolverChange, SolverDelete and SolverOptions the sheet's own names
(`solver_*`, name and RefersTo); what SolverGet answers for each TypeNum;
SolverSolve's return code and the cells after SolverFinish. Then an
AllDifferent model (an assignment of four jobs) under each engine.

    python tools/metrics/_xlsx_vba_solver_probe.py
"""
import json
import time

import pywintypes
import win32com.client

OUT = "E:/oxi-scratch/vba_solver_probe.json"


def retry(fn):
    for _ in range(80):
        try:
            return fn()
        except pywintypes.com_error as e:
            if e.args[0] in (-2147418111, -2146777998):
                time.sleep(0.25)
                continue
            raise


def attempt(fn):
    try:
        return retry(fn)
    except pywintypes.com_error as e:
        return ("ERR", e.args[0], str(e.args[2])[:120] if len(e.args) > 2 else None)


def plain(v):
    if isinstance(v, tuple):
        return [plain(x) for x in v]
    if isinstance(v, (int, float, str, bool)) or v is None:
        return v
    return str(v)


def names(sh):
    out = {}
    for n in sh.Names:
        out[n.Name.split("!")[-1]] = n.RefersTo
    return dict(sorted(out.items()))


def main():
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    log = {}

    class Log(dict):
        # each step goes to disk as it lands: a dialog that hangs Excel
        # loses nothing already read
        def __setitem__(self, k, v):
            super().__setitem__(k, v)
            with open(OUT, "w", encoding="utf-8") as f:
                json.dump(self, f, ensure_ascii=False, indent=1, default=str)
            print(k, flush=True)

    log = Log()
    try:
        addin = [a for a in xl.AddIns if "SOLVER" in a.FullName.upper()][0]
        addin.Installed = True
        retry(lambda: xl.Workbooks.Open(addin.FullName))
        run = lambda *a: attempt(lambda: xl.Run(*a))
        wb = retry(lambda: xl.Workbooks.Add())
        sh = wb.Worksheets(1)
        sh.Name = "M"
        sh.Activate()
        for a1, v in {"A1": 0, "A2": 0, "B1": 6, "A3": "=3*A1+2*A2", "A4": "=A1+A2", "A5": "=A1+3*A2"}.items():
            if isinstance(v, str):
                sh.Range(a1).Formula = v
            else:
                sh.Range(a1).Value = v
        log["names_before"] = names(sh)
        log["SolverReset"] = run("Solver.xlam!SolverReset")
        log["names_reset"] = names(sh)
        log["SolverOk"] = run("Solver.xlam!SolverOk", "$A$3", 1, 0, "$A$1:$A$2", 2)
        log["names_ok"] = names(sh)
        log["SolverAdd1"] = run("Solver.xlam!SolverAdd", "$A$4", 1, "4")
        log["SolverAdd2"] = run("Solver.xlam!SolverAdd", "$A$5", 1, "=$B$1")
        log["SolverAdd3"] = run("Solver.xlam!SolverAdd", "$A$1:$A$2", 4)
        log["names_add"] = names(sh)
        log["SolverChange"] = run("Solver.xlam!SolverChange", "$A$4", 1, "5")
        log["names_change"] = names(sh)
        log["SolverDelete"] = run("Solver.xlam!SolverDelete", "$A$1:$A$2", 4)
        log["names_delete"] = names(sh)
        log["SolverGet"] = {t: plain(run("Solver.xlam!SolverGet", t)) for t in range(1, 11)}
        log["SolverSolve"] = run("Solver.xlam!SolverSolve", True)
        log["SolverFinish"] = run("Solver.xlam!SolverFinish", 1)
        log["cells_after"] = [sh.Range(a).Value for a in ("A1", "A2", "A3")]
        log["names_after"] = names(sh)
        # SolverOk with ValueOf and GRG; SolverOk again (does it keep constraints?)
        log["SolverOk2"] = run("Solver.xlam!SolverOk", "$A$3", 3, 12, "$A$1:$A$2", 1)
        log["names_ok2"] = names(sh)
        log["SolverSolve2"] = run("Solver.xlam!SolverSolve", True)
        log["SolverFinish2"] = run("Solver.xlam!SolverFinish", 2)
        log["cells_after2"] = [sh.Range(a).Value for a in ("A1", "A2", "A3")]
        wb.Close(False)

        # AllDifferent: four jobs to four slots, cost = SUMPRODUCT of weight x slot
        for eng in (1, 2, 3):
            wb = retry(lambda: xl.Workbooks.Add())
            sh = wb.Worksheets(1)
            sh.Activate()
            for i, (w, start) in enumerate([(4, 1), (9, 2), (2, 3), (7, 4)], 1):
                sh.Cells(i, 1).Value = start
                sh.Cells(i, 2).Value = w
            sh.Range("C1").Formula = "=SUMPRODUCT(A1:A4,B1:B4)+ABS(A1-A3)*5"
            run("Solver.xlam!SolverReset")
            run("Solver.xlam!SolverOk", "$C$1", 2, 0, "$A$1:$A$4", eng)
            run("Solver.xlam!SolverAdd", "$A$1:$A$4", 6)
            log[f"alldiff_names_{eng}"] = names(sh)
            log[f"alldiff_solve_{eng}"] = run("Solver.xlam!SolverSolve", True)
            run("Solver.xlam!SolverFinish", 1)
            log[f"alldiff_cells_{eng}"] = [sh.Cells(i, 1).Value for i in range(1, 5)] + [sh.Range("C1").Value]
            wb.Close(False)
    finally:
        xl.Quit()


if __name__ == "__main__":
    main()
