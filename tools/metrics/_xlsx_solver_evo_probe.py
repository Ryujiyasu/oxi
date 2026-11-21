# -*- coding: utf-8 -*-
"""Solver's Evolutionary engine, through COM and Solver.xlam.

Three models, each step saved as it lands (a dialog can hang Excel): no
bounds on the variables (what code does Solver give?), a bounded nonlinear
function with several minima, and a bounded integer problem. Return code,
cells and objective for each.

    python tools/metrics/_xlsx_solver_evo_probe.py
"""
import json
import time

import threading

import pywintypes
import win32com.client
import win32con
import win32gui
import win32process

OUT = "E:/oxi-scratch/solver_evo_probe.json"


def retry(fn):
    for _ in range(80):
        try:
            return fn()
        except pywintypes.com_error as e:
            if e.args[0] in (-2147418111, -2146777998):
                time.sleep(0.25)
                continue
            raise


def main():
    log = {}

    def say(k, v):
        log[k] = v
        with open(OUT, "w", encoding="utf-8") as f:
            json.dump(log, f, ensure_ascii=False, indent=1, default=str)
        print(k, v, flush=True)

    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    excel_pid = win32process.GetWindowThreadProcessId(xl.Hwnd)[1]
    stop = threading.Event()
    seen = []

    def watch():
        # Solver's dialogs (time limit reached, show trial solution...) block
        # SolverSolve even with UserFinish: read each one's text into the log
        # and answer it with its default button.
        while not stop.is_set():
            def each(hwnd, _):
                kind = win32gui.GetClassName(hwnd)
                if kind in ("#32770", "ThunderDFrame") and win32gui.IsWindowVisible(hwnd):
                    if win32process.GetWindowThreadProcessId(hwnd)[1] == excel_pid:
                        texts = []
                        win32gui.EnumChildWindows(hwnd, lambda c, acc: acc.append(win32gui.GetWindowText(c)), texts)
                        seen.append([win32gui.GetWindowText(hwnd)] + [t for t in texts if t])
                        # Solver's own forms (試行状況の表示) stop on Esc; a
                        # message box takes its default button.
                        key = win32con.VK_ESCAPE if kind == "ThunderDFrame" else win32con.VK_RETURN
                        win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, key, 0)
                        win32gui.PostMessage(hwnd, win32con.WM_KEYUP, key, 0)
                return True
            try:
                win32gui.EnumWindows(each, None)
            except Exception:
                pass
            time.sleep(1.0)

    threading.Thread(target=watch, daemon=True).start()
    try:
        addin = [a for a in xl.AddIns if "SOLVER" in a.FullName.upper()][0]
        addin.Installed = True
        retry(lambda: xl.Workbooks.Open(addin.FullName))
        run = lambda *a: retry(lambda: xl.Run(*a))
        models = {
            # f = (x-2)^2 + (y+1)^2 + 3*SIN(3x)^2 : minima at several x
            "bounded_nonlinear": ({"A1": 0, "A2": 0, "A3": "=(A1-2)^2+(A2+1)^2+3*SIN(3*A1)^2"},
                                  "$A$3", 2, "$A$1:$A$2", [("$A$1:$A$2", 1, "5"), ("$A$1:$A$2", 3, "-5")]),
            "bounded_integer": ({"A1": 0, "A2": 0, "A3": "=ABS(7*A1-5*A2-3)+A1+A2"},
                                "$A$3", 2, "$A$1:$A$2", [("$A$1:$A$2", 1, "10"), ("$A$1:$A$2", 3, "1"), ("$A$1:$A$2", 4, None)]),
            "unbounded": ({"A1": 0, "A2": "=(A1-3)^2"}, "$A$2", 2, "$A$1", []),
        }
        for name, (cells, obj, goal, adj, cons) in models.items():
            wb = retry(lambda: xl.Workbooks.Add())
            sh = wb.Worksheets(1)
            sh.Activate()
            for a1, v in cells.items():
                if isinstance(v, str):
                    sh.Range(a1).Formula = v
                else:
                    sh.Range(a1).Value = v
            run("Solver.xlam!SolverReset")
            # No limit that stops it with Solver's "show trial solution" form
            # (time, iterations, subproblems): it ends when it stops
            # improving (MaxTimeNoImp 5 s), seed fixed.
            run("Solver.xlam!SolverOptions", 300, 32767, 0.000001, False, False, 1, 1, 1, 1, False, 0.0001, True,
                100, 1, False, True, 100000, 100000, 0.075, 5)
            run("Solver.xlam!SolverOk", obj, goal, 0, adj, 3)
            for lhs, rel, rhs in cons:
                if rhs is None:
                    run("Solver.xlam!SolverAdd", lhs, rel)
                else:
                    run("Solver.xlam!SolverAdd", lhs, rel, rhs)
            say(f"{name}_start", time.time())
            say(f"{name}_code", run("Solver.xlam!SolverSolve", True))
            say(f"{name}_dialogs", list(seen))
            seen.clear()
            run("Solver.xlam!SolverFinish", 1)
            say(f"{name}_cells", [sh.Range(a).Value for a in cells if a != obj.replace("$", "")] + [sh.Range(obj).Value])
            say(f"{name}_end", time.time())
            wb.Close(False)
    finally:
        stop.set()
        xl.Quit()


if __name__ == "__main__":
    main()
