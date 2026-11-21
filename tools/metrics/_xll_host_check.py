# -*- coding: utf-8 -*-
"""Hold oxi-xll-host against Excel: the same XLL (tools/oxi-xll-host/testxll)
is registered in Excel (RegisterXLL) and loaded by the host, and each
function is called both ways with the same arguments.

    python tools/metrics/_xll_host_check.py
"""
import json, subprocess
from pathlib import Path
import win32com.client

REPO = Path(__file__).resolve().parents[2]
HOST = REPO / "tools/oxi-xll-host/target/release/oxi-xll-host.exe"
XLL = REPO / "tools/oxi-xll-host/testxll/target/release/oxi_test.xll"
CASES = [("OxiAdd", [2, 3]), ("OxiAdd", [-1.25, 0.5]), ("OxiScale", [4, 2.5]), ("OxiScale", [-3, 7]),
         ("OxiUpper", ["hello xll"]), ("OxiUpper", ["日本語 abc"]), ("OxiHypot4", [1, 2, 3, 4]),
         ("OxiSum6", [1, 2, 3, 4, 5, 6]), ("OxiSum6", [0.5, -1, 2, 0, 3, 1]), ("OxiHalf", [7]),
         ("OxiHalf", ["x"]), ("OxiHalf", [True])]


def formula(name, args):
    lit = lambda v: ("TRUE" if v else "FALSE") if isinstance(v, bool) else json.dumps(v, ensure_ascii=False) if isinstance(v, str) else repr(v)
    return f"={name}({','.join(lit(a) for a in args)})"


def main():
    xl = win32com.client.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
    ok = 0
    try:
        assert xl.RegisterXLL(str(XLL))
        sh = xl.Workbooks.Add().Worksheets(1)
        for i, (name, args) in enumerate(CASES, 1):
            sh.Cells(i, 1).Formula = formula(name, args)
        xl.Calculate()
        for i, (name, args) in enumerate(CASES, 1):
            excel = sh.Cells(i, 1).Value
            text = sh.Cells(i, 1).Text
            run = subprocess.run([str(HOST), str(XLL), "call", name, json.dumps(args, ensure_ascii=False)], capture_output=True, text=True, encoding="utf-8")
            mine = json.loads(run.stdout).get("value")
            same = (text.startswith("#") and mine == text) or (isinstance(excel, float) and isinstance(mine, (int, float)) and abs(excel - mine) < 1e-12) or excel == mine
            ok += same
            print(f"  {formula(name, args):34s} excel {text!r:14s} oxi {mine!r} {'OK' if same else 'DIFF'}")
    finally:
        xl.Quit()
    print(f"{ok}/{len(CASES)} XLL calls match Excel")


if __name__ == "__main__":
    main()
