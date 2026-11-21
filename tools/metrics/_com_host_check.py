# -*- coding: utf-8 -*-
"""Hold oxi-com-host against Excel: the same COM Automation add-in
(tools/oxi-com-host/testcom, ProgID Oxi.TestCom) is registered for this user,
added to Excel as an Automation add-in, and each of its functions is called
from a worksheet formula and through the host with the same arguments. A
list-of-rows argument is laid on the sheet and passed as its range (Excel
hands the add-in a Range object; the host hands a SAFEARRAY of values).

The class's registration (HKCU\\Software\\Classes) and Excel's add-in entry
are removed again at the end.

    python tools/metrics/_com_host_check.py
"""
import ctypes
import json
import struct
import subprocess
import winreg
from pathlib import Path

import win32com.client

REPO = Path(__file__).resolve().parents[2]
HOST = REPO / "tools/oxi-com-host/target/release/oxi-com-host.exe"
DLL = REPO / "tools/oxi-com-host/testcom/target/release/oxi_test_com.dll"
PROGID = "Oxi.TestCom"
CLSID = "{6F1C2B7E-3A55-4E2B-9C1D-0A7E5B4C3D21}"
LIBID = "{6F1C2B7E-3A55-4E2B-9C1D-0A7E5B4C3D22}"
TLB = DLL.with_suffix(".tlb")
CASES = [("Add", [2, 3]), ("Add", [-1.25, 0.5]), ("Upper", ["hello com"]), ("Upper", ["日本語 abc"]),
         ("Sum", [[[1, 2], [3, "x"]]]), ("Sum", [7]), ("Shape", [[[1, 2, 3], [4, 5, 6]]]), ("Shape", [7]),
         ("Twice", [[[1, 2], [3, 4]]])]


def register():
    classes = winreg.HKEY_CURRENT_USER
    base = r"Software\Classes"
    for path, values in [
        (rf"{base}\{PROGID}", {"": PROGID}),
        (rf"{base}\{PROGID}\CLSID", {"": CLSID}),
        (rf"{base}\CLSID\{CLSID}", {"": PROGID}),
        (rf"{base}\CLSID\{CLSID}\ProgID", {"": PROGID}),
        (rf"{base}\CLSID\{CLSID}\Programmable", {}),
        (rf"{base}\CLSID\{CLSID}\InprocServer32", {"": str(DLL), "ThreadingModel": "Apartment"}),
        (rf"{base}\CLSID\{CLSID}\TypeLib", {"": LIBID}),
        (rf"{base}\CLSID\{CLSID}\Version", {"": "1.0"}),
    ]:
        key = winreg.CreateKey(classes, path)
        for name, value in values.items():
            winreg.SetValueEx(key, name, 0, winreg.REG_SZ, value)
        winreg.CloseKey(key)
    # Excel learns an Automation add-in's functions from its type library.
    made = ctypes.WinDLL(str(DLL)).OxiMakeTypeLib(ctypes.c_wchar_p(str(TLB)))
    assert made == 0, hex(made & 0xFFFFFFFF)


def unregister():
    def drop(path):
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, path) as key:
                children = []
                i = 0
                while True:
                    try:
                        children.append(winreg.EnumKey(key, i))
                        i += 1
                    except OSError:
                        break
            for child in children:
                drop(path + "\\" + child)
            winreg.DeleteKey(winreg.HKEY_CURRENT_USER, path)
        except FileNotFoundError:
            pass
    try:
        ctypes.WinDLL(str(DLL)).OxiDropTypeLib()
    except OSError:
        pass
    drop(rf"Software\Classes\{PROGID}")
    # Excel keeps an uninstalled add-in in its Add-in Manager list.
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Office\16.0\Excel\Add-in Manager", 0, winreg.KEY_SET_VALUE) as key:
            winreg.DeleteValue(key, PROGID)
    except OSError:
        pass
    drop(rf"Software\Classes\CLSID\{CLSID}")


def main():
    register()
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    ok = 0
    addin = None
    try:
        wb = xl.Workbooks.Add()
        addin = xl.AddIns.Add(PROGID)
        addin.Installed = True
        sh = wb.Worksheets(1)
        top = [1]

        def place(rows):
            r0 = top[0]
            for r, row in enumerate(rows):
                for c, v in enumerate(row):
                    sh.Cells(r0 + r, 10 + c).Value = v
            top[0] += len(rows) + 1
            return sh.Range(sh.Cells(r0, 10), sh.Cells(r0 + len(rows) - 1, 10 + max(map(len, rows)) - 1)).GetAddress(False, False)

        def lit(v):
            if isinstance(v, list):
                return place(v)
            if isinstance(v, bool):
                return "TRUE" if v else "FALSE"
            return json.dumps(v, ensure_ascii=False) if isinstance(v, str) else repr(v)

        formulas = []
        for i, (name, args) in enumerate(CASES, 1):
            text = f"={PROGID}.{name}({','.join(lit(a) for a in args)})"
            formulas.append(text)
            sh.Cells(i * 4, 1).Formula2 = text
        xl.CalculateFull()
        for i, (name, args) in enumerate(CASES, 1):
            cell = sh.Cells(i * 4, 1)
            excel = cell.Value
            text = cell.Text
            if cell.HasSpill:
                excel = [list(r) for r in cell.SpillingToRange.Value]
                text = json.dumps(excel)
            run = subprocess.run([str(HOST), PROGID, "call", name, json.dumps(args, ensure_ascii=False)], capture_output=True, text=True, encoding="utf-8")
            answer = json.loads(run.stdout or "{}")
            mine = answer.get("value", answer.get("error"))
            same = (text.startswith("#") and mine == text) or (isinstance(excel, float) and isinstance(mine, (int, float)) and abs(excel - mine) < 1e-12) or excel == mine
            ok += same
            print(f"  {formulas[i - 1]:42s} excel {text!r:22s} oxi {mine!r} {'OK' if same else 'DIFF'}")
        wb.Close(False)
        # The same host as the browser extension starts it (native messaging).
        host = subprocess.Popen([str(HOST), "chrome-extension://oxi-test/"], stdin=subprocess.PIPE, stdout=subprocess.PIPE)
        body = json.dumps({"id": 9, "progId": PROGID, "function": "Twice", "args": [[[1.5, 2]]]}).encode()
        host.stdin.write(struct.pack("<I", len(body)) + body)
        host.stdin.flush()
        size = struct.unpack("<I", host.stdout.read(4))[0]
        answer = json.loads(host.stdout.read(size))
        host.stdin.close()
        host.wait(timeout=10)
        fine = answer == {"id": 9, "value": [[3.0, 4.0]]}
        print(f"  native messaging Twice([[1.5, 2]]) -> {answer} {'OK' if fine else 'DIFF'}")
        ok += fine
    finally:
        if addin is not None:
            try:
                addin.Installed = False
            except Exception:
                pass
        xl.Quit()
        unregister()
    print(f"{ok}/{len(CASES) + 1} COM add-in calls match Excel (the last through native messaging)")


if __name__ == "__main__":
    main()
