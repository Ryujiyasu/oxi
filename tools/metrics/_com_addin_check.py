# -*- coding: utf-8 -*-
"""Hold oxi-com-addin-host against Excel: the same COM add-in
(tools/oxi-com-addin-host/testaddin, ProgID Oxi.TestAddin, IDTExtensibility2)
is connected in Excel (COMAddIns(...).Connect = True) and run by the host on
the same empty workbook; then A2 is set to 5 (in Excel through COM, in the
host with --edit) and the add-in's SheetChange sink writes F1:G1. What it
wrote is read back from both (the host's saved book opened in Excel).

The class and Excel's add-in entry are registered for this user for the run
and removed again at the end.

    python tools/metrics/_com_addin_check.py
"""
import subprocess
import tempfile
import winreg
from pathlib import Path

import win32com.client

REPO = Path(__file__).resolve().parents[2]
HOST = REPO / "target/release/oxi-com-addin-host.exe"
DLL = REPO / "tools/oxi-com-addin-host/testaddin/target/release/oxi_test_addin.dll"
PROGID = "Oxi.TestAddin"
CLSID = "{6F1C2B7E-3A55-4E2B-9C1D-0A7E5B4C3D31}"
ADDINS = rf"Software\Microsoft\Office\Excel\Addins\{PROGID}"


def register():
    hk = winreg.HKEY_CURRENT_USER
    for path, values in [
        (rf"Software\Classes\{PROGID}\CLSID", {"": CLSID}),
        (rf"Software\Classes\CLSID\{CLSID}", {"": PROGID}),
        (rf"Software\Classes\CLSID\{CLSID}\ProgID", {"": PROGID}),
        (rf"Software\Classes\CLSID\{CLSID}\InprocServer32", {"": str(DLL), "ThreadingModel": "Apartment"}),
    ]:
        with winreg.CreateKey(hk, path) as key:
            for name, value in values.items():
                winreg.SetValueEx(key, name, 0, winreg.REG_SZ, value)
    with winreg.CreateKey(hk, ADDINS) as key:
        winreg.SetValueEx(key, "FriendlyName", 0, winreg.REG_SZ, "Oxi test add-in")
        winreg.SetValueEx(key, "LoadBehavior", 0, winreg.REG_DWORD, 0)


def drop(path):
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, path) as key:
            children = []
            while True:
                try:
                    children.append(winreg.EnumKey(key, len(children)))
                except OSError:
                    break
        for child in children:
            drop(path + "\\" + child)
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, path)
    except FileNotFoundError:
        pass


def unregister():
    drop(rf"Software\Classes\{PROGID}")
    drop(rf"Software\Classes\CLSID\{CLSID}")
    drop(ADDINS)


def read(sh):
    return [sh.Range(a).Value for a in ("A1", "B1", "C1", "D1", "E1", "F1", "G1")] + [sh.Range("B1").Formula]


def main():
    register()
    work = Path(tempfile.mkdtemp(prefix="oxi_addin_"))
    base, mine = work / "base.xlsx", work / "oxi.xlsx"
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Add()
        wb.SaveAs(str(base), 51)
        xl.COMAddIns.Update()
        addin = xl.COMAddIns(PROGID)
        addin.Connect = True
        # A change after connection: the add-in's SheetChange sink answers it.
        wb.ActiveSheet.Range("A2").Value = 5
        excel = read(wb.ActiveSheet)
        addin.Connect = False
        wb.Close(False)
        run = subprocess.run([str(HOST), PROGID, str(base), str(mine), "--edit", "A2=5"], capture_output=True, text=True)
        print("  host:", run.stdout.strip() or run.stderr.strip())
        wb = xl.Workbooks.Open(str(mine))
        wb.Application.CalculateFull()
        oxi = read(wb.ActiveSheet)
        wb.Close(False)
    finally:
        xl.Quit()
        unregister()
    labels = ["A1 text", "B1 value", "C1 from B1", "D1 Application.Name", "E1 connect mode",
              "F1 SheetChange address", "G1 SheetChange value", "B1 formula"]
    ok = 0
    for label, a, b in zip(labels, excel, oxi):
        same = a == b
        ok += same
        print(f"  {label:22s} excel {a!r:22s} oxi {b!r} {'OK' if same else 'DIFF'}")
    print(f"{ok}/{len(labels)} of what the add-in wrote match Excel")


if __name__ == "__main__":
    main()
