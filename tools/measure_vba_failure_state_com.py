# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.
"""Measure workbook state after ending an unhandled VBA runtime error.

Requires Windows, pywin32 and caller-enabled VBA project access. Never changes
Office trust settings. Only dialogs belonging to our DispatchEx instance are
handled. A timeout kills that instance through its held process handle and is
recorded as a failed measurement, never as an expected VBA result.
"""
import argparse
import gc
import json
from pathlib import Path
import re
import threading
import time

import pythoncom
import win32api
import win32com.client
import win32con
import win32gui
import win32process

PROCEDURES = (
    'Success SuccessNoWrite DivideFailure RaiseFailure AbortFailure ManualFailure '
    'FormulaFailure ReadThenFailure FormatFailure StatusFailure NoWriteFailure '
    'HostFailure AddSheetFailure DeleteSheetFailure RenameSheetFailure '
    'ActivateFailure NameFailure HiddenRowFailure ResumeNextSuccess HandlerSuccess '
    'NestedFailure'
).split()


def measure(output):
    root = Path(__file__).resolve().parents[1]
    source = (root / 'tests/fixtures/vba/failure_state.bas').read_text(encoding='utf-8-sig')
    excel = book = component = process = None
    result = {'cases': {}}
    dead = False
    try:
        excel = win32com.client.gencache.EnsureDispatch(win32com.client.DispatchEx('Excel.Application'))
        pid = win32process.GetWindowThreadProcessId(excel.Hwnd)[1]
        process = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_TERMINATE, False, pid)
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AutomationSecurity = 1
        result['environment'] = {'version': excel.Version, 'build': excel.Build, 'ui': excel.LanguageSettings.LanguageID(2)}
        for procedure in PROCEDURES:
            record = {}
            result['cases'][procedure] = record
            done = threading.Event()

            def monitor():
                started = time.monotonic()
                clicked = set()
                while not done.wait(.1):
                    def visit(hwnd, unused):
                        try:
                            if win32process.GetWindowThreadProcessId(hwnd)[1] != pid or not win32gui.IsWindowVisible(hwnd):
                                return
                            controls = []
                            win32gui.EnumChildWindows(hwnd, lambda child, unused: controls.append((child, win32gui.GetDlgCtrlID(child), win32gui.GetWindowText(child))), None)
                            # The runtime error dialog has a dedicated End command.
                            ends = [item for item in controls if item[1] == 4800]
                            messages = [text for _, ident, text in controls if ident == 4803]
                            if ends and messages and hwnd not in clicked:
                                clicked.add(hwnd)
                                record.setdefault('dialogs', []).append({'title': win32gui.GetWindowText(hwnd), 'text': messages})
                                match = re.search(r"'(-?\d+)'", messages[0])
                                if match:
                                    record['runtime_number'] = int(match[1])
                                child, ident, _ = ends[0]
                                win32gui.PostMessage(hwnd, win32con.WM_COMMAND, ident, child)
                        except win32gui.error:
                            pass
                    win32gui.EnumWindows(visit, None)
                    if time.monotonic() - started > 40:
                        record['timeout'] = True
                        win32api.TerminateProcess(process, 1)
                        return

            thread = threading.Thread(target=monitor, daemon=True)
            thread.start()
            try:
                book = excel.Workbooks.Add()
                record['book_name'] = book.Name
                sheet = book.Worksheets(1)
                sheet.Name = 'Sheet1'
                while book.Worksheets.Count > 1:
                    book.Worksheets(book.Worksheets.Count).Delete()
                book.Worksheets.Add(After=sheet).Name = 'Sheet2'
                sheet.Activate()
                excel.Calculation = -4105
                sheet.Range('A1').Value2 = 1
                sheet.Range('B1').Formula = '=A1*2'
                component = book.VBProject.VBComponents.Add(1)
                component.Name = 'FailureProbe'
                component.CodeModule.AddFromString(source)
                excel.Run('FailureProbe.Seed')
                book.Saved = True
                record['book_name_before_run'] = book.Name
                try:
                    excel.Run('FailureProbe.' + procedure)
                    record['returned'] = True
                except Exception as failure:
                    record['run_error'] = repr(failure)
                names = [item.Name for item in book.Worksheets]
                record['state'] = {
                    'a1': sheet.Range('A1').Value2, 'b1': sheet.Range('B1').Value2,
                    'b1_formula': sheet.Range('B1').Formula, 'c1': sheet.Range('C1').Value2,
                    'bold': bool(sheet.Range('A1').Font.Bold), 'number_format': sheet.Range('A1').NumberFormat,
                    'status_bar': excel._oleobj_.InvokeTypes(excel._oleobj_.GetIDsOfNames('StatusBar'), 0, 2, (pythoncom.VT_VARIANT, 0), ()),
                    'calculation': excel.Calculation, 'saved': bool(book.Saved),
                    'sheets': names, 'active_sheet': int(excel.ActiveSheet.Index) - 1,
                    'row2_hidden': bool(sheet.Rows(2).Hidden),
                    'created_d1': book.Worksheets('Created').Range('D1').Value2 if 'Created' in names else None,
                    'names': {item.Name: item.RefersTo for item in book.Names},
                }
                book.Close(False)
                component = book = sheet = None
            except Exception as failure:
                record['measurement_error'] = repr(failure)
            finally:
                done.set()
                thread.join(2)
            output.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
            print(procedure, 'MEASUREMENT FAILED' if 'state' not in record else 'recorded', flush=True)
            if record.get('timeout'):
                dead = True
                break
    finally:
        try:
            if not dead and excel is not None:
                try:
                    if book is not None:
                        book.Close(False)
                finally:
                    excel.Quit()
        finally:
            if process is not None:
                win32api.CloseHandle(process)
            component = book = excel = process = None
            gc.collect()
    output.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
    if len(result['cases']) != len(PROCEDURES) or any('state' not in value for value in result['cases'].values()):
        raise RuntimeError('Incomplete Excel measurements; inspect the output before using any expectations.')
    for procedure, value in result['cases'].items():
        if procedure in ('Success', 'SuccessNoWrite', 'ResumeNextSuccess', 'HandlerSuccess'):
            assert value.get('returned'), (procedure, value)
        elif procedure != 'AbortFailure':
            assert value.get('runtime_number') in (11, 513, 1004), (procedure, value)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output', required=True, type=Path)
    args = parser.parse_args()
    measure(args.output.resolve())
