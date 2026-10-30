# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.
"""Measure project storage across Excel Application.Run invocations.

Uses only owned DispatchEx instances. Requires caller-enabled VBA project
access; never changes trust settings or saves any workbook.
"""
from pathlib import Path
import argparse, gc, json, threading, time, uuid
import win32api, win32com.client, win32con, win32event, win32gui, win32process


def measure(output):
    source = (Path(__file__).resolve().parents[1] / 'tests/fixtures/vba/project_state.bas').read_text(encoding='utf-8-sig')
    records = []
    for mode in range(7):
        excel = book = component = sheet = process = None
        done = threading.Event()
        row = {'mode': mode}
        try:
            excel = win32com.client.DispatchEx('Excel.Application')
            pid = win32process.GetWindowThreadProcessId(excel.Hwnd)[1]
            process = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_TERMINATE | win32con.SYNCHRONIZE, False, pid)
            def monitor():
                started = time.monotonic()
                while not done.wait(.1):
                    def visit(hwnd, unused):
                        if win32process.GetWindowThreadProcessId(hwnd)[1] != pid or not win32gui.IsWindowVisible(hwnd):
                            return
                        controls = []
                        win32gui.EnumChildWindows(hwnd, lambda child, unused: controls.append((child, win32gui.GetDlgCtrlID(child), win32gui.GetWindowText(child))), None)
                        for child, ident, text in controls:
                            if ident == 4803:
                                row['dialog'] = text
                            if ident == 4800:
                                win32gui.PostMessage(hwnd, win32con.WM_COMMAND, ident, child)
                    win32gui.EnumWindows(visit, None)
                    if time.monotonic() - started > 30:
                        row['timeout'] = True
                        win32api.TerminateProcess(process, 1)
                        return
            threading.Thread(target=monitor, daemon=True).start()
            excel.Visible = False
            excel.DisplayAlerts = False
            book = excel.Workbooks.Add()
            sheet = book.Worksheets(1)
            sheet.Name = 'Sheet1'
            component = book.VBProject.VBComponents.Add(1)
            component.Name = 'StateProbe'
            component.CodeModule.AddFromString(source)
            try:
                row['returned'] = excel.Run('StateProbe.Prime', mode)
            except Exception as failure:
                row['run_error'] = repr(failure)
            row['inside'] = sheet.Range('D1').Value2
            row['after'] = excel.Run('StateProbe.ReadGlobals')
            sheet.Range('A1').Value2 = 9
            row['again'] = excel.Run('StateProbe.ReadGlobals')
            book.Close(False)
            book = component = sheet = None
            excel.Quit()
        finally:
            done.set()
            book = component = sheet = excel = None
            gc.collect()
            if process is not None:
                if win32event.WaitForSingleObject(process, 5000) == win32event.WAIT_TIMEOUT:
                    win32api.TerminateProcess(process, 1)
                    assert win32event.WaitForSingleObject(process, 5000) == win32event.WAIT_OBJECT_0
                win32api.CloseHandle(process)
        assert 'after' in row and 'again' in row and not row.get('timeout'), row
        records.append(row)
        output.write_text(json.dumps(records, ensure_ascii=False, indent=2), encoding='utf-8')
        print(mode, row['after'], row['again'], flush=True)
    return records


def measure_objects(output, scratch):
    """Class/Collection aliases, array/static storage, files, Rnd and lifetime.

    The VBA file lives at a new UUID path inside the supplied scratch directory.
    Existing files and workbooks are never overwritten.
    """
    source = (Path(__file__).resolve().parents[1] / 'tests/fixtures/vba/project_objects.bas').read_text(encoding='utf-8-sig')
    standard, class_source = source.split('VERSION 1.0 CLASS\nAttribute VB_Name = "Counter"\n')
    scratch = scratch.resolve()
    scratch.mkdir(parents=True, exist_ok=True)
    path = scratch / (uuid.uuid4().hex + '.txt')
    assert not path.exists()
    standard = standard.replace('"state.txt"', '"' + str(path) + '"')
    excel = book = component = class_component = process = None
    done = threading.Event()
    result = {}
    try:
        excel = win32com.client.DispatchEx('Excel.Application')
        pid = win32process.GetWindowThreadProcessId(excel.Hwnd)[1]
        process = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_TERMINATE | win32con.SYNCHRONIZE, False, pid)
        def monitor():
            if not done.wait(30):
                result['timeout'] = True
                win32api.TerminateProcess(process, 1)
        threading.Thread(target=monitor, daemon=True).start()
        excel.Visible = False
        excel.DisplayAlerts = False
        book = excel.Workbooks.Add()
        class_component = book.VBProject.VBComponents.Add(2)
        class_component.Name = 'Counter'
        class_component.CodeModule.AddFromString(class_source)
        component = book.VBProject.VBComponents.Add(1)
        component.Name = 'GraphProbe'
        component.CodeModule.AddFromString(standard)
        book.Worksheets(1).Range('A1').Value2 = 5
        excel.Run('GraphProbe.Prime')
        result['first'] = excel.Run('GraphProbe.ReadState')
        book.Worksheets(1).Range('A1').Value2 = 9
        result['second'] = excel.Run('GraphProbe.ReadState')
        result['file'] = excel.Run('GraphProbe.ReadFile')
        result['random'] = [excel.Run('GraphProbe.NextRandom') for _ in range(4)]
        result['release'] = excel.Run('GraphProbe.Release')
        book.Close(False)
        book = component = class_component = None
        excel.Quit()
    finally:
        done.set()
        book = component = class_component = excel = None
        gc.collect()
        if process is not None:
            if win32event.WaitForSingleObject(process, 5000) == win32event.WAIT_TIMEOUT:
                win32api.TerminateProcess(process, 1)
                assert win32event.WaitForSingleObject(process, 5000) == win32event.WAIT_OBJECT_0
            win32api.CloseHandle(process)
    assert not result.get('timeout') and 'release' in result, result
    output.write_text(json.dumps(result, indent=2), encoding='utf-8')
    return result


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output', required=True, type=Path)
    parser.add_argument('--objects', action='store_true', help='Measure object graph instead of termination controls')
    parser.add_argument('--scratch', type=Path, help='Directory for new files created by --objects')
    args = parser.parse_args()
    if args.objects:
        if args.scratch is None:
            parser.error('--objects requires --scratch')
        measure_objects(args.output.resolve(), args.scratch)
    else:
        measure(args.output.resolve())
