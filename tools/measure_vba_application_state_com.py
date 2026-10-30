# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.
"""Measure application settings across successive Excel VBA invocations.

Requires Windows, pywin32 and caller-enabled project access. Never changes
Office trust settings. Own DispatchEx instances are monitored through held
process handles; runtime error End commands are scoped to those instances.
"""
import argparse
import gc
import json
from pathlib import Path
import threading
import time

import win32api
import win32com.client
import win32con
import win32event
import win32gui
import win32process


def measure(output):
    source = (Path(__file__).resolve().parents[1] / 'tests/fixtures/vba/application_state.bas').read_text(encoding='utf-8-sig')
    results = {'cases': []}
    for mode in range(6):
        excel = book = process = component = None
        done = threading.Event()
        record = {'mode': mode}
        results['cases'].append(record)
        thread = None
        try:
            excel = win32com.client.gencache.EnsureDispatch(win32com.client.DispatchEx('Excel.Application'))
            pid = win32process.GetWindowThreadProcessId(excel.Hwnd)[1]
            process = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_TERMINATE | win32con.SYNCHRONIZE, False, pid)
            results['environment'] = {'version': excel.Version, 'build': excel.Build, 'ui': excel.LanguageSettings.LanguageID(2)}
            excel.Visible = False
            excel.DisplayAlerts = False
            book = excel.Workbooks.Add()
            component = book.VBProject.VBComponents.Add(1)
            component.Name = 'ApplicationProbe'
            component.CodeModule.AddFromString(source)

            def monitor():
                started = time.monotonic()
                clicked = set()
                while not done.wait(.1):
                    def visit(hwnd, unused):
                        try:
                            if win32process.GetWindowThreadProcessId(hwnd)[1] != pid or not win32gui.IsWindowVisible(hwnd): return
                            controls = []
                            win32gui.EnumChildWindows(hwnd, lambda child, unused: controls.append((child, win32gui.GetDlgCtrlID(child), win32gui.GetWindowText(child))), None)
                            ends = [item for item in controls if item[1] == 4800]
                            messages = [text for _, ident, text in controls if ident == 4803]
                            if ends and messages and hwnd not in clicked:
                                clicked.add(hwnd)
                                record.setdefault('dialogs', []).append(messages)
                                child, ident, _ = ends[0]
                                win32gui.PostMessage(hwnd, win32con.WM_COMMAND, ident, child)
                        except win32gui.error: pass
                    win32gui.EnumWindows(visit, None)
                    if time.monotonic() - started > 40:
                        record['timeout'] = True
                        win32api.TerminateProcess(process, 1)
                        return
            thread = threading.Thread(target=monitor, daemon=True)
            thread.start()
            excel.Run('ApplicationProbe.ResetState')
            book.Saved = True
            try:
                record['returned'] = excel.Run('ApplicationProbe.Configure', mode)
            except Exception as failure:
                record['run_error'] = repr(failure)
            record['inside'] = book.Worksheets(1).Range('D1').Value2
            record['after_com'] = {name: getattr(excel, name) for name in ('Calculation', 'ScreenUpdating', 'EnableEvents', 'DisplayAlerts', 'StatusBar', 'Cursor', 'Interactive')}
            record['after'] = excel.Run('ApplicationProbe.ReadState')
            record['after_write'] = excel.Run('ApplicationProbe.WriteMore')
            record['automatic'] = excel.Run('ApplicationProbe.Automatic')
            record['reset_text'] = excel.Run('ApplicationProbe.ResetText')
            record['transitions'] = []
            for initial in (-4105, -4135, 2):
                for target in (-4105, -4135, 2):
                    excel.Run('ApplicationProbe.ResetState')
                    excel.Run('ApplicationProbe.SwitchMode', initial)
                    before = excel.Run('ApplicationProbe.ChangeInput', 7)
                    after = excel.Run('ApplicationProbe.SwitchMode', target)
                    following = excel.Run('ApplicationProbe.ChangeInput', 9)
                    record['transitions'].append({'initial': initial, 'target': target, 'before': before, 'after': after, 'following': following})
            print(mode, record['inside'], '->', record['after'], '->', record['automatic'], flush=True)
        except Exception as failure:
            record['measurement_error'] = repr(failure)
        finally:
            try:
                if not record.get('timeout') and excel is not None:
                    excel.Interactive = True
                    excel.EnableEvents = True
                    excel.ScreenUpdating = True
                    excel.DisplayAlerts = False
                    excel.Cursor = -4143
                    try:
                        if book is not None: book.Close(False)
                    finally: excel.Quit()
            finally:
                done.set()
                if thread: thread.join(2)
                component = book = excel = None
                gc.collect()
                if process is not None:
                    # Quit can return before the process exits. Finish owned
                    # cleanup before a caller restores temporary trust access.
                    if win32event.WaitForSingleObject(process, 5000) == win32event.WAIT_TIMEOUT:
                        record['forced_cleanup'] = True
                        win32api.TerminateProcess(process, 1)
                        assert win32event.WaitForSingleObject(process, 5000) == win32event.WAIT_OBJECT_0
                    win32api.CloseHandle(process)
                    process = None
        output.write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding='utf-8')
    assert len(results['cases']) == 6 and all('reset_text' in row and 'measurement_error' not in row for row in results['cases'])
    return results


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output', type=Path, required=True)
    measure(parser.parse_args().output.resolve())
