# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.
"""Compare special_errors.bas with Japanese Excel (Windows, pywin32 required).

Uses a new Excel instance and only self-authored source. Does not change Office
trust settings; VBA project access must already be enabled by the caller.
"""
from pathlib import Path
import argparse
import gc
import gzip
import json
import threading
import time

import win32api
import win32con
import win32process
import win32com.client


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--json-output', type=Path)
    parser.add_argument('--sweep-output', type=Path, help='Also capture both full 0..65535 sweeps in this directory.')
    args = parser.parse_args()
    fixtures = Path(__file__).resolve().parents[1] / 'tests/fixtures/vba'
    source = (fixtures / 'special_errors.bas').read_text(encoding='utf-8')
    cases = [line.split('\t') for line in (fixtures / 'special_errors_ja.tsv').read_text(encoding='utf-8').splitlines()]
    excel = book = component = process = None
    done = threading.Event()
    last_activity = [time.monotonic()]
    timed_out = threading.Event()
    thread = None
    failures = []
    compared = 0
    try:
        excel = win32com.client.DispatchEx('Excel.Application')
        pid = win32process.GetWindowThreadProcessId(excel.Hwnd)[1]
        process = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_TERMINATE, False, pid)

        def watchdog():
            while not done.wait(0.2):
                if time.monotonic() - last_activity[0] > 90:
                    timed_out.set()
                    # This handle belongs solely to the newly created instance.
                    win32api.TerminateProcess(process, 1)
                    return

        last_activity[0] = time.monotonic()
        thread = threading.Thread(target=watchdog, daemon=True)
        thread.start()
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AutomationSecurity = 1
        if excel.LanguageSettings.LanguageID(2) != 1041:
            raise RuntimeError('This fixture requires Japanese Excel UI language 1041.')
        environment = {'version': excel.Version, 'build': excel.Build}
        for procedure, number, mode, outcome, expected_progress, expected in cases:
            last_activity[0] = time.monotonic()
            book = excel.Workbooks.Add()
            try:
                component = book.VBProject.VBComponents.Add(1)
            except Exception as failure:
                raise RuntimeError('VBA project access is required. This script does not change Office trust settings.') from failure
            component.Name = 'SpecialErrors'
            component.CodeModule.AddFromString(source)
            qualified = "'" + book.Name + "'!SpecialErrors."
            call_args = [int(number)]
            if mode != 'omitted':
                call_args.append('' if mode == 'empty' else mode)
            actual_outcome = 'return'
            try:
                actual = str(excel.Run(qualified + procedure, *call_args))
                parts = actual.split('|')
                if len(parts) == 5 and parts[3] and parts[3].lower().endswith('vblr6.chm'):
                    parts[3] = '{VBA_HELP}'
                    actual = '|'.join(parts)
            except Exception as failure:
                info = getattr(failure, 'excepinfo', None)
                if info is None:
                    raise
                actual_outcome = 'terminate'
                actual = str(info[5])
            progress = str(excel.Run(qualified + 'ReadProgress'))
            compared += 1
            if (actual_outcome, progress, actual) != (outcome, expected_progress, expected):
                failures.append({'procedure': procedure, 'number': int(number), 'mode': mode,
                                 'expected': [outcome, expected_progress, expected],
                                 'actual': [actual_outcome, progress, actual]})
            book.Close(False)
            component = book = None
        book = excel.Workbooks.Add()
        component = book.VBProject.VBComponents.Add(1)
        component.Name = 'DispatchProbe'
        component.CodeModule.AddFromString((fixtures / 'err_dispatch.bas').read_text(encoding='utf-8'))
        for line in (fixtures / 'err_dispatch_ja.tsv').read_text(encoding='utf-8').splitlines():
            member, kind, argc, expected = line.split('\t', 3)
            last_activity[0] = time.monotonic()
            actual = str(excel.Run("'" + book.Name + "'!DispatchProbe.DispatchMatrix", member, int(kind), int(argc)))
            actual = '|'.join('{VBA_HELP}' if part.lower().endswith('vblr6.chm') else part for part in actual.split('|'))
            compared += 1
            if actual != expected:
                failures.append({'member': member, 'kind': int(kind), 'argc': int(argc), 'expected': expected, 'actual': actual})
        book.Close(False)
        component = book = None
        if args.sweep_output:
            args.sweep_output.mkdir(parents=True, exist_ok=True)
            book = excel.Workbooks.Add()
            component = book.VBProject.VBComponents.Add(1)
            component.Name = 'SpecialErrors'
            component.CodeModule.AddFromString(source)
            for use_raise in [False, True]:
                last_activity[0] = time.monotonic()
                values = excel.Run("'" + book.Name + "'!SpecialErrors.Sweep", use_raise)
                if len(values) != 65536:
                    raise AssertionError('Incomplete error sweep')
                name = 'raise' if use_raise else 'statement'
                with gzip.open(args.sweep_output / (name + '-sweep.json.gz'), 'wt', encoding='utf-8') as output:
                    json.dump(values, output, ensure_ascii=False)
        result = {**environment, 'total': compared, 'passed': compared - len(failures), 'mismatches': failures}
        if args.json_output:
            args.json_output.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
        print(json.dumps(result, ensure_ascii=True, indent=2))
        if failures:
            raise AssertionError('Excel special error comparisons differ')
    finally:
        try:
            try:
                if book is not None:
                    book.Close(False)
            finally:
                if excel is not None:
                    excel.Quit()
        finally:
            done.set()
            if thread is not None:
                thread.join(2)
            if process is not None:
                win32api.CloseHandle(process)
            component = book = excel = None
            gc.collect()
        if timed_out.is_set():
            raise TimeoutError('The owned Excel probe exceeded 90 seconds without progress')


if __name__ == '__main__':
    main()
