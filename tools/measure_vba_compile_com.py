# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.
"""Compare procedure compilation and Collection dispatch with Japanese Excel.

Requires Windows, pywin32, and caller-enabled VBA project access. Never changes
Office trust settings. Compilation dialogs are recorded before terminating only
the newly created Excel instance through its owned process handle.
"""
from pathlib import Path
import argparse
import gc
import json
import threading
import time
import uuid

import win32api
import win32con
import win32gui
import win32process
import win32com.client


class ExcelProbe:
    def __init__(self):
        self.excel = self.book = self.component = self.process = None
        self.environment = None

    def close(self, dead=False):
        try:
            if not dead:
                try:
                    if self.book is not None:
                        self.book.Close(False)
                finally:
                    if self.excel is not None:
                        self.excel.Quit()
        finally:
            if self.process is not None:
                win32api.CloseHandle(self.process)
            self.component = self.book = self.excel = self.process = None
            gc.collect()

    def observe(self, source, other='', entry='Probe', args=(), trace=False):
        if self.excel is None:
            self.excel = win32com.client.DispatchEx('Excel.Application')
            self.pid = win32process.GetWindowThreadProcessId(self.excel.Hwnd)[1]
            self.process = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_TERMINATE, False, self.pid)
            self.excel.Visible = False
            self.excel.DisplayAlerts = False
            self.excel.AutomationSecurity = 1
            self.environment = {'version': self.excel.Version, 'build': self.excel.Build, 'ui': self.excel.LanguageSettings.LanguageID(2)}
            if self.environment['ui'] != 1041:
                raise RuntimeError('These measured compiler descriptions require Japanese Excel UI1041.')
        pid, process = self.pid, self.process
        done = threading.Event()
        watch = {}

        def monitor():
            started = time.monotonic()
            while not done.wait(0.1):
                dialogs = []

                def visit(hwnd, unused):
                    try:
                        if win32process.GetWindowThreadProcessId(hwnd)[1] != pid or not win32gui.IsWindowVisible(hwnd):
                            return
                        texts = []
                        win32gui.EnumChildWindows(hwnd, lambda child, unused: texts.append(win32gui.GetWindowText(child)), None)
                        for text in texts:
                            if 'コンパイル エラー:' in text:
                                description = text.replace('\r', '').split('\n')[-1].strip()
                                if description and not description.endswith(':'):
                                    dialogs.append({'title': win32gui.GetWindowText(hwnd), 'description': description, 'text': texts})
                    except win32gui.error:
                        pass  # A completed call may already have closed its window.

                win32gui.EnumWindows(visit, None)
                if dialogs or time.monotonic() - started > 30:
                    watch.update({'dialogs': dialogs, 'timeout': not bool(dialogs)})
                    win32api.TerminateProcess(process, 1)
                    return

        thread = threading.Thread(target=monitor, daemon=True)
        thread.start()
        result = {'pid': pid}
        try:
            self.book = self.excel.Workbooks.Add()
            try:
                self.component = self.book.VBProject.VBComponents.Add(1)
            except Exception as failure:
                raise RuntimeError('Caller-enabled VBA project access is required; this script does not enable it.') from failure
            self.component.Name = 'EntryProbe'
            self.component.CodeModule.AddFromString(source)
            if other:
                self.component = self.book.VBProject.VBComponents.Add(1)
                self.component.Name = 'OtherProbe'
                self.component.CodeModule.AddFromString(other)
            result.update(outcome='return', value=str(self.excel.Run("'" + self.book.Name + "'!EntryProbe." + entry, *args)))
            if trace:
                result['trace'] = str(self.excel.Run("'" + self.book.Name + "'!EntryProbe.ReadTrace"))
        except Exception as failure:
            result.update(outcome='exception', error=repr(failure))
        finally:
            done.set()
            thread.join(2)
        result['watchdog'] = watch
        if watch.get('dialogs'):
            result.update(outcome='compile', value=watch['dialogs'][0]['description'])
        if watch:
            self.close(dead=True)
        else:
            if self.book is not None:
                self.book.Close(False)
            self.component = self.book = None
        return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output-dir', type=Path)
    args = parser.parse_args()
    root = Path(__file__).resolve().parents[1]
    fixtures = root / 'tests/fixtures/vba'
    output = args.output_dir or root / 'pipeline_data' / ('vba-compile-' + uuid.uuid4().hex)
    output = output.resolve()
    output.mkdir(parents=True, exist_ok=False)
    observations, failures = [], []
    probe = ExcelProbe()
    try:
        compiler_cases = (fixtures / 'compile/cases.tsv').read_text(encoding='utf-8').splitlines()
        for line in compiler_cases:
            name, main_file, other_file, expected_outcome, expected, expected_trace = line.split('\t', 5)
            marker = output / (name + '.trace.txt')

            def instrument(source):
                for point in ['ENTRY', 'ARGUMENT']:
                    code = 'Open "' + str(marker).replace('"', '""') + '" For Append As #94\nPrint #94, "' + point.lower() + '"\nClose #94'
                    source = source.replace("' PROBE_" + point + '_MARK', code)
                return source

            source = instrument((fixtures / 'compile' / main_file).read_text(encoding='utf-8'))
            other = instrument((fixtures / 'compile' / other_file).read_text(encoding='utf-8')) if other_file else ''
            actual = probe.observe(source, other, trace=True)
            actual_trace = '|'.join(marker.read_text(encoding='cp932').splitlines()) if marker.exists() else ''
            actual.update(name=name, file_trace=actual_trace)
            observations.append(actual)
            if (actual['outcome'], actual.get('value'), actual_trace) != (expected_outcome, expected, expected_trace) or (actual['outcome'] == 'return' and actual.get('trace') != expected_trace):
                failures.append({'name': name, 'expected': [expected_outcome, expected, expected_trace], 'actual': actual})
            (output / 'observations.json').write_text(json.dumps(observations, ensure_ascii=False, indent=2), encoding='utf-8')
        collection_cases = (fixtures / 'collection_dispatch_ja.tsv').read_text(encoding='utf-8').splitlines()
        source = (fixtures / 'collection_dispatch.bas').read_text(encoding='utf-8')
        for line in collection_cases:
            member, kind, argc, expected = line.split('\t', 3)
            actual = probe.observe(source, entry='DispatchMatrix', args=(member, int(kind), int(argc)))
            value = '|'.join('{VBA_HELP}' if part.lower().endswith('vblr6.chm') else part for part in actual.get('value', '').split('|'))
            if actual['outcome'] != 'return' or value != expected:
                failures.append({'member': member, 'kind': int(kind), 'argc': int(argc), 'expected': expected, 'actual': actual})
        total = len(compiler_cases) + len(collection_cases)
        report = {'environment': probe.environment, 'compiler_cases': len(compiler_cases), 'collection_cases': len(collection_cases), 'total': total, 'passed': total - len(failures), 'failures': failures}
        (output / 'results.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
        print(json.dumps(report, ensure_ascii=True, indent=2))
        if failures:
            raise AssertionError('Excel compiler/dispatch measurements differ')
    finally:
        probe.close()


if __name__ == '__main__':
    main()
