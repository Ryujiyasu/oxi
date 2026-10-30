# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.
"""Measure VBA editing operations and subsequent project storage in Excel.

Uses owned DispatchEx instances and caller-enabled project access. Never
changes trust settings, writes VBA files or saves workbooks.
"""
from pathlib import Path
import argparse, gc, json, threading
import win32api, win32com.client, win32con, win32event, win32process


def measure(output):
    fixtures = Path(__file__).resolve().parents[1] / 'tests/fixtures/vba'
    result = {}
    for corpus in ('standard', 'classes', 'controls'):
        source = (fixtures / ('source_edit_classes.bas' if corpus == 'classes' else 'source_edit_standard.bas')).read_text(encoding='utf-8-sig')
        cases = json.loads((fixtures / f'source_edit_{corpus}.json').read_text(encoding='utf-8'))
        result[corpus] = []
        for case in cases:
            excel = book = component = class_component = extra_component = code = process = None
            done = threading.Event()
            row = {'name': case['name']}
            try:
                excel = win32com.client.DispatchEx('Excel.Application')
                process = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_TERMINATE | win32con.SYNCHRONIZE, False, win32process.GetWindowThreadProcessId(excel.Hwnd)[1])
                def monitor():
                    if not done.wait(30):
                        row['timeout'] = True
                        win32api.TerminateProcess(process, 1)
                threading.Thread(target=monitor, daemon=True).start()
                excel.Visible = False
                excel.DisplayAlerts = False
                result['environment'] = {'version': excel.Version, 'build': excel.Build, 'ui': excel.LanguageSettings.LanguageID(2)}
                book = excel.Workbooks.Add()
                book.Worksheets(1).Name = 'Sheet1'
                component = book.VBProject.VBComponents.Add(1)
                component.Name = 'EditProbe'
                if corpus == 'classes':
                    standard, class_source = source.split('VERSION 1.0 CLASS\nAttribute VB_Name = "StateClass"\n')
                    component.CodeModule.AddFromString(standard)
                    class_component = book.VBProject.VBComponents.Add(2)
                    class_component.Name = 'StateClass'
                    code = class_component.CodeModule
                    code.AddFromString(class_source)
                    row['prime'] = excel.Run('EditProbe.Prime')
                    read = 'EditProbe.ReadState'
                else:
                    code = component.CodeModule
                    code.AddFromString(source)
                    row['prime'] = excel.Run('EditProbe.Prime', 0)
                    read = 'EditProbe.ReadGlobals'
                row['before'] = excel.Run(read)
                if case.get('new_component'):
                    extra_component = book.VBProject.VBComponents.Add(2)
                    extra_component.Name = 'ExtraClass'
                elif 'append' in case:
                    code.InsertLines(code.CountOfLines + 1, case['append'])
                elif 'insert' in case:
                    for number in range(1, code.CountOfLines + 1):
                        if case['insert_after'] in code.Lines(number, 1):
                            code.InsertLines(number + 1, case['insert'])
                            break
                    else:
                        raise AssertionError(case['name'])
                else:
                    lines = []
                    for number in range(1, code.CountOfLines + 1):
                        text = code.Lines(number, 1)
                        if case['old'] in text:
                            lines.append((number, text.replace(case['old'], case['new'])))
                            if not case.get('replace_all'):
                                break
                    assert lines, case['name']
                    for number, text in reversed(lines):
                        code.ReplaceLine(number, text)
                book.Worksheets(1).Range('A1').Value2 = 9
                row['after'] = excel.Run(read)
                row['actual_module_source'] = code.Lines(1, code.CountOfLines)
            except Exception as failure:
                row['measurement_error'] = repr(failure)
            finally:
                if book is not None and not row.get('timeout'):
                    try: book.Close(False)
                    except Exception: pass
                book = component = class_component = extra_component = code = None
                if excel is not None and not row.get('timeout'):
                    try: excel.Quit()
                    except Exception: pass
                excel = None
                gc.collect()
                if process is not None:
                    if win32event.WaitForSingleObject(process, 5000) == win32event.WAIT_TIMEOUT:
                        row['forced_exit'] = True
                        win32api.TerminateProcess(process, 1)
                        assert win32event.WaitForSingleObject(process, 5000) == win32event.WAIT_OBJECT_0
                    win32api.CloseHandle(process)
                done.set()
            result[corpus].append(row)
            output.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
            assert not row.get('timeout') and not row.get('measurement_error'), row
            assert all(row[key] == case[key] for key in ('prime', 'before', 'after')), row
            print(corpus, case['name'], row['after'], flush=True)
    return result


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output', required=True, type=Path)
    measure(parser.parse_args().output.resolve())
