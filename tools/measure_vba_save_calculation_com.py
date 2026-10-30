# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.
"""Measure SaveAs calculation and calculation-setting dirty state in Excel.

Uses only owned DispatchEx instances and caller-enabled VBA project access.
Never changes trust settings. All saved files are new files in --scratch.
"""
from pathlib import Path
import argparse, gc, json, threading, uuid, zipfile
import xml.etree.ElementTree as ET
import win32api, win32com.client, win32con, win32event, win32process


def measure(output, scratch):
    source = (Path(__file__).resolve().parents[1] / 'tests/fixtures/vba/save_calculation.bas').read_text(encoding='utf-8-sig')
    scratch.mkdir(parents=True, exist_ok=True)
    result = {'saves': [], 'modes': [], 'on_save': []}
    jobs = [('save', mode, flag, udf) for mode in (-4105, -4135, 2) for flag in (True, False) for udf in (False, True)]
    jobs += [('mode', initial, target, False) for initial in (-4105, -4135, 2) for target in (-4105, -4135, 2)]
    jobs += [('on_save', initial, target, False) for initial in (True, False) for target in (True, False)]
    for kind, initial, target, udf in jobs:
        excel = book = component = sheet = process = None
        done = threading.Event()
        row = {'kind': kind, 'initial': initial, 'target': target, 'udf': udf}
        try:
            excel = win32com.client.gencache.EnsureDispatch(win32com.client.DispatchEx('Excel.Application'))
            pid = win32process.GetWindowThreadProcessId(excel.Hwnd)[1]
            process = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_TERMINATE | win32con.SYNCHRONIZE, False, pid)
            def watchdog():
                if not done.wait(30):
                    row['timeout'] = True
                    win32api.TerminateProcess(process, 1)
            thread = threading.Thread(target=watchdog, daemon=True)
            thread.start()
            excel.Visible = False
            excel.DisplayAlerts = False
            book = excel.Workbooks.Add()
            sheet = book.Worksheets(1)
            sheet.Name = 'Sheet1'
            prior = (excel.Calculation, excel.CalculateBeforeSave)
            result['environment'] = {'version': excel.Version, 'build': excel.Build, 'ui': excel.LanguageSettings.LanguageID(2)}
            component = book.VBProject.VBComponents.Add(1)
            component.Name = 'SaveProbe'
            component.CodeModule.AddFromString(source)
            excel.Run('SaveProbe.Reset', udf)
            if kind == 'save':
                excel.Run('SaveProbe.Configure', initial, target)
                row['before'] = sheet.Range('B1').Value2
                path = scratch / (uuid.uuid4().hex + '.xlsm')
                assert not path.exists()
                book.SaveAs(str(path), FileFormat=52)
                row['after'] = sheet.Range('B1').Value2
                row['saved'] = bool(book.Saved)
                with zipfile.ZipFile(path) as package:
                    calc = ET.fromstring(package.read('xl/workbook.xml')).find('{*}calcPr')
                    row['calcPr'] = dict(calc.attrib) if calc is not None else {}
                    cell = ET.fromstring(package.read('xl/worksheets/sheet1.xml')).find('.//{*}c[@r="B1"]')
                    row['cached'] = float(cell.find('{*}v').text)
                    row['formula'] = cell.find('{*}f').text
                result['saves'].append(row)
            else:
                procedure = 'ModeOnly' if kind == 'mode' else 'OnSaveOnly'
                excel.Run('SaveProbe.' + procedure, initial)
                book.Saved = True
                excel.Run('SaveProbe.' + procedure, target)
                row['saved'] = bool(book.Saved)
                result['modes' if kind == 'mode' else 'on_save'].append(row)
            excel.Calculation, excel.CalculateBeforeSave = prior
            book.Close(False)
            book = sheet = component = None
            excel.Quit()
        except Exception as failure:
            row['measurement_error'] = repr(failure)
            result.setdefault('failures', []).append(row)
            if excel is not None and not row.get('timeout'):
                try: excel.Quit()
                except Exception: pass
        finally:
            done.set()
            excel = book = component = sheet = None
            gc.collect()
            if process is not None:
                if win32event.WaitForSingleObject(process, 5000) == win32event.WAIT_TIMEOUT:
                    win32api.TerminateProcess(process, 1)
                    assert win32event.WaitForSingleObject(process, 5000) == win32event.WAIT_OBJECT_0
                win32api.CloseHandle(process)
            output.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
        print(kind, initial, target, udf, 'FAILED' if row.get('measurement_error') else 'recorded', flush=True)
    assert not result.get('failures'), result.get('failures')
    return result


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output', required=True, type=Path)
    parser.add_argument('--scratch', required=True, type=Path)
    args = parser.parse_args()
    measure(args.output.resolve(), args.scratch.resolve())
