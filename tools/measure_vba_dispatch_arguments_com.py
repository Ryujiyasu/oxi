# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.
"""Measure typed host references, Err argument coercion, call syntax and interface pointers in Excel.

Uses owned DispatchEx instances and caller-enabled VBA project access. Never
changes trust settings, opens user files or saves workbooks.
"""
from pathlib import Path
import argparse, gc, json, threading, re, time, win32gui
import win32api, win32com.client, win32con, win32event, win32process


def measure(output):
    fixtures = Path(__file__).resolve().parents[1] / 'tests/fixtures/vba'
    cases = json.loads((fixtures / 'dispatch_arguments.json').read_text(encoding='utf-8'))
    result = {'cases': []}
    for case in cases:
        excel = book = component = process = None
        components = []
        done = threading.Event()
        row = {'name': case['name'], 'source': case['source'], 'steps': []}
        try:
            excel = win32com.client.DispatchEx('Excel.Application')
            process = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_TERMINATE | win32con.SYNCHRONIZE, False, win32process.GetWindowThreadProcessId(excel.Hwnd)[1])
            pid=win32process.GetWindowThreadProcessId(excel.Hwnd)[1]
            row['owned_pid'] = pid
            def monitor(pid=pid,process=process,done=done,row=row):
                start=time.monotonic()
                while not done.wait(.1):
                    dialogs=[]
                    def visit(hwnd,unused):
                        if win32process.GetWindowThreadProcessId(hwnd)[1]!=pid or not win32gui.IsWindowVisible(hwnd) or win32gui.GetClassName(hwnd)!='#32770':return
                        texts=[]
                        win32gui.EnumChildWindows(hwnd,lambda child,unused:texts.append(win32gui.GetWindowText(child)),None)
                        if any('コンパイル' in text or 'compile error' in text.lower() for text in texts): dialogs.append(dict(title=win32gui.GetWindowText(hwnd),texts=texts))
                    win32gui.EnumWindows(visit,None)
                    if dialogs or time.monotonic()-start>25:
                        row['dialogs']=dialogs
                        row['timeout']=not bool(dialogs)
                        win32api.TerminateProcess(process,1)
                        return
            threading.Thread(target=monitor, daemon=True).start()
            excel.Visible = False
            excel.DisplayAlerts = False
            result['environment'] = {'version': excel.Version, 'build': excel.Build, 'ui': excel.LanguageSettings.LanguageID(2)}
            book = excel.Workbooks.Add()
            source = (fixtures / case['source']).read_text(encoding='utf-8-sig')
            # Exported module boundaries, preserving standard/class identity.
            blocks = re.split(r'(?=^VERSION 1\.0 CLASS$|^Attribute VB_Name =)', source, flags=re.M)
            class_next = False
            for block in blocks:
                if block.startswith('VERSION 1.0 CLASS'):
                    class_next = True
                    continue
                if not block.startswith('Attribute VB_Name ='):
                    continue
                attribute, code = block.split('\n', 1)
                component = book.VBProject.VBComponents.Add(2 if class_next else 1)
                component.Name = attribute.split('"')[1]
                if '\nAttribute ' in code:
                    book.VBProject.VBComponents.Remove(component)
                    temp=output.parent / ('import_' + case['name'] + ('.cls' if class_next else '.bas'))
                    exported=('VERSION 1.0 CLASS\nBEGIN\n  MultiUse = -1\nEND\n' if class_next else '')+attribute+'\n'+code
                    temp.write_text(exported,encoding='cp932')
                    component=book.VBProject.VBComponents.Import(str(temp))
                else:
                    component.CodeModule.AddFromString(code)
                components.append(component)
                class_next = False
            for step in [dict(procedure=case['procedure'])]*2:
                procedure = step['procedure']
                row['steps'].append({'procedure': procedure, 'value': excel.Run(procedure)})
        except Exception as failure:
            row['measurement_error'] = repr(failure)
        finally:
            if book is not None and not row.get('timeout') and not row.get('dialogs'):
                try: book.Close(False)
                except Exception: pass
            book = component = None
            components.clear()
            if excel is not None and not row.get('timeout') and not row.get('dialogs'):
                try: excel.Quit()
                except Exception: pass
            excel = None
            gc.collect()
            if process is not None:
                if win32event.WaitForSingleObject(process, 5000) == win32event.WAIT_TIMEOUT:
                    row['forced_exit'] = True
                    win32api.TerminateProcess(process, 1)
                    assert win32event.WaitForSingleObject(process, 5000) == win32event.WAIT_OBJECT_0
                row['owned_exit_code'] = win32process.GetExitCodeProcess(process)
                done.set()
                win32api.CloseHandle(process)
            done.set()
        result['cases'].append(row)
        output.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
        assert not row.get('timeout'), row
        if row.get('dialogs'):
            row['compile_message']=row['dialogs'][0]['texts'][-1].split('\n\n')[-1]
        else:
            assert not row.get('measurement_error') and len(row['steps'])==2,row

        output.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
        if case['outcome']=='compile': assert row.get('compile_message')==case['message'],row
        else: assert not row.get('dialogs') and row['steps']==case['steps'],row
        print(case['name'], 'compile' if row.get('dialogs') else 'values', flush=True)
    return result


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output', required=True, type=Path)
    measure(parser.parse_args().output.resolve())
