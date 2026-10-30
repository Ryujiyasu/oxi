# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.
"""Measure document-module compilation and persistent execution in owned Excel.

Requires caller-enabled VBA project access. Creates and closes its own workbooks;
never changes trust settings. Only explicitly named package_fixture cases save
new conformance workbooks under tests/fixtures/vba. Run outputs use qualified
macros in the owned workbook and retain native host context for reproducibility.
"""
from pathlib import Path
import argparse, gc, json, threading, re, time, tempfile, win32gui
import win32api, win32com.client, win32con, win32event, win32process


def measure(output, names=None, package_directory=None):
    if package_directory is None:
        with tempfile.TemporaryDirectory(prefix="oxi-vba-document-identity-") as directory:
            return measure(output,names,Path(directory))
    fixtures = Path(__file__).resolve().parents[1] / 'tests/fixtures/vba'
    cases = json.loads((fixtures / 'document_identity.json').read_text(encoding='utf-8'))
    if names is not None:
        cases=[case for case in cases if case['name'] in names]
        assert len(cases)==len(names), 'Unknown or duplicate case names'
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
            while book.Worksheets.Count < case.get('setup_sheets',1):
                book.Worksheets.Add(After=book.Worksheets(book.Worksheets.Count))
            if 'references' not in result:
                result['references']=[dict(name=reference.Name,guid=reference.GUID,major=reference.Major,minor=reference.Minor) for reference in book.VBProject.References]
            source = (fixtures / case.get('measurement_source',case['source'])).read_text(encoding='utf-8-sig')
            if case.get('edit'):
                source=source.replace(case['edit']['after'],case['edit']['before'],1)
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
                    temp=fixtures / ('import_' + case['name'] + ('.cls' if class_next else '.bas'))
                    exported=('VERSION 1.0 CLASS\nBEGIN\n  MultiUse = -1\nEND\n' if class_next else '')+attribute+'\n'+code
                    temp.write_text(exported,encoding='cp932')
                    component=book.VBProject.VBComponents.Import(str(temp))
                else:
                    component.CodeModule.AddFromString(code)
                components.append(component)
                class_next = False
            for specification in case['documents']:
                owner = book.Worksheets(specification['sheet']+1) if specification['kind']=='sheet' else book
                document = book.VBProject.VBComponents.Item(owner.CodeName)
                document.Name = specification['name']
                document.CodeModule.AddFromString((fixtures/specification['source']).read_text(encoding='utf-8'))
                components.append(document)
            owner=None
            packages=package_directory
            book.SaveAs(str((packages/(case['name']+'.xlsm')).resolve()),52)
            if case.get('package_fixture'):
                book.SaveCopyAs(str((fixtures/case['package_fixture']).resolve()))
            row['document_exports']=[]
            with tempfile.TemporaryDirectory(prefix='oxi-vba-doc-research-') as directory:
                for component in components:
                    exported=Path(directory)/(component.Name+('.bas' if component.Type==1 else '.cls'))
                    component.Export(str(exported))
                    row['document_exports'].append(dict(name=component.Name,component_type=component.Type,source=exported.read_text(encoding='cp932')))
            row['measurement_mode']='VBE whole-project compilation with explicit active code pane'
            if case.get('edit'):
                row['warmup_steps']=[dict(procedure=case['procedure'],value=excel.Run(case['procedure']))]
                edit=case['edit'];target=next(item for item in components if item.Name==edit['module'])
                lines=target.CodeModule.Lines(1,target.CodeModule.CountOfLines).splitlines()
                line=next(index+1 for index,text in enumerate(lines) if text.strip()==edit['before'])
                row['edit_line']=line;target.CodeModule.ReplaceLine(line,edit['after'])
                row['source_after_edit']=target.CodeModule.Lines(1,target.CodeModule.CountOfLines)
            if case.get('compile_project'):
                focus=next(item for item in components if item.Name==case.get('focus_module','DriverModule'))
                focus.CodeModule.CodePane.Show()
                row['active_code_module']=excel.VBE.ActiveCodePane.CodeModule.Parent.Name
                row['active_project_name']=excel.VBE.ActiveVBProject.Name
                control=excel.VBE.CommandBars.FindControl(1,578)
                row['compile_control']=None if control is None else dict(caption=control.Caption,id=control.Id,type=control.Type,enabled=control.Enabled)
                if control is None or not ('コンパイル' in control.Caption or 'compile' in control.Caption.lower()):
                    raise RuntimeError('Compile control identity not established')
                if control.Enabled:
                    row['compile_requested']=True;control.Execute();time.sleep(.5)
                    row['compile_enabled_after']=control.Enabled
                control=None
            row['host_context']=dict(book_name=book.Name,book_code_name=book.CodeName,active_sheet=book.ActiveSheet.Index-1,sheets=[dict(name=sheet.Name,code_name=sheet.CodeName) for sheet in book.Worksheets])
            def metadata():
                rows=[]
                for specification in case['documents']:
                    if specification['kind']=='sheet':
                        sheet=next(sheet for sheet in book.Worksheets if sheet.CodeName==specification['name'])
                        rows.append(dict(name=specification['name'],baseType='Excel.Worksheet',sheetIndex=sheet.Index-1,codeName=sheet.CodeName))
                    else:rows.append(dict(name=specification['name'],baseType='Excel.Workbook',codeName=book.CodeName))
                return rows
            row['metadata_before']=metadata()
            for step_index in range(3):
                if step_index==1:
                    if case['after_first']['move']:
                        sheet=next(sheet for sheet in book.Worksheets if sheet.CodeName=='Ledger')
                        sheet.Move(After=book.Worksheets(book.Worksheets.Count));sheet=None
                    if case['after_first']['save']:
                        book.Save()
                        row['refreshed_package']=str((fixtures/case['package_refreshed'] if case.get('package_refreshed') else packages/(case['name']+'_refreshed.xlsm')).resolve())
                        book.SaveCopyAs(row['refreshed_package'])
                row['steps'].append(dict(procedure=case['procedure'],metadata=metadata(),value=excel.Run("'"+book.Name.replace("'","''")+"'!"+case['procedure'])))
                row['steps'][-1]['sheet_names_after']=[sheet.Name for sheet in book.Worksheets]
        except Exception as failure:
            row['measurement_error'] = repr(failure)
        finally:
            if book is not None and not row.get('timeout') and not row.get('dialogs'):
                try: book.Close(False)
                except Exception: pass
            book = component = document = focus = owner = None
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
            assert not row.get('measurement_error') and len(row['steps'])==3,row

        output.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
        assert [step['value'] for step in row['steps']]==[step['value'] for step in case['steps']],row
        assert row['metadata_before']==case['metadata_before'],row
        assert [step['metadata'] for step in row['steps']]==[step['metadata'] for step in case['steps']],row
        print(case['name'], 'compile' if row.get('dialogs') else 'values', flush=True)
    return result


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output', required=True, type=Path)
    parser.add_argument('--names', nargs='+', help='Measure only the named conformance cases')
    args=parser.parse_args()
    measure(args.output.resolve(), args.names)
