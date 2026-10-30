# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.
"""Measure Unicode module/class/document identities and execution in owned Excel.

Requires caller-enabled VBA project access; never changes trust settings.
Only --write-fixtures replaces these explicitly named self-authored packages.
"""
from pathlib import Path
import argparse,ctypes,csv,gc,io,json,re,subprocess,threading,time,tempfile
import pythoncom
import win32api,win32com.client,win32con,win32event,win32gui,win32process


def measure(output,names=None,package_directory=None,write_fixtures=False,fixture_name='unicode_modules.json'):
    if package_directory is None:
        with tempfile.TemporaryDirectory(prefix='oxi-vba-unicode-') as directory:
            return measure(output,names,Path(directory),write_fixtures,fixture_name)
    fixtures=Path(__file__).resolve().parents[1]/'tests/fixtures/vba'
    cases=json.loads((fixtures/fixture_name).read_text(encoding='utf-8'))
    if names is not None:
        cases=[case for case in cases if case['name'] in names]
        assert len(cases)==len(names),'Unknown or duplicate case names'
    package_directory=package_directory.resolve();package_directory.mkdir(parents=True,exist_ok=True)
    result=dict(cases=[],windows_ansi_code_page=ctypes.windll.kernel32.GetACP())
    for case in cases:
        assert case['source_code_page']==result['windows_ansi_code_page'],'Target ANSI code page differs from this oracle host'
        excel = book = component = control = chart = None
        process = None
        done = threading.Event()
        row = dict(name=case['name'], entry=case['procedure'], exports=[])
        try:
            existing_excel_pids={int(fields[1]) for fields in csv.reader(io.StringIO(subprocess.check_output(['tasklist','/FI','IMAGENAME eq EXCEL.EXE','/FO','CSV','/NH'],text=True))) if len(fields)>1 and fields[0].upper()=='EXCEL.EXE'}
            excel = win32com.client.DispatchEx('Excel.Application')
            pid = win32process.GetWindowThreadProcessId(excel.Hwnd)[1]
            if pid in existing_excel_pids:
                row['reused_existing_pid']=pid
                excel=None
                raise RuntimeError('DispatchEx reused an existing Excel process; leave it untouched')
            row['owned_pid'] = pid
            row['measurement_status']='running'
            output.write_text(json.dumps(dict(result,cases=result['cases']+[row]),ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
            process = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_TERMINATE | win32con.SYNCHRONIZE, False, pid)
            def monitor(pid=pid, process=process, done=done, row=row):
                start = time.monotonic()
                while not done.wait(.1):
                    dialogs=[]
                    def visit(hwnd, unused):
                        if win32process.GetWindowThreadProcessId(hwnd)[1]!=pid or not win32gui.IsWindowVisible(hwnd) or win32gui.GetClassName(hwnd)!='#32770': return
                        texts=[]
                        win32gui.EnumChildWindows(hwnd, lambda child, unused: texts.append(win32gui.GetWindowText(child)), None)
                        dialogs.append(dict(title=win32gui.GetWindowText(hwnd), texts=texts))
                    win32gui.EnumWindows(visit, None)
                    if dialogs or time.monotonic()-start>40:
                        row['dialogs']=dialogs;row['timeout']=not bool(dialogs)
                        win32api.TerminateProcess(process,1);return
            threading.Thread(target=monitor, daemon=True).start()
            excel.Visible=False;excel.DisplayAlerts=False
            result['environment']=dict(version=excel.Version, build=excel.Build, ui=excel.LanguageSettings.LanguageID(2))
            book=excel.Workbooks.Add()
            while book.Worksheets.Count>1: book.Worksheets(book.Worksheets.Count).Delete()
            if case['measurement'].get('chart'):
                old_names=[c.Name for c in book.VBProject.VBComponents]
                chart=book.Charts.Add()
                chart.Name='ChartTab'
                chart_name=next(c.Name for c in book.VBProject.VBComponents if c.Name not in old_names)
                chart=None
            for spec in case['measurement']['modules']:
                if spec['kind']==100:
                    old=chart_name if spec['old']=='@chart' else spec['old']
                    component=book.VBProject.VBComponents(old)
                else:
                    component=book.VBProject.VBComponents.Add(spec['kind'])
                    if spec['old']: component.Name=spec['old']
                if spec['kind']==2 and case['measurement'].get('expose_class'):component.Properties('Instancing').Value=2
                component.CodeModule.AddFromString(spec['body'])
                component.Name=spec['name']
                if spec.get('predeclared') or spec.get('member_attributes'):
                    assert spec['kind']==2
                    with tempfile.TemporaryDirectory(prefix='oxi-vba-class-') as temp:
                        exported=Path(temp)/'component.cls'
                        component.Export(str(exported.resolve()))
                        content=exported.read_bytes().decode('cp932')
                        if spec.get('predeclared'):
                            assert content.count('Attribute VB_PredeclaredId = False')==1
                            content=content.replace('Attribute VB_PredeclaredId = False','Attribute VB_PredeclaredId = True')
                        for member,attributes in spec.get('member_attributes',{}).items():
                            lines=[]
                            for attribute,value in attributes.items():
                                if isinstance(value,bool):literal='True' if value else 'False'
                                elif isinstance(value,int):literal=str(value)
                                elif isinstance(value,str):literal='"'+value.replace('"','""')+'"'
                                else:raise ValueError('Unsupported member attribute value')
                                lines.append((attribute,literal))
                            pattern=r'(?im)^((?:Public|Private|Friend)\s+(?:Property\s+Get|Function|Sub)\s+('+re.escape(member)+r')\b[^\r\n]*\r\n)'
                            content,count=re.subn(pattern,lambda match:match.group(1)+''.join('Attribute '+match.group(2)+'.'+attribute+' = '+literal+'\r\n' for attribute,literal in lines),content,count=1)
                            assert count==1,'Missing member declaration for attributes: '+member
                        exported.write_bytes(content.encode('cp932'))
                        book.VBProject.VBComponents.Remove(component)
                        component=book.VBProject.VBComponents.Import(str(exported.resolve()))
                        assert component.Name==spec['name']
                component=None
            package=(package_directory/(case['name']+'.xlsm')).resolve()
            assert package.parent==package_directory
            book.SaveAs(str(package),52)
            row['package']=str(package)
            with tempfile.TemporaryDirectory(prefix='oxi-vba-unicode-') as temp:
                for index in range(1,book.VBProject.VBComponents.Count+1):
                    component=book.VBProject.VBComponents(index)
                    file=Path(temp)/(str(index)+('.bas' if component.Type==1 else '.cls'))
                    component.Export(str(file))
                    row['exports'].append(dict(name=component.Name, kind=component.Type, source=file.read_text(encoding='cp932')))
                    component=None
            component=book.VBProject.VBComponents(case['procedure'].split('.')[0])
            component.CodeModule.CodePane.Show()
            control=excel.VBE.CommandBars.FindControl(1,578)
            assert control is not None
            row['compile_enabled_before']=bool(control.Enabled)
            if control.Enabled: control.Execute();time.sleep(.5)
            row['compile_enabled_after']=bool(control.Enabled)
            control=component=None
            row['steps']=[dict(value=excel.Run("'"+book.Name.replace("'","''")+"'!"+case['procedure'])) for _ in range(3)]
            row['value']=row['steps'][0]['value']
            row['host_context']=dict(book_name=book.Name,book_code_name=book.CodeName,active_sheet=book.ActiveSheet.Index-1,sheets=[dict(name=sheet.Name,code_name=sheet.CodeName,kind='chart' if sheet.Index in [c.Index for c in book.Charts] else 'worksheet') for sheet in book.Sheets])
            row['document_modules']=[]
            for exported in row['exports']:
                if exported['kind']!=100:continue
                if exported['name'] in [sheet.CodeName for sheet in book.Sheets]:
                    sheet=next(sheet for sheet in book.Sheets if sheet.CodeName==exported['name'])
                    row['document_modules'].append(dict(name=exported['name'],baseType='Excel.Chart' if sheet.Index in [c.Index for c in book.Charts] else 'Excel.Worksheet',sheetIndex=sheet.Index-1,codeName=sheet.CodeName))
                    sheet=None
                else:row['document_modules'].append(dict(name=exported['name'],baseType='Excel.Workbook',codeName=book.CodeName))
            book.Save()
            source='\n'.join(entry['source'] for entry in row['exports']).replace('\r\n','\n')
            assert source==(fixtures/case['source']).read_text(encoding='utf-8')
            assert all(step['value']==case['expected'] for step in row['steps'])
            assert sorted(row['document_modules'],key=lambda d:d['name'])==sorted(case['document_modules'],key=lambda d:d['name'])
            if write_fixtures:
                target=(fixtures/case['package_fixture']).resolve()
                assert target.parent==fixtures.resolve()
                book.SaveCopyAs(str(target))
        except Exception as failure:
            row['measurement_error']=repr(failure)
        finally:
            if book is not None and not row.get('timeout') and not row.get('dialogs'):
                try:book.Close(False)
                except Exception as failure:row['close_error']=repr(failure)
            book=component=control=chart=None
            if excel is not None and not row.get('timeout') and not row.get('dialogs'):
                try:excel.VBE.MainWindow.Visible=False
                except Exception as failure:row['hide_vbe_error']=repr(failure)
                try:excel.Quit()
                except Exception as failure:row['quit_error']=repr(failure)
            excel=None;gc.collect()
            if process is not None:
                row['cleanup_grace_ms']=15000
                cleanup_start=time.monotonic()
                while win32event.WaitForSingleObject(process,100)==win32event.WAIT_TIMEOUT and time.monotonic()-cleanup_start<15:
                    pythoncom.PumpWaitingMessages()
                row['cleanup_elapsed_seconds']=time.monotonic()-cleanup_start
                if win32event.WaitForSingleObject(process,0)==win32event.WAIT_TIMEOUT:
                    row['forced_exit']=True;win32api.TerminateProcess(process,1)
                    assert win32event.WaitForSingleObject(process,5000)==win32event.WAIT_OBJECT_0
                row['owned_exit_code']=win32process.GetExitCodeProcess(process)
                done.set();win32api.CloseHandle(process)
            done.set()
        row['measurement_status']='finished'
        result['cases'].append(row)
        output.write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
        assert not row.get('measurement_error') and not row.get('timeout') and not row.get('dialogs'),row
        print(case['name']+' measured',flush=True)

    return result


if __name__=='__main__':
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output',required=True,type=Path)
    parser.add_argument('--names',nargs='+')
    parser.add_argument('--package-dir',type=Path)
    parser.add_argument('--write-fixtures',action='store_true')
    args=parser.parse_args()
    measure(args.output.resolve(),args.names,args.package_dir,args.write_fixtures)
