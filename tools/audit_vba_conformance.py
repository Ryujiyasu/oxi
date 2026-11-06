# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.
"""Build an evidence-based VBA audit; declarations never imply conformance.

Outputs are local research data. Run from the repository root:
    python tools/audit_vba_conformance.py --native

Run a reviewed manifest containing existing String-report VBA functions:
    python tools/audit_vba_conformance.py --capture RECEIPT.json --pkg PKG
    python tools/audit_vba_conformance.py --browser RECEIPT.json --pkg PKG

The unit is (library, interface, member, invocation kind). Interface aliases
are retained explicitly; totals are not percentages of the entire language.
Native hidden/restricted members remain visible. Only IUnknown/IDispatch
plumbing is excluded, by name, with the exclusion recorded.
"""

import argparse
import collections
import datetime
import hashlib
import json
from pathlib import Path
import re
import subprocess


GATES = ("same_input", "value", "vba_type", "error", "state", "persistence", "browser")
PLUMBING = {"QueryInterface", "AddRef", "Release", "GetTypeInfoCount",
            "GetTypeInfo", "GetIDsOfNames", "Invoke"}
INVOKE = {1: "call", 2: "get", 4: "let", 8: "set"}

# The manifest pins fixtures and the compiled artifact; no VBA project access or
# macro security changes are needed. Existing embedded code runs in Excel.
CAPTURE_NODE = r"""
import fs from 'node:fs';import {pathToFileURL} from 'node:url';import {createHash} from 'node:crypto';
const input=JSON.parse(fs.readFileSync(0,'utf8')),hash=b=>createHash('sha256').update(b).digest('hex');
for(const [name,expected] of Object.entries(input.artifact_hashes)){
 if(hash(fs.readFileSync(input.pkg+'/'+name))!==expected)throw Error('artifact changed: '+name);
}
const w=await import(pathToFileURL(input.pkg+'/oxidocs_wasm.js'));
w.initSync({module:fs.readFileSync(input.pkg+'/oxidocs_wasm_bg.wasm')});
const output={sources:{},cases:[]};
for(const [fixture,meta] of Object.entries(input.fixtures)){
 const bytes=fs.readFileSync(input.root+'/'+meta.path),source=w.spreadsheet_vba_source(bytes);
 if(hash(bytes)!==meta.sha256||hash(source)!==meta.source_sha256)throw Error('fixture changed: '+fixture);
 output.sources[fixture]={source,source_sha256:hash(source),fixture_sha256:hash(bytes),projectName:w.spreadsheet_vba_project_name(bytes)};
 for(const c of input.cases.filter(c=>c.fixture===fixture)){
  try{
   const options={...input.options,projectName:output.sources[fixture].projectName,documentModules:w.spreadsheet_vba_document_modules(bytes)};
   const r=w.run_spreadsheet_vba_with_outcome(w.parse_spreadsheet(bytes),source,c.procedure,c.args,0,meta.path.split('/').at(-1),options);
   const saved=w.edit_xlsx_from_workbook(bytes,r.workbook),reopened=w.parse_spreadsheet(saved);
   const cells=book=>book.sheets.map(s=>({name:s.name,rows:s.rows.map(row=>({index:row.index,cells:row.cells.map(cell=>({col:cell.col,value:cell.value,formula:cell.formula}))}))}));
   const before=JSON.stringify(cells(r.workbook)),after=JSON.stringify(cells(reopened));
   output.cases.push({...c,result:r.result,error:r.error||null,application_state:r.application_state,workbook_changed:r.workbook_changed,
    oxi_roundtrip_equal:before===after,oxi_before_sha256:hash(before),oxi_reopened_sha256:hash(after),
    saved_sha256:hash(saved),saved_base64:Buffer.from(saved).toString('base64')});
  }catch(error){output.cases.push({...c,harness_error:String(error)})}
 }
}
console.log(JSON.stringify(output));
"""


BROWSER_NODE = r"""

import fs from 'node:fs';import path from 'node:path';import http from 'node:http';import {spawn} from 'node:child_process';import {once} from 'node:events';import {createHash,randomUUID} from 'node:crypto';
const [root,pkg,receiptPath]=process.argv.slice(1),receipt=JSON.parse(fs.readFileSync(receiptPath,'utf8'));
const hash=b=>createHash('sha256').update(b).digest('hex'),cache=path.dirname(pkg),profile=path.join(cache,'owned-conformance-browser-'+randomUUID());
const cases=receipt.manifest?.cases||receipt.case_input,options=receipt.manifest?.options||receipt.node_options;
const artifactHashes=receipt.manifest?.artifact_hashes||Object.fromEntries(['oxidocs_wasm.js','oxidocs_wasm_bg.wasm'].map(n=>[n,receipt.oxi_artifact.artifacts[n].sha256]));
const fixtures=receipt.manifest?.fixtures||Object.fromEntries(Object.entries(receipt.sources).map(([f,s])=>[f,{path:'tests/fixtures/vba/'+f+'.xlsm',sha256:s.fixture_sha256,source_sha256:s.source_sha256}]));
for(const [n,s]of Object.entries(artifactHashes))if(hash(fs.readFileSync(path.join(pkg,n)))!==s)throw Error('artifact changed '+n);
const driver=String.raw`
export async function run(useEditor=false){
 const start=performance.now(),input=await(await fetch(location.origin+'/request.json')).json();
 const hash=async b=>Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256',typeof b==='string'?new TextEncoder().encode(b):b)),x=>x.toString(16).padStart(2,'0')).join('');
 const js=await(await fetch(location.origin+'/pkg/oxidocs_wasm.js')).arrayBuffer(),wasm=await(await fetch(location.origin+'/pkg/oxidocs_wasm_bg.wasm')).arrayBuffer();
 if(await hash(js)!==input.artifactHashes['oxidocs_wasm.js']||await hash(wasm)!==input.artifactHashes['oxidocs_wasm_bg.wasm'])throw Error('browser artifact mismatch');
 const w=await import(location.origin+'/pkg/oxidocs_wasm.js');await w.default({module_or_path:new Uint8Array(wasm)});
 const runner=useEditor?await import(location.origin+'/web/vba-runner.js'):null;
 let result={userAgent:navigator.userAgent,mode:useEditor?'editor_worker':typeof document==='undefined'?'worker':'page',cases:[],saved_packages:{}};
 for(const [fixture,meta]of Object.entries(input.fixtures)){
  const bytes=new Uint8Array(await(await fetch(location.origin+'/fixtures/'+fixture)).arrayBuffer()),source=w.spreadsheet_vba_source(bytes);
  if(await hash(bytes)!==meta.sha256||await hash(source)!==meta.source_sha256)throw Error('browser fixture mismatch '+fixture);
  for(const c of input.cases.filter(c=>c.fixture===fixture)){
   try{
    const opts={...input.options,projectName:w.spreadsheet_vba_project_name(bytes),documentModules:w.spreadsheet_vba_document_modules(bytes)};
    const r=runner?await runner.runVbaInWorker(w.parse_spreadsheet(bytes),source,c.procedure,c.args,0,meta.path.split('/').at(-1),opts,bytes):w.run_spreadsheet_vba_with_outcome(w.parse_spreadsheet(bytes),source,c.procedure,c.args,0,meta.path.split('/').at(-1),opts);
    const saved=w.edit_xlsx_from_workbook(bytes,r.workbook),reopened=w.parse_spreadsheet(saved);
    const cells=book=>book.sheets.map(s=>({name:s.name,rows:s.rows.map(row=>({index:row.index,cells:row.cells.map(cell=>({col:cell.col,value:cell.value,formula:cell.formula}))}))}));
    const before=JSON.stringify(cells(r.workbook)),after=JSON.stringify(cells(reopened)),savedHash=await hash(saved);
    if(savedHash!==input.expectedSaved[c.id]&&!result.saved_packages[savedHash]){
     let binary='';for(let offset=0;offset<saved.length;offset+=32768)binary+=String.fromCharCode(...saved.subarray(offset,offset+32768));result.saved_packages[savedHash]=btoa(binary);
    }
    result.cases.push({id:c.id,result:r.result,error:r.error||null,application_state:r.application_state,workbook_changed:r.workbook_changed,
     oxi_roundtrip_equal:before===after,oxi_before_sha256:await hash(before),oxi_reopened_sha256:await hash(after),saved_sha256:savedHash});
   }catch(e){result.cases.push({id:c.id,harness_error:String(e)})}
  }
  runner?.closeVbaProject();
 }
 result.seconds=(performance.now()-start)/1000;return result;
}
`;
const input=JSON.stringify({cases,fixtures,options,artifactHashes,expectedSaved:Object.fromEntries(receipt.cases.map(c=>[c.id,c.saved_sha256]))});let child,edgeExit=false,server,ws,browserWs,diagnostics={},error;
const sockets=new Set();
function cdp(socket){let id=0,waiting=new Map();socket.addEventListener('message',event=>{let m=JSON.parse(event.data);if(waiting.has(m.id)){let p=waiting.get(m.id);waiting.delete(m.id);clearTimeout(p.timer);m.error?p.reject(Error(JSON.stringify(m.error))):p.resolve(m.result)}});return(method,params={})=>new Promise((resolve,reject)=>{let n=++id;const timer=setTimeout(()=>{waiting.delete(n);reject(Error('CDP timeout '+method))},180000);waiting.set(n,{resolve,reject,timer});socket.send(JSON.stringify({id:n,method,params}))})}
try{
 fs.mkdirSync(profile);server=http.createServer((req,res)=>{const p=new URL(req.url,'http://localhost').pathname;res.setHeader('Cache-Control','no-store');
  if(p==='/request.json'){res.setHeader('Content-Type','application/json');return res.end(input)}
  if(p==='/run.mjs'){res.setHeader('Content-Type','text/javascript');return res.end(driver)}
  if(p==='/'){res.setHeader('Content-Type','text/html');return res.end('<!doctype html><meta charset="utf-8"><title>Owned VBA conformance capture</title>')}
  if(p==='/overview'){res.setHeader('Content-Type','text/html');return res.end(fs.readFileSync(path.join(root,'pipeline_data/vba_overview_20261027.html')))}
  let file,mime;
  if(p.startsWith('/fixtures/')&&fixtures[p.slice(10)]){file=path.join(root,fixtures[p.slice(10)].path);mime='application/octet-stream'}
  if(p==='/pkg/oxidocs_wasm.js'||p==='/web/oxidocs_wasm.js'){file=path.join(pkg,'oxidocs_wasm.js');mime='text/javascript'}
  if(p==='/pkg/oxidocs_wasm_bg.wasm'||p==='/web/oxidocs_wasm_bg.wasm'){file=path.join(pkg,'oxidocs_wasm_bg.wasm');mime='application/wasm'}
  if(p==='/web/vba-runner.js'||p==='/web/vba-worker.js'){file=path.join(root,p.slice(1));mime='text/javascript'}
  if(!file){res.statusCode=404;return res.end('not found')}
  res.setHeader('Content-Type',mime);fs.createReadStream(file).pipe(res);
 });server.on('connection',s=>{sockets.add(s);s.on('close',()=>sockets.delete(s))});await new Promise(resolve=>server.listen(0,'127.0.0.1',resolve));
 const origin='http://127.0.0.1:'+server.address().port;
 const edge=path.join(process.env['ProgramFiles(x86)'],'Microsoft/Edge/Application/msedge.exe');
 child=spawn(edge,['--headless=new','--remote-debugging-port=0','--user-data-dir='+profile,'--no-first-run','--no-default-browser-check','--disable-extensions','--disable-background-networking','--disable-sync','--disable-component-update','about:blank'],{windowsHide:true,stdio:'ignore'});child.on('exit',()=>{edgeExit=true});
 diagnostics.owned_pid=child.pid;
 const portfile=path.join(profile,'DevToolsActivePort');
 for(let i=0;!fs.existsSync(portfile)&&i<200;i++){if(edgeExit)throw Error('Edge exited before CDP');await new Promise(r=>setTimeout(r,100))}
 const port=Number(fs.readFileSync(portfile,'utf8').split('\n')[0]),version=await(await fetch('http://127.0.0.1:'+port+'/json/version')).json();
 diagnostics.version=version;const targets=await(await fetch('http://127.0.0.1:'+port+'/json/list')).json(),target=targets.find(x=>x.type==='page');
 ws=new WebSocket(target.webSocketDebuggerUrl);await once(ws,'open');const call=cdp(ws);
 browserWs=new WebSocket(version.webSocketDebuggerUrl);await once(browserWs,'open');const browserCall=cdp(browserWs);
 await call('Page.enable');await call('Page.navigate',{url:origin+'/'});
 for(let i=0;i<100;i++){await new Promise(r=>setTimeout(r,100));const q=await call('Runtime.evaluate',{expression:'location.origin',returnByValue:true});if(q.result.value===origin)break}
 async function evaluate(expression){const v=await call('Runtime.evaluate',{expression,awaitPromise:true,returnByValue:true,timeout:180000});if(v.exceptionDetails)throw Error(JSON.stringify(v.exceptionDetails));return v.result.value}
 const page=await evaluate("(async()=>{const m=await import('/run.mjs');return await m.run()})()");
 console.log(JSON.stringify({event:'page_capture_finished',cases:page.cases.length,seconds:page.seconds}));
 diagnostics.page=page;
 const workerScript="onmessage=async()=>{try{const m=await import('"+origin+"/run.mjs');postMessage({value:await m.run()})}catch(e){postMessage({error:String(e)})}}";
 const worker=await evaluate("new Promise((resolve,reject)=>{const url=URL.createObjectURL(new Blob(["+JSON.stringify(workerScript)+"],{type:'text/javascript'})),w=new Worker(url,{type:'module'});w.onmessage=e=>{w.terminate();URL.revokeObjectURL(url);e.data.error?reject(Error(e.data.error)):resolve(e.data.value)};w.onerror=e=>{w.terminate();URL.revokeObjectURL(url);reject(Error(e.message))};w.postMessage(1)})");
 console.log(JSON.stringify({event:'worker_capture_finished',cases:worker.cases.length,seconds:worker.seconds}));
 diagnostics.page=page;diagnostics.worker=worker;
 diagnostics.editor_worker_sources=Object.fromEntries(['vba-runner.js','vba-worker.js'].map(n=>[n,hash(fs.readFileSync(path.join(root,'web',n)))]));
 const editorWorker=await evaluate("(async()=>{const m=await import('/run.mjs');return await m.run(true)})()");
 diagnostics.editor_worker=editorWorker;
 console.log(JSON.stringify({event:'editor_worker_capture_finished',cases:editorWorker.cases.length,seconds:editorWorker.seconds}));
 for(const [mode,result]of [['page',page],['worker',worker],['editor_worker',editorWorker]]){
  const expected=new Map(receipt.cases.map(c=>[c.id,c]));result.differences=[];
  if(result.cases.length!==expected.size)throw Error(mode+' case count');
  for(const c of result.cases){let old=expected.get(c.id);if(!old)throw Error('unknown case');
   for(const k of ['result','error','application_state','workbook_changed','oxi_roundtrip_equal','oxi_before_sha256','oxi_reopened_sha256','saved_sha256'])
    if(JSON.stringify(c[k])!==JSON.stringify(old[k]))result.differences.push({id:c.id,field:k,actual:c[k],expected:old[k]});
  }
 }
 // Validate the existing local status view in the same owned browser.
 await call('Page.navigate',{url:origin+'/overview'});let ui;
 for(let i=0;i<100;i++){await new Promise(r=>setTimeout(r,100));ui=await evaluate("(()=>{const t=document.getElementById('uniform-vba-rows');return t?{rows:t.children.length,count:document.getElementById('uniform-vba-count').textContent}:null})()");if(ui?.rows)break}
 if(!ui?.rows)throw Error('overview did not render');diagnostics.overview_initial=ui;
 diagnostics.overview_search=await evaluate("(()=>{const q=document.getElementById('uniform-vba-query');q.value='Range.Value2';q.dispatchEvent(new Event('input'));return {rows:document.getElementById('uniform-vba-rows').children.length,count:document.getElementById('uniform-vba-count').textContent}})()");
 await browserCall('Browser.close');
}catch(e){error=String(e)}
finally{
 if(browserWs?.readyState===1&&!edgeExit){try{await cdp(browserWs)('Browser.close')}catch{}}
 ws?.close();browserWs?.close();
 if(child&&!edgeExit){await Promise.race([once(child,'exit'),new Promise(r=>setTimeout(r,15000))])}
 diagnostics.owned_exit_verified=edgeExit;
 if(child&&!edgeExit){child.kill();diagnostics.owned_forced_stop=true;await new Promise(r=>setTimeout(r,500))}
 if(server){for(const s of sockets)s.destroy();await new Promise(resolve=>server.close(resolve))}
 if(fs.existsSync(profile)&&edgeExit){
  const actual=fs.realpathSync(profile),base=fs.realpathSync(cache);
  if(!actual.startsWith(base+path.sep)||!path.basename(actual).startsWith('owned-conformance-browser-'))throw Error('profile cleanup boundary');
  try{fs.rmSync(actual,{recursive:true,force:true,maxRetries:5,retryDelay:200});diagnostics.profile_removed=true}catch(e){diagnostics.profile_cleanup_error=String(e)}
 }
 diagnostics.artifact_guards_unchanged=Object.entries(artifactHashes).every(([n,s])=>hash(fs.readFileSync(path.join(pkg,n)))===s);
 diagnostics.editor_worker_guards_unchanged=!!diagnostics.editor_worker_sources&&Object.entries(diagnostics.editor_worker_sources).every(([n,s])=>hash(fs.readFileSync(path.join(root,'web',n)))===s);
 diagnostics.at=new Date().toISOString();diagnostics.error=error||null;diagnostics.passed=!error&&edgeExit&&diagnostics.artifact_guards_unchanged&&diagnostics.editor_worker_guards_unchanged&&diagnostics.page.differences.length===0&&diagnostics.worker.differences.length===0&&diagnostics.editor_worker.differences.length===0;
 diagnostics.scope='Actual isolated headless Edge page, module Web Worker and the unchanged editor vba-runner.js/vba-worker.js route. Results and serialized package bytes compared with pinned Node evidence; DOM interaction is not covered.';
 receipt.browser_previous_attempts=[...(receipt.browser_previous_attempts||[]),...(receipt.browser_capture?[receipt.browser_capture]:[])];receipt.browser_capture=diagnostics;receipt.browser_driver_sha256=hash(fs.readFileSync(process.argv[4]));
 fs.writeFileSync(receiptPath,JSON.stringify(receipt)+'\n');
 console.log(JSON.stringify({passed:diagnostics.passed,page_cases:diagnostics.page?.cases.length,worker_cases:diagnostics.worker?.cases.length,editor_worker_cases:diagnostics.editor_worker?.cases.length,page_differences:diagnostics.page?.differences.length,worker_differences:diagnostics.worker?.differences.length,editor_worker_differences:diagnostics.editor_worker?.differences.length,error:diagnostics.error,exit:edgeExit,profile_removed:diagnostics.profile_removed}));
}

"""


def browser_capture(receipt_path, pkg, root):
    """Run the pinned cases in an owned Edge page and module Web Worker."""
    report = json.loads(receipt_path.read_bytes())
    fixtures = report.get("manifest", {}).get("fixtures") or {
        name: {"path": f"tests/fixtures/vba/{name}.xlsm", "sha256": source["fixture_sha256"]}
        for name, source in report["sources"].items()}
    for fixture in fixtures.values():
        path = (root / fixture["path"]).resolve()
        assert path.is_relative_to((root / "tests/fixtures/vba").resolve())
        assert sha(path.read_bytes()) == fixture["sha256"]
    process = subprocess.run(["node", "--max-old-space-size=4096", "--input-type=module", "-e", BROWSER_NODE,
                              str(root), str(pkg), str(receipt_path), str(Path(__file__).resolve())])
    fresh = json.loads(receipt_path.read_bytes())
    return process.returncode == 0 and fresh.get("browser_capture", {}).get("passed") is True


def capture(manifest_path, pkg, root):
    """Capture a reviewed manifest in owned Excel and Node WASM, with readbacks.

    Only bounded cell state is compared; numeric COM values do not prove VBA
    subtypes. The receipt records partial gates and never promotes API members.
    Failed runs persist their exact outcome instead of reporting a vacuous pass.
    """
    import base64
    import gc
    import struct
    import time
    import uuid
    import win32api
    import win32com.client
    import win32event
    import win32process

    supplied = json.loads(manifest_path.read_bytes())
    plan = supplied["manifest"]
    assert plan["schema"] == 1 and plan["cases"]
    assert len({c["id"] for c in plan["cases"]}) == len(plan["cases"])
    for meta in plan["fixtures"].values():
        path = (root / meta["path"]).resolve()
        assert path.is_relative_to((root / "tests/fixtures/vba").resolve())
        assert meta["reviewed"] and sha(path.read_bytes()) == meta["sha256"]
    for name, digest in plan["artifact_hashes"].items():
        assert Path(name).name == name and sha((pkg / name).read_bytes()) == digest
    assert plan["artifact_commit"] and len(plan["artifact_hashes"]["oxidocs_wasm_bg.wasm"]) == 64
    owned_dir = pkg.parent.resolve()
    assert owned_dir.is_dir()
    token = uuid.uuid4().hex
    scratch = [owned_dir / f"owned-conformance-{token}-{kind}.xlsm" for kind in ("native", "oxi")]
    assert all(p.parent.resolve() == owned_dir and not p.exists() for p in scratch)
    guarded = [root / ".git/index", root / "crates/oxicells-core/src/row_defaults.rs",
               root / "crates/oxidocs-wasm/src/spreadsheet_vba.rs"]
    guarded.extend(pkg / name for name in plan["artifact_hashes"])
    guards = {str(path): sha(path.read_bytes()) for path in guarded}
    start = time.monotonic()
    report = {"schema": 1, "at": datetime.datetime.now(datetime.timezone.utc).isoformat(),
              "manifest": plan, "whole_goal_complete": False, "cases": [],
              "source_guards": guards, "feature_evidence": [],
              "limitations": ["Node WASM only; browser is a separate gate",
                              "Bounded cell state, not complete workbook/event/object state",
                              "COM numeric representations are not VBA subtype evidence"]}
    app = book = handle = None

    def snapshot(workbook):
        sheet = workbook.Worksheets(1)
        return {"cells": {address: {"value2": sheet.Range(address).Value2,
                                   "formula2": sheet.Range(address).Formula2,
                                   "has_formula": sheet.Range(address).HasFormula}
                          for address in plan["sample_cells"]},
                "sheet_names": [workbook.Worksheets(i).Name
                                for i in range(1, workbook.Worksheets.Count + 1)]}

    try:
        app = win32com.client.DispatchEx("Excel.Application")
        report["owned_pid"] = win32process.GetWindowThreadProcessId(app.Hwnd)[1]
        handle = win32api.OpenProcess(0x100000, False, report["owned_pid"])
        app.Visible = False
        app.DisplayAlerts = False
        app.EnableEvents = False
        with (Path(app.Path) / "EXCEL.EXE").open("rb") as stream:
            stream.seek(0x3c)
            offset = struct.unpack("<I", stream.read(4))[0]
            stream.seek(offset)
            assert stream.read(4) == b"PE\0\0"
            machine = struct.unpack("<H", stream.read(2))[0]
        report["excel"] = {"version": app.Version, "build": app.Build,
                           "bitness": {0x8664: 64, 0x14c: 32, 0xaa64: 64}[machine],
                           "os": app.OperatingSystem, "locale": app.LanguageSettings.LanguageID(2),
                           "security_observed": app.AutomationSecurity}
        assert report["excel"]["locale"] == plan["excel_ui_language"]
        request = {"root": str(root), "pkg": str(pkg), **plan}
        node = subprocess.run(["node", "--max-old-space-size=4096", "--input-type=module", "-e", CAPTURE_NODE],
                              input=json.dumps(request).encode(), capture_output=True,
                              creationflags=subprocess.CREATE_NO_WINDOW, timeout=300)
        if node.returncode:
            raise RuntimeError(node.stderr.decode("utf-8", errors="replace"))
        captured = json.loads(node.stdout)
        report["sources"] = captured["sources"]
        assert len(captured["cases"]) == len(plan["cases"])
        for index, case in enumerate(captured["cases"]):
            native_path, oxi_path = scratch
            original = root / plan["fixtures"][case["fixture"]]["path"]
            try:
                book = app.Workbooks.Open(str(original), UpdateLinks=0, ReadOnly=True)
                app.Calculation = plan["options"]["applicationState"]["calculation"]
                case["native_result"] = app.Run("'" + book.Name.replace("'", "''") + "'!" + case["procedure"], *case["args"])
                case["native_state"] = snapshot(book)
                case["native_application"] = {"calculation": app.Calculation,
                                              "enableEvents": app.EnableEvents, "displayAlerts": app.DisplayAlerts}
                book.SaveCopyAs(str(native_path))
                book.Close(False)
                book = app.Workbooks.Open(str(native_path), UpdateLinks=0, ReadOnly=True)
                case["native_reopened"] = snapshot(book)
                book.Close(False)
                book = None
                if "saved_base64" in case:
                    body = base64.b64decode(case.pop("saved_base64"))
                    assert sha(body) == case["saved_sha256"]
                    oxi_path.write_bytes(body)
                    book = app.Workbooks.Open(str(oxi_path), UpdateLinks=0, ReadOnly=True)
                    case["oxi_excel_reopened"] = snapshot(book)
                    book.Close(False)
                    book = None
                case["value_equal"] = (isinstance(case["native_result"], str)
                    and not case.get("error") and case.get("result") == {"type": "string", "value": case["native_result"]})
                case["native_roundtrip_equal"] = case["native_state"] == case["native_reopened"]
                case["excel_reopen_equal"] = case["native_reopened"] == case.get("oxi_excel_reopened")
                case["application_equal"] = all(case.get("application_state", {}).get(k) == v
                                                for k, v in case["native_application"].items())
            except Exception as error:
                case["native_capture_error"] = type(error).__name__ + ": " + str(error)
            finally:
                if book is not None:
                    book.Close(False)
                    book = None
                for path in scratch:
                    if path.exists():
                        assert path.parent.resolve() == owned_dir
                        path.unlink()
            report["cases"].append(case)
            if (index + 1) % 20 == 0:
                print(json.dumps({"event": "capture", "done": index + 1, "planned": len(plan["cases"]),
                                  "return_differences": sum(c.get("value_equal") is False for c in report["cases"])}), flush=True)
    except Exception as error:
        report["error"] = type(error).__name__ + ": " + str(error)
    finally:
        if app is not None:
            app.Quit()
            app = None
        gc.collect()
        report["owned_exit_verified"] = handle is not None and win32event.WaitForSingleObject(handle, 15000) == 0
        if handle is not None:
            handle.Close()
        report["guards_unchanged"] = all(sha(Path(path).read_bytes()) == digest for path, digest in guards.items())
        report["summary"] = {"planned": len(plan["cases"]), "completed": len(report["cases"]),
                             "capture_errors": sum(bool(c.get("native_capture_error") or c.get("harness_error")) for c in report["cases"])}
        for label, field in (("return", "value_equal"), ("native_save", "native_roundtrip_equal"),
                             ("oxi_save", "oxi_roundtrip_equal"), ("excel_reopen", "excel_reopen_equal"),
                             ("application", "application_equal")):
            report["summary"][label + "_differences"] = sum(c.get(field) is False for c in report["cases"])
            report["summary"][label + "_unmeasured"] = sum(field not in c for c in report["cases"])
        report["completed"] = (not report.get("error") and not report["summary"]["capture_errors"]
                               and len(report["cases"]) == len(plan["cases"])
                               and report["owned_exit_verified"] and report["guards_unchanged"])
        report["seconds"] = round(time.monotonic() - start, 2)
        manifest_path.write_text(json.dumps(report, ensure_ascii=True, separators=(",", ":")) + "\n", encoding="utf-8")
    print(json.dumps({"completed": report["completed"], "summary": report["summary"],
                      "error": report.get("error"), "seconds": report["seconds"]}), flush=True)
    return report


def sha(data):
    return hashlib.sha256(data).hexdigest()


def feature_id(library, interface, member, operation):
    return f"{library}.{interface}.{member}:{operation}"


def read_declarations(path):
    rows = {}
    for line in path.read_text(encoding="utf-8").splitlines():
        parts = line.split("\t")
        if parts[0] != "M":
            continue
        library, interface = parts[1].split(".", 1)
        operation = INVOKE[int(parts[3])]
        key = feature_id(library, interface, parts[2], operation)
        row = rows.setdefault(key, {
            "id": key, "library": library, "interface": interface,
            "member": parts[2], "operation": operation, "declarations": [],
            "status": "未検証", "evidence": [],
        })
        row["declarations"].append(parts[4:])
    return rows


def read_native(rows):
    import pythoncom
    import pywintypes

    library = pythoncom.LoadRegTypeLib(
        pywintypes.IID("{00020813-0000-0000-C000-000000000046}"), 1, 9, 0)
    aliases, sources, excluded = {}, set(), 0
    for index in range(library.GetTypeInfoCount()):
        info = library.GetTypeInfo(index)
        attr = info.GetTypeAttr()
        if attr.typekind != pythoncom.TKIND_COCLASS:
            continue
        for implementation in range(attr.cImplTypes):
            flags = info.GetImplTypeFlags(implementation)
            target = info.GetRefTypeInfo(info.GetRefTypeOfImplType(implementation))
            target_name = target.GetDocumentation(-1)[0]
            if flags & pythoncom.IMPLTYPEFLAG_FSOURCE:
                sources.add(target_name)
            elif flags & pythoncom.IMPLTYPEFLAG_FDEFAULT:
                aliases[library.GetDocumentation(index)[0]] = target_name
    native_keys = set()
    for index in range(library.GetTypeInfoCount()):
        info = library.GetTypeInfo(index)
        attr = info.GetTypeAttr()
        if attr.typekind not in (pythoncom.TKIND_INTERFACE, pythoncom.TKIND_DISPATCH):
            continue
        name = library.GetDocumentation(index)[0]
        for function in range(attr.cFuncs):
            desc = info.GetFuncDesc(function)
            member = info.GetDocumentation(desc.memid)[0]
            if member in PLUMBING:
                excluded += 1
                continue
            operation = INVOKE[desc.invkind]
            key = feature_id("Excel", name, member, operation)
            native_keys.add(key)
            row = rows.setdefault(key, {
                "id": key, "library": "Excel", "interface": name,
                "member": member, "operation": operation,
                "declarations": [], "status": "未検証", "evidence": [],
            })
            row["native"] = {"dispid": desc.memid, "flags": desc.wFuncFlags,
                             "event_source": name in sources}
    declared = {key for key, row in rows.items()
                if row["library"] == "Excel" and row["declarations"]}
    return {
        "registered_version": "1.9", "type_count": library.GetTypeInfoCount(),
        "native_operations": len(native_keys), "excluded_plumbing_rows": excluded,
        "excluded_plumbing_names": sorted(PLUMBING), "aliases": aliases,
        "event_interfaces": sorted(sources),
        "native_only_operations": sorted(native_keys - declared),
        "declarations_absent_from_native": sorted(declared - native_keys),
    }


def verdict(evidence):
    """No aggregate test count can satisfy a gate or imply a feature pass."""
    gates = evidence.get("gates", {})
    provenance = evidence.get("provenance", {})
    required = ("excel_version", "excel_build", "excel_bitness", "os", "locale",
                "oxi_commit", "wasm_sha256", "source_sha256", "fixture_sha256")
    if not all(provenance.get(key) for key in required):
        return "未検証"
    if not evidence.get("case_ids") or not evidence.get("coverage_contract"):
        return "未検証"
    if any(item.get("status") == "fail" and item.get("receipt") for item in gates.values()):
        return "不一致"
    if evidence.get("unsupported") and evidence.get("reproduction"):
        return "未対応"
    for gate in GATES:
        item = gates.get(gate, {})
        if item.get("status") == "pass" and item.get("receipt"):
            continue
        if item.get("status") == "not_applicable" and item.get("reason") and item.get("receipt"):
            continue
        return "未検証"
    return "合格"


def build(root, native=False, evidence_paths=(), target_artifact=None):
    declarations = root / "crates/oxidocs-wasm/src/vba_reference_signatures.tsv"
    rows = read_declarations(declarations)
    declared_pairs = {(r["interface"], r["member"]) for r in rows.values()
                      if r["library"] == "Excel"}
    native_report = read_native(rows) if native else None
    overview = root / "pipeline_data/vba_overview_20261027.html"
    legacy = None
    if overview.exists():
        match = re.search(r'<script[^>]*id=["\']vba-api-inventory-data["\'][^>]*>(.*?)</script>',
                          overview.read_text(encoding="utf-8"), re.S)
        if match:
            original = json.loads(match[1])["rows"]
            legacy = {"rows": len(original), "assessments": dict(collections.Counter(
                row.get("assessment", "unknown") for row in original)),
                "scope": "Historical member inventory, not uniform gate results"}
    unmapped, receipts = [], []
    measured_batches = []
    artifact_hashes = set()
    for path in evidence_paths:
        body = path.read_bytes()
        report = json.loads(body)
        receipts.append({"path": str(path), "sha256": sha(body)})
        artifact = report.get("manifest", {}).get("artifact_hashes", {}).get("oxidocs_wasm_bg.wasm")
        if not artifact:
            artifact = report.get("oxi_artifact", {}).get("artifacts", {}).get("oxidocs_wasm_bg.wasm", {}).get("sha256")
        if artifact:
            artifact_hashes.add(artifact)
        if report.get("summary"):
            measured_batches.append({"path": str(path), "completed": report.get("completed", False),
                                     "summary": report["summary"],
                                     "browser": {key: report.get("browser_capture", {}).get(key)
                                                 for key in ("passed", "scope")}})
        for item in report.get("feature_evidence", []):
            key = item["id"]
            if key not in rows:
                unmapped.append(key)
                continue
            rows[key]["evidence"].append(item)
            if item.get("provenance", {}).get("wasm_sha256"):
                artifact_hashes.add(item["provenance"]["wasm_sha256"])
    if target_artifact is None:
        if len(artifact_hashes) > 1:
            raise ValueError("Multiple WASM artifacts: select the current one with --artifact-sha; historical evidence is retained")
        target_artifact = next(iter(artifact_hashes), None)
    if target_artifact is not None and not re.fullmatch(r"[0-9a-f]{64}", target_artifact):
        raise ValueError("--artifact-sha must be a lowercase SHA-256")
    for row in rows.values():
        current = [item for item in row["evidence"]
                   if item.get("provenance", {}).get("wasm_sha256") == target_artifact]
        row["current_evidence_count"] = len(current)
        outcomes = [verdict(item) for item in current]
        # Conflicting receipts retain a failure until an explicit reviewed resolution.
        row["status"] = next((status for status in ("不一致", "未対応", "合格")
                              if status in outcomes), "未検証")
    fixtures = root / "tests/fixtures/vba"
    counts = dict(collections.Counter(p.suffix for p in fixtures.iterdir() if p.is_file()))
    scope = [
        "Excel object model: properties, methods, events, including hidden members",
        "VBA syntax, compiler diagnostics, operators, coercion, arrays and ByRef",
        "VBA standard library and VBA/Office/stdole reference declarations",
        "Projects, document/class modules, import/export, lifetime and save/reopen",
        "WorksheetFunction, Evaluate, UDFs, calculation and dependency propagation",
        "Browser/Worker integration, events, UserForms and controls",
        "External automation, native API/DLL, files and OS-dependent behavior",
    ]
    return {
        "schema": 1, "generated_at": datetime.datetime.now(datetime.timezone.utc).isoformat(),
        "root_head": subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=root, text=True).strip(),
        "declaration_sha256": sha(declarations.read_bytes()), "legacy": legacy,
        "target_wasm_sha256": target_artifact,
        "excel_declared_member_pairs": len(declared_pairs), "native": native_report,
        "gates": list(GATES), "scope": [{"domain": item, "status": "未検証"} for item in scope],
        "scope_note": "API operation totals do not enumerate the entire VBA language; no whole-goal percentage is defined.",
        "fixture_file_counts": counts, "evidence_receipts": receipts,
        "measured_batches": measured_batches,
        "operations_with_case_evidence": sum(bool(row["evidence"]) for row in rows.values()),
        "unmapped_evidence_ids": unmapped,
        "status_counts": dict(collections.Counter(row["status"] for row in rows.values())),
        "library_counts": dict(collections.Counter(row["library"] for row in rows.values())),
        "rows": sorted(rows.values(), key=lambda item: item["id"]),
    }


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--browser", type=Path, help="Replay an existing receipt in isolated Edge and a Web Worker")
    parser.add_argument("--capture", type=Path, help="Run the reviewed manifest in this receipt file and replace it with fresh evidence")
    parser.add_argument("--pkg", type=Path, help="Existing WASM package to capture; hashes must match the manifest")
    parser.add_argument("--native", action="store_true", help="Read installed Excel typelib; does not launch Excel")
    parser.add_argument("--evidence", type=Path, action="append", default=[])
    parser.add_argument("--artifact-sha", help="WASM SHA-256 whose current status is being assessed; other evidence remains historical")
    parser.add_argument("--overview", action="store_true", help="Refresh a summary in the existing local HTML overview")
    parser.add_argument("--output", type=Path, default=Path("pipeline_data/vba_conformance_audit.json"))
    args = parser.parse_args()
    root = Path(__file__).resolve().parents[1]
    if args.browser:
        if not args.pkg or args.capture:
            parser.error("--browser requires --pkg and cannot be combined with --capture")
        raise SystemExit(0 if browser_capture(args.browser.resolve(), args.pkg.resolve(), root) else 1)
    if args.capture:
        if not args.pkg:
            parser.error("--capture requires --pkg")
        result = capture(args.capture.resolve(), args.pkg.resolve(), root)
        raise SystemExit(0 if result["completed"] else 1)
    report = build(root, args.native, args.evidence, args.artifact_sha)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(report, ensure_ascii=False, separators=(",", ":")) + "\n", encoding="utf-8")
    if args.overview:
        update_overview(root, report)
    print(json.dumps({key: report[key] for key in ("excel_declared_member_pairs", "status_counts", "library_counts", "unmapped_evidence_ids")}, ensure_ascii=True))


def update_overview(root, report):
    """Embed searchable operation rows so the local file needs no HTTP server."""
    from html import escape

    path = root / "pipeline_data/vba_overview_20261027.html"
    html = path.read_text(encoding="utf-8")
    start, end = "<!-- uniform-vba-audit:start -->", "<!-- uniform-vba-audit:end -->"
    rows = [[r["id"], r["status"], bool(r.get("native", {}).get("event_source")),
             r["member"].startswith("_Dummy"),
             len({case for item in r["evidence"] for case in item.get("case_ids", [])})]
            for r in report["rows"]]
    data = json.dumps(rows, ensure_ascii=True, separators=(",", ":")).replace("<", "\\u003c")
    counts = " / ".join(f"{escape(status)} {report['status_counts'].get(status, 0):,}"
                        for status in ("合格", "不一致", "未検証", "未対応"))
    batches = report["measured_batches"]
    batch_lines = []
    for batch in batches:
        summary = batch["summary"]
        browser = batch["browser"]
        batch_lines.append(f"{Path(batch['path']).name}: {summary['completed']}/{summary['planned']} ケース、"
                           f"返却文字列の差 {summary['return_differences']}、"
                           f"保存後の Excel 読み直しの差 {summary['excel_reopen_differences']}。"
                           + ("実 Edge のページ・Web Worker でも同じ結果と保存バイト列を確認。"
                              if browser.get("passed") else "実ブラウザーの比較は未完了。"))
    batch_text = " ".join(batch_lines) if batch_lines else "実測結果はまだ取り込み前です。"
    batch_text += (f" ケース証拠の対応付けは {report['operations_with_case_evidence']} 操作。"
                   "限定ケースの比較であり、全機能の合格数には加算していません。")
    panel = '''
<section id="uniform-vba-audit" style="margin:20px;padding:24px;background:white;border:3px solid #00796c;border-radius:12px">
<h2>全機能を共通基準で検証する台帳</h2>
<p>更新: DATE。既存一覧の LEGACY 件と、参照定義の Excel 型・メンバー DECLARED 組を照合しました。
下表は取得・代入・呼び出しを分けた操作単位です。Office / VBA / stdole も含み、予約用 _Dummy DUMMY 件も削除せず残しています。
言語構文、コンパイル、プロジェクト、イベント、UserForm、外部 COM / DLL などは別の検証範囲として残ります。<strong>VBA 全体の完成率はまだ算出できません。</strong></p>
<p><strong>COUNTS</strong></p><p>「未検証」は共通基準を満たす証拠が未登録という意味です。既存の実装や個別テストがゼロという意味ではありません。</p>
<p>共通基準: 同じ入力と実行環境／値／実際の VBA 型／エラー／状態変化／保存して Oxi・Excel で再読込／実ブラウザー。
適用外には理由と証拠を必須とし、ソース内の実装名やテスト総数だけでは合格にしません。</p>
<p>BATCH</p><p>順序: 既存の比較証拠を台帳へ対応付け → 日常操作と VBA 言語・実行系を横断する検証 → 残る全 API・イベント・外部連携。
各範囲で同じ基準を使い、部分的な一致と未検証を併記します。全件の終了日は、ケース不足と実測所要時間が判明するまで未確定です。</p>
<label>機能を検索 <input id="uniform-vba-query" placeholder="Range.Value2 / WorksheetFunction / AppEvents" style="width:min(100%,500px)"></label>
<label>判定 <select id="uniform-vba-status"><option value="">すべて</option><option>合格</option><option>不一致</option><option>未検証</option><option>未対応</option></select></label>
<p id="uniform-vba-count"></p><table><thead><tr><th>ライブラリ・型・メンバー : 操作</th><th>共通判定</th><th>補足</th><th>対応付けた限定ケース</th></tr></thead><tbody id="uniform-vba-rows"></tbody></table>
<button id="uniform-vba-more" type="button">次の 100 件を表示</button>
</section>
<script id="uniform-vba-data" type="application/json">DATA</script>
<script>(()=>{const data=JSON.parse(document.getElementById('uniform-vba-data').textContent),q=document.getElementById('uniform-vba-query'),s=document.getElementById('uniform-vba-status');let limit=100;
function render(){const f=data.filter(r=>r[0].toLowerCase().includes(q.value.toLowerCase())&&(!s.value||r[1]===s.value)),body=document.getElementById('uniform-vba-rows');body.replaceChildren();
for(const r of f.slice(0,limit)){const row=document.createElement('tr');for(const text of [r[0],r[1],r[3]?'予約メンバー':r[2]?'イベント元インターフェイス':'',r[4]||'未対応付け']){const cell=document.createElement('td');cell.textContent=text;row.appendChild(cell)}body.appendChild(row)}
document.getElementById('uniform-vba-count').textContent=f.length.toLocaleString()+' 件中 '+Math.min(f.length,limit).toLocaleString()+' 件を表示';document.getElementById('uniform-vba-more').hidden=f.length<=limit;}
q.addEventListener('input',()=>{limit=100;render()});s.addEventListener('change',()=>{limit=100;render()});document.getElementById('uniform-vba-more').addEventListener('click',()=>{limit+=100;render()});render();})();</script>
'''.replace("DATE", escape(report["generated_at"])).replace("COUNTS", counts).replace("BATCH", escape(batch_text))
    panel = panel.replace("LEGACY", f"{(report.get('legacy') or {}).get('rows', 0):,}")
    panel = panel.replace("DECLARED", f"{report['excel_declared_member_pairs']:,}")
    panel = panel.replace("DUMMY", f"{sum(r['member'].startswith('_Dummy') for r in report['rows']):,}")
    panel = panel.replace("DATA", data)
    block = start + panel + end
    if start in html:
        html = re.sub(re.escape(start) + ".*?" + re.escape(end), lambda _: block, html, count=1, flags=re.S)
    else:
        html, count = re.subn(r"(<body[^>]*>)", lambda m: m[1] + block, html, count=1)
        if count != 1:
            raise ValueError("overview has no body element")
    path.write_text(html, encoding="utf-8")


if __name__ == "__main__":
    main()
