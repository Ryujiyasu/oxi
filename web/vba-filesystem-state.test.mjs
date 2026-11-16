// SPDX-License-Identifier: MIT OR Apache-2.0
import assert from 'node:assert/strict';
import {readFile} from 'node:fs/promises';
import path from 'node:path';
import {fileURLToPath,pathToFileURL} from 'node:url';
import {Worker as ThreadWorker} from 'node:worker_threads';
const root=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
const pkg=path.resolve(process.argv[2]||path.join(root,'crates/oxidocs-wasm/pkg'));
const bindingUrl='data:text/javascript;base64,'+Buffer.from(await readFile(path.join(pkg,'oxidocs_wasm.js'))).toString('base64');
const wasm=await import(bindingUrl);await wasm.default({module_or_path:await readFile(path.join(pkg,'oxidocs_wasm_bg.wasm'))});
const fixture=JSON.parse(await readFile(path.join(root,'tests/fixtures/vba/namespace_lifetime.json'),'utf8'));
const source=await readFile(path.join(root,'tests/fixtures/vba',fixture.source),'utf8');
const bytes=new Uint8Array(await readFile(path.join(root,'tests/fixtures/vba',fixture.package)));
const initialApplication={displayAlerts:false,directoryState:fixture.initial_directory_state};
const book=()=>wasm.parse_spreadsheet(bytes);
const options=(item,state,fileSystemState)=>({errorLanguage:'ja-JP',applicationState:state,fileSystemState,
    documentModules:wasm.spreadsheet_vba_document_modules(bytes),fileContext:{packageBytes:bytes,fullName:item.full_name,readOnly:true,folders:fixture.folders,files:fixture.files}});
function check(execution,item,index){
    if(item.expected_termination){assert.equal(execution.error?.kind,'terminated',`step ${index}`);assert.deepEqual(execution.result,{type:'empty'});}
    else {assert.equal(execution.error,undefined,`step ${index}`);assert.deepEqual(execution.result,{type:'string',value:item.expected},`step ${index}`);}
    assert.ok(execution.file_system_state);
}
const emptyState=()=>({folders:[],files:[],mountedFolders:[],mountedFiles:[]});
const invalidStates=[0,false,'',[],{},
    {...emptyState(),folders:['relative']},
    {...emptyState(),folders:['C:\\work','c:/WORK']},
    {...emptyState(),files:[{path:'C:\\same',bytes:[1]},{path:'c:/SAME',bytes:[2]}]},
    {...emptyState(),unknown:1},
    {...emptyState(),files:[{path:'C:\\bad',bytes:[1],unknown:1}]},
    {...emptyState(),mountedFiles:['relative']},
    {...emptyState(),mountedFolders:['C:\\work','c:/WORK']},
    {...emptyState(),files:[{path:'C:\\bad',bytes:[256]}]},
    {...emptyState(),files:[{path:'C:\\bad'}]},
    {...emptyState(),folders:null}];
const first=fixture.cases[0];
check(wasm.run_spreadsheet_vba_with_outcome(book(),source,first.procedure,first.arguments,0,'input-a.xlsm',options(first,initialApplication)),first,0);
let session,project,workbook,applicationState=initialApplication,namespace;
try {
    for(const [index,item]of fixture.cases.entries()){
        const changed=project!==item.project;
        if(changed){session?.free();session=new wasm.SpreadsheetVbaSession();project=item.project;workbook=book();}
        const result=session.run(workbook,source,item.procedure,item.arguments,0,'input-'+item.project+'.xlsm',options(item,applicationState,changed?namespace:undefined));
        check(result,item,index);workbook=JSON.parse(JSON.stringify(result.workbook));applicationState=JSON.parse(JSON.stringify(result.application_state));
        namespace=JSON.parse(JSON.stringify(result.file_system_state));
        if(index===4){
            for(const invalid of invalidStates){
                assert.throws(()=>session.run(workbook,source,'NamespaceRead',item.arguments,0,'input-b.xlsm',options(item,applicationState,invalid)),/invalid VBA/,JSON.stringify(invalid));
            }
        }
    }
} finally {session?.free();}
const shim=`import init, {SpreadsheetVbaSession,spreadsheet_vba_source,spreadsheet_vba_project_name} from ${JSON.stringify(bindingUrl)};
import {readFile} from 'node:fs/promises';await init({module_or_path:await readFile(${JSON.stringify(path.join(pkg,'oxidocs_wasm_bg.wasm'))})});
export {SpreadsheetVbaSession,spreadsheet_vba_source,spreadsheet_vba_project_name};export default function ready(){return Promise.resolve();}`;
const workerSource=(await readFile(path.join(root,'web/vba-worker.js'),'utf8')).replace("'./vba-editor-settings.js'", JSON.stringify(new URL('./vba-editor-settings.js', import.meta.url).href)).replace("'./oxidocs_wasm.js'",JSON.stringify('data:text/javascript;base64,'+Buffer.from(shim).toString('base64')));
const bridge=`import {parentPort,workerData} from 'node:worker_threads';const queued=[];let listener;
globalThis.self={postMessage:value=>parentPort.postMessage(value),addEventListener:(type,handler)=>{listener=handler;for(const data of queued)listener({data});}};
parentPort.on('message',data=>listener?listener({data}):queued.push(data));await import('data:text/javascript;base64,'+Buffer.from(workerData).toString('base64'));`;
class BrowserWorker {
    constructor(){this.thread=new ThreadWorker(new URL('data:text/javascript;base64,'+Buffer.from(bridge).toString('base64')),{workerData:workerSource});}
    addEventListener(type,handler){this.thread.on(type,value=>handler(type==='message'?{data:value}:value));}
    postMessage(value){this.thread.postMessage(value);}
    terminate(){void this.thread.terminate();}
}
const originalWorker=globalThis.Worker;globalThis.Worker=BrowserWorker;
const runner=await import(pathToFileURL(path.join(root,'web/vba-runner.js')));project=undefined;
try {
    for(const [index,item]of fixture.cases.entries()){
        if(project!==item.project){runner.closeVbaProject();project=item.project;workbook=book();}
        let execution;
        try {execution=await runner.runVbaInWorker(workbook,source,item.procedure,item.arguments,0,'input-'+item.project+'.xlsm',options(item,index===0?initialApplication:undefined),bytes);}
        catch(error){if(!error.execution)throw error;execution=error.execution;}
        check(execution,item,index);workbook=execution.workbook;
        if(index===4){
            for(const invalid of [false,{},invalidStates[8],invalidStates[9]]){
                await assert.rejects(runner.runVbaInWorker(workbook,source,'NamespaceRead',item.arguments,0,'input-b.xlsm',options(item,undefined,invalid),bytes),/invalid VBA/);
            }
        }
        // Returned bytes, empty directories and consumed-mount identities are
        // caller-owned; mutating them must not change the runner's next mount.
        const returned=execution.file_system_state;
        for(const file of returned.files)file.bytes.fill(255);
        returned.files.length=0;returned.folders.length=0;returned.mountedFiles.length=0;returned.mountedFolders.length=0;
    }
} finally {runner.closeVbaProject();globalThis.Worker=originalWorker;}
async function validateBridge(w,b){
 const ensure=(condition,label)=>{if(!condition)throw Error(label)};
 const decode=v=>v.type==='array'?v.value.values.map(decode):v.type==='empty'?null:v.value;
 const source=`Attribute VB_Name = "BrowserDates"
Public Function Probe(ByVal mode As Long) As Variant
 Dim value As Variant, octet As Byte
 On Error Resume Next
 If mode=1 Then
  octet=90
  Open "C:\\work\\data.bin" For Binary As #1
  Put #1,,octet
  Close #1
 ElseIf mode=2 Then
  Kill "C:\\work\\data.bin"
 End If
 Err.Clear
 value=FileDateTime("C:\\work\\data.bin")
 Probe=Array(Err.Number,TypeName(value),VarType(value),value)
End Function
`;
 const bytes=new Uint8Array(b.groups[0].bytes),old=43890.50283564815,path='C:\\work\\data.bin';
 const context={packageBytes:bytes,fullName:'C:\\work\\bridge.xlsm',folders:['C:\\work'],
  files:[{path,bytes:new Uint8Array([65,66,67])}],modifiedTimes:[{path,value:old}]};
 const options=()=>({errorLanguage:'ja-JP',projectName:'VBAProject',hostAnsiCodePage:932,
  applicationState:{displayAlerts:false,enableEvents:false},fileContext:context});
 let workbook=w.parse_spreadsheet(bytes),session=new w.SpreadsheetVbaSession();
 const validation=[];
 try{
  let out=session.run(workbook,source,'BrowserDates.Probe',[0],0,'bridge.xlsm',options());
  ensure(JSON.stringify(decode(out.result))===JSON.stringify([0,'Date',7,old]),'initial Date metadata');
  workbook=out.workbook;
  const invalid=[null,0,{},[null],[{path,value:NaN}],[{path,value:Infinity}],
   [{path,value:true}],[{path,value:'43890'}],[{path}],[{path,value:old,unknown:1}],
   [{path,value:-657435}],[{path,value:2958466}],
   [{path,value:old},{path:'c:/WORK/DATA.bin',value:old}],
   [{path:'C:\\work\\absent.bin',value:old}]];
  for(let index=0;index<invalid.length;index++){
   let message='';
   try{session.run(workbook,source,'BrowserDates.Probe',[0],0,'bridge.xlsm',{
    ...options(),fileContext:{...context,modifiedTimes:invalid[index]}})}catch(error){message=String(error)}
   ensure(/invalid VBA|modifiedTimes/i.test(message),'invalid metadata rejected '+index+': '+message);
   out=session.run(workbook,source,'BrowserDates.Probe',[0],0,'bridge.xlsm',options());
   ensure(JSON.stringify(decode(out.result))===JSON.stringify([0,'Date',7,old]),'invalid metadata leaves live state '+index);
   workbook=out.workbook;validation.push({index,message,result:out.result});
  }
  const snapshot=structuredClone(out.file_system_state);snapshot.modifiedTimes[0].value=NaN;
  let message='';try{session.run(workbook,source,'BrowserDates.Probe',[0],0,'bridge.xlsm',{
   ...options(),fileSystemState:snapshot})}catch(error){message=String(error)}
  ensure(/invalid VBA|modifiedTimes/i.test(message),'invalid snapshot time rejected');
  out=session.run(workbook,source,'BrowserDates.Probe',[0],0,'bridge.xlsm',options());
  ensure(JSON.stringify(decode(out.result))===JSON.stringify([0,'Date',7,old]),'invalid snapshot leaves live state');
  validation.push({index:'snapshot',message,result:out.result});
 }finally{session.free()}
 const observations=[];
 workbook=w.parse_spreadsheet(bytes);
 async function run(mode,extra={}){
  const out=await runner.runVbaInWorker(workbook,source,'BrowserDates.Probe',[mode],0,'bridge.xlsm',{
   ...options(),...extra},bytes);
  ensure(!out.error,'runner macro error '+JSON.stringify(out.error));workbook=out.workbook;
  observations.push(structuredClone(out));return out;
 }
 try{
  runner.closeVbaProject();
  let out=await run(0,{fileSystemState:{folders:[],files:[],mountedFolders:[],mountedFiles:[]}});
  ensure(JSON.stringify(decode(out.result))===JSON.stringify([0,'Date',7,old]),'real Worker initial metadata');
  out=await run(1);const current=decode(out.result)[3];
  ensure(JSON.stringify(decode(out.result).slice(0,3))===JSON.stringify([0,'Date',7])&&Number.isFinite(current)&&current!==old,'real Worker writes new time');
  out=await run(0);ensure(decode(out.result)[3]===current,'repeated bootstrap preserves written date');
  const returned=out.file_system_state.modifiedTimes.find(x=>x.path.toLowerCase()===path.toLowerCase());
  ensure(returned,'returned snapshot includes file date');returned.value=old;
  ensure(runner.closeVbaProject(),'owned Worker closed before restoring snapshot');
  out=await run(0);ensure(decode(out.result)[3]===current,'cached snapshot owns a deep copy of dates across actual Worker restart');
  out=await run(2);ensure(JSON.stringify(decode(out.result))===JSON.stringify([53,'Empty',0,null]),'deleted metadata is absent');
  ensure(runner.closeVbaProject(),'owned Worker closed after deletion');
  out=await run(0);ensure(JSON.stringify(decode(out.result))===JSON.stringify([53,'Empty',0,null]),'consumed mount does not resurrect after Worker restart');
 }finally{runner.closeVbaProject()}
 return {validation,observations,passed:true};
}

let modificationDateControls;
try{globalThis.Worker=BrowserWorker;modificationDateControls=await validateBridge(wasm,{groups:[{bytes}]});}
finally{runner.closeVbaProject();globalThis.Worker=originalWorker;}
assert(modificationDateControls.passed);


async function validateDirectoryEntries(w, bytes) {
 const source=`Attribute VB_Name = "DirectoryDates"
Public Function Probe(ByVal mode As Long) As Variant
 Dim stamp As Variant, attributes As Long
 On Error Resume Next
 If mode=1 Then
  Open "C:\\work\\new.bin" For Output As #1
  Close #1
 ElseIf mode=2 Then
  attributes=GetAttr("C:\\work")
 End If
 stamp=FileDateTime("C:\\work")
 Probe=Array(Err.Number,TypeName(stamp),VarType(stamp),CDbl(stamp))
End Function
`;
 const old=43890.50283564815;
 const context={packageBytes:bytes,fullName:'C:\\work\\dates.xlsm',folders:['C:\\work'],
  files:[{path:'C:\\work\\data.bin',bytes:new Uint8Array([65,66,67])}],modifiedTimes:[{path:'C:\\work',value:old}]};
 const options={fileContext:context,errorLanguage:'ja-JP'};
 const decode=out=>out.result.value.values.map(x=>x.value);
 const expected=value=>[0,'Date',7,value];
 let workbook=w.parse_spreadsheet(bytes);
 const run=async (mode,extra={})=>{
  const out=await runner.runVbaInWorker(workbook,source,'DirectoryDates.Probe',[mode],0,'dates.xlsm',{...options,...extra},bytes);
  assert.equal(out.error,undefined);workbook=out.workbook;return out;
 };
 const key=path=>path.replaceAll('/','\\').toLowerCase().replace(/\\+$/,'');
 runner.closeVbaProject();
 try {
  let out=await run(1,{fileSystemState:{folders:[],files:[],mountedFolders:[],mountedFiles:[]}});assert.deepEqual(decode(out),expected(old));
  const physical=out.file_system_state.modifiedTimes.find(x=>key(x.path)==='c:\\work').value;
  assert.notEqual(physical,old,'new child updates the physical folder date');
  const snapshot=structuredClone(out.file_system_state);
  const entry=out.file_system_state.directoryEntryTimes.find(x=>key(x.path)==='c:\\work');
  assert.equal(entry.value,old,'FileDateTime retains the unrefreshed directory entry');
  entry.value=physical;
  assert(runner.closeVbaProject());
  out=await run(0);assert.deepEqual(decode(out),expected(old),'caller mutation cannot refresh the cached snapshot on Worker restart');
  const session=new w.SpreadsheetVbaSession();
  let validations=0;
  try {
   const invoke=(state)=>session.run(workbook,source,'DirectoryDates.Probe',[0],0,'dates.xlsm',{...options,...(state?{fileSystemState:state}:{})});
   assert.deepEqual(decode(invoke(snapshot)),expected(old));
   const invalid=[null,0,{},[null],[{path:'C:\\work',value:NaN}],[{path:'C:\\work',value:Infinity}],
    [{path:'C:\\work',value:true}],[{path:'C:\\work',value:'43890'}],[{path:'C:\\work'}],
    [{path:'C:\\work',value:old,unknown:1}], [{path:'C:\\work',value:-657435}],[{path:'C:\\work',value:2958466}],
    [{path:'C:\\work',value:old},{path:'c:/WORK',value:old}],
    [{path:'C:\\work\\absent',value:old}],[{path:'C:\\work\\data.bin',value:old}]];
   for(const facts of invalid) {
    assert.throws(()=>invoke({...snapshot,directoryEntryTimes:facts}),/invalid VBA|directoryEntryTimes/i);
    assert.deepEqual(decode(invoke()),expected(old),'rejected directory metadata leaves the live state intact');validations++;
   }
  } finally {session.free();}
  out=await run(2);assert.deepEqual(decode(out),expected(physical),'GetAttr publishes the saved physical time');
  assert(runner.closeVbaProject());
  out=await run(0);assert.deepEqual(decode(out),expected(physical),'refreshed time persists through another Worker restart');
  return {validations,worker_calls:4,passed:true};
 } finally {runner.closeVbaProject();}
}
let directoryEntryControls;
try{globalThis.Worker=BrowserWorker;directoryEntryControls=await validateDirectoryEntries(wasm,bytes);}
finally{runner.closeVbaProject();globalThis.Worker=originalWorker;}
assert(directoryEntryControls.passed);

console.log(JSON.stringify({directory_entry_validation_cases:directoryEntryControls.validations,directory_entry_worker_calls:directoryEntryControls.worker_calls,file_date_validation_cases:modificationDateControls.validation.length,file_date_worker_calls:modificationDateControls.observations.length,native_steps:fixture.cases.length,expected_end_terminations:1,invalid_snapshots:invalidStates.length,direct_session_worker:true,caller_snapshot_isolation:true,passed:true}));
