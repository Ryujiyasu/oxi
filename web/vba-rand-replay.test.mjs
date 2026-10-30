// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.
import assert from 'node:assert/strict';
import {readFile} from 'node:fs/promises';
import path from 'node:path';
import {fileURLToPath,pathToFileURL} from 'node:url';
import {Worker as ThreadWorker} from 'node:worker_threads';
const root=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
const pkg=path.resolve(process.argv[2]||path.join(root,'crates/oxidocs-wasm/pkg'));
const bindingUrl='data:text/javascript;base64,'+Buffer.from(await readFile(path.join(pkg,'oxidocs_wasm.js'))).toString('base64');
const wasm=await import(bindingUrl);await wasm.default({module_or_path:await readFile(path.join(pkg,'oxidocs_wasm_bg.wasm'))});
const fixture=path.join(root,'tests/fixtures/vba');
const corpus=JSON.parse(await readFile(path.join(fixture,'rand_replay.json'),'utf8'));
const groups=await Promise.all(corpus.groups.map(async group=>({...group,source:await readFile(path.join(fixture,group.source),'utf8'),packet:await readFile(path.join(fixture,group.package))})));
// Excel 20527: Volatile False suppresses RAND's volatility for the whole
// formula. Both orders were measured in three separate Excel processes.
const intrinsicSource=`Option Explicit
Private hitsB As Long
Private trace As String, traceStage As String
Private savedCalculation As Long
Public Function VolB() As Long
    Application.Volatile False
    hitsB = hitsB + 1
    trace = trace & traceStage & ":B|"
    VolB = 20
End Function
Public Function CallTrace() As String
    CallTrace = trace
End Function
Public Function IntrinsicProbe(ByVal kind As Long) As String
    Dim previous As Long, report As String, expression As String, old As Double
    previous = Application.Calculation: Application.Calculation = xlCalculationManual
    Range("Z1").ClearContents: hitsB = 0: trace = "": traceStage = "seed"
    If kind = 0 Then expression = "=RAND()+VolB()" Else expression = "=VolB()+RAND()"
    Range("Z1").Formula2 = expression: Application.CalculateFull
    old = Range("Z1").Value2: report = "seed:" & CStr(hitsB)
    traceStage = "clean": Application.Calculate
    report = report & "|clean:" & CStr(hitsB) & "," & CStr(old <> Range("Z1").Value2): old = Range("Z1").Value2
    traceStage = "dirty": Range("D9").Value2 = 1: Application.Calculate
    report = report & "|dirty:" & CStr(hitsB) & "," & CStr(old <> Range("Z1").Value2): old = Range("Z1").Value2
    traceStage = "again": Application.Calculate
    report = report & "|again:" & CStr(hitsB) & "," & CStr(old <> Range("Z1").Value2)
    IntrinsicProbe = report
    Range("Z1").ClearContents: Application.Calculation = previous
End Function
Public Function VolPrime(ByVal kind As Long) As String
    Dim expression As String
    savedCalculation = Application.Calculation: Application.Calculation = xlCalculationManual
    Range("Z1").ClearContents: hitsB = 0: trace = "": traceStage = "seed"
    If kind = 0 Then expression = "=RAND()+VolB()" Else expression = "=VolB()+RAND()"
    Range("Z1").Formula2 = expression: Application.CalculateFull
    VolPrime = CStr(hitsB)
End Function
Public Function VolNext() As String
    traceStage = "next": Application.Calculate
    VolNext = CStr(hitsB)
End Function
Public Sub VolFinish()
    Range("Z1").ClearContents: Application.Calculation = savedCalculation
End Sub
`;
groups.unshift({source:intrinsicSource,packet:await readFile(path.join(fixture,'udf_returned_references.xlsm')),cases:[
    ...[0,1].map(id=>({id:`intrinsic-${id}`,steps:[
        {procedure:'IntrinsicProbe',arguments:[id],expected:'seed:2|clean:2,False|dirty:2,False|again:2,False'},
        {procedure:'CallTrace',arguments:[],expected:'seed:B|seed:B|',sessionOnly:true},
    ]})),
    ...[0,1].map(id=>({id:`intrinsic-session-${id}`,sessionOnly:true,steps:[
        {procedure:'VolPrime',arguments:[id],expected:'2'},
        {procedure:'VolNext',arguments:[],expected:'2'},
        {procedure:'VolNext',arguments:[],expected:'2'},
        {procedure:'CallTrace',arguments:[],expected:'seed:B|seed:B|'},
    ]})),
]});
const options={errorLanguage:'ja-JP',hostAnsiCodePage:932};
function check(outcome,step,label){
    assert.equal(outcome.error,undefined,label+': '+JSON.stringify(outcome.error));
    if(step.expected!==undefined)assert.deepEqual(outcome.result,{type:'string',value:step.expected},label);
    else assert.equal(outcome.result.value,step.expected_number,label+' seed');
    return JSON.parse(JSON.stringify(outcome.workbook));
}
const counts={standalone:0,direct_session:0,worker:0};
for(const route of ['standalone','direct_session'])for(const group of groups)for(const item of group.cases){
    if(route==='standalone' && item.sessionOnly)continue;
    const session=route==='direct_session'?new wasm.SpreadsheetVbaSession():undefined;
    try{
        let book=wasm.parse_spreadsheet(group.packet);
        if(session)assert.deepEqual(session.compile(group.source,{}),{compiled:true});
        for(const step of item.steps){
            if(route==='standalone' && step.sessionOnly)continue;
            const args=[book,group.source,step.procedure,step.arguments,0,`case-${item.id}.xlsm`,options];
            book=check(session?session.run(...args):wasm.run_spreadsheet_vba_with_outcome(...args),step,route+' Native case '+item.id);
        }
        counts[route]++;
    }finally{session?.free();}
}
const shim=`import init,{SpreadsheetVbaSession,spreadsheet_vba_project_name,spreadsheet_vba_source}from ${JSON.stringify(bindingUrl)};import{readFile}from'node:fs/promises';await init({module_or_path:await readFile(${JSON.stringify(path.join(pkg,'oxidocs_wasm_bg.wasm'))})});export{SpreadsheetVbaSession,spreadsheet_vba_project_name,spreadsheet_vba_source};export default function ready(){return Promise.resolve();}`;
const workerSource=(await readFile(path.join(root,'web/vba-worker.js'),'utf8')).replace("'./oxidocs_wasm.js'",JSON.stringify('data:text/javascript;base64,'+Buffer.from(shim).toString('base64')));
const bridge=`import{parentPort,workerData}from'node:worker_threads';const queued=[];let listener;globalThis.self={postMessage:value=>parentPort.postMessage(value),addEventListener:(type,handler)=>{listener=handler;for(const data of queued)listener({data});}};parentPort.on('message',data=>listener?listener({data}):queued.push(data));await import('data:text/javascript;base64,'+Buffer.from(workerData).toString('base64'));`;
class BrowserWorker{
    constructor(){this.thread=new ThreadWorker(new URL('data:text/javascript;base64,'+Buffer.from(bridge).toString('base64')),{workerData:workerSource});}
    addEventListener(type,handler){this.thread.on(type,value=>handler(type==='message'?{data:value}:value));}
    postMessage(value){this.thread.postMessage(value);}
    terminate(){return this.thread.terminate();}
}
const previous=globalThis.Worker;globalThis.Worker=BrowserWorker;
const runner=await import(pathToFileURL(path.join(root,'web/vba-runner.js')));
try{
    for(const group of groups)for(const item of group.cases){
        runner.closeVbaProject();let book=wasm.parse_spreadsheet(group.packet);
        assert.deepEqual(await runner.compileVbaInWorker(group.source,options,book),{compiled:true});
        for(const step of item.steps)book=check(await runner.runVbaInWorker(book,group.source,step.procedure,step.arguments,0,`case-${item.id}.xlsm`,options),step,'Worker Native case '+item.id);
        counts.worker++;
    }
}finally{runner.closeVbaProject();globalThis.Worker=previous;}
assert.deepEqual(counts,{standalone:106,direct_session:108,worker:108});
console.log(JSON.stringify({native_cases:108,...counts,complete_results:true,seed_results:true,workbook_json_between_calls:true,passed:true}));
