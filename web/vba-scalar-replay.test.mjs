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
const corpus=JSON.parse(await readFile(path.join(fixture,'udf_scalar_replay_current_module.json'),'utf8'));
const referenceCorpus=JSON.parse(await readFile(path.join(fixture,'udf_reference_text_chain_current_module.json'),'utf8'));
const objectCorpus=JSON.parse(await readFile(path.join(fixture,'udf_object_error_current_module.json'),'utf8'));
const directObjectCorpus=JSON.parse(await readFile(path.join(fixture,'udf_object_scalar_boundary_current_module.json'),'utf8'));
corpus.groups.push(...referenceCorpus.groups,...objectCorpus.groups,...directObjectCorpus.groups);
assert.equal(corpus.groups.reduce((total,group)=>total+group.cases.length,0),309);
const groups=await Promise.all(corpus.groups.map(async group=>({...group,source:await readFile(path.join(fixture,group.source),'utf8'),packet:await readFile(path.join(fixture,group.package))})));
const options={errorLanguage:'ja-JP',hostAnsiCodePage:932};
function check(outcome,step,label){
    assert.equal(outcome.error,undefined,label+': '+JSON.stringify(outcome.error));
    assert.ok(step.expected_returns.some(value=>outcome.result?.type==='string'&&outcome.result.value===value),label+': '+JSON.stringify(outcome.result));
    return JSON.parse(JSON.stringify(outcome.workbook));
}
const counts={standalone:0,direct_session:0,worker:0};
for(const route of ['standalone','direct_session'])for(const group of groups)for(const item of group.cases){
    const session=route==='direct_session'?new wasm.SpreadsheetVbaSession():undefined;
    try{
        let book=wasm.parse_spreadsheet(group.packet);
        if(session)assert.deepEqual(session.compile(group.source,{}),{compiled:true});
        for(const step of item.steps){
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
assert.deepEqual(counts,{standalone:309,direct_session:309,worker:309});
console.log(JSON.stringify({native_variable_cases:23,native_fixed_cases:286,total_native_cases:309,...counts,complete_results:true,workbook_json_between_calls:true,passed:true}));
