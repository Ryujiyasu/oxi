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
const workerSource=(await readFile(path.join(root,'web/vba-worker.js'),'utf8')).replace("'./oxidocs_wasm.js'",JSON.stringify('data:text/javascript;base64,'+Buffer.from(shim).toString('base64')));
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
console.log(JSON.stringify({native_steps:fixture.cases.length,expected_end_terminations:1,invalid_snapshots:invalidStates.length,direct_session_worker:true,caller_snapshot_isolation:true,passed:true}));
