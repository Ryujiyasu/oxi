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
const groups=[];
for(const name of ['file_attributes','file_attributes_more','file_attributes_edges']){
    const fixture=JSON.parse(await readFile(path.join(root,'tests/fixtures/vba',name+'.json'),'utf8'));
    groups.push({fixture,source:await readFile(path.join(root,'tests/fixtures/vba',fixture.source),'utf8'),
        bytes:new Uint8Array(await readFile(path.join(root,'tests/fixtures/vba',fixture.package)))});
}
const options=(group,item,state)=>({errorLanguage:'ja-JP',hostAnsiCodePage:932,applicationState:state,
    documentModules:wasm.spreadsheet_vba_document_modules(group.bytes),fileContext:{packageBytes:group.bytes,fullName:item.full_name,readOnly:true,folders:item.folders,files:item.files}});
const sorted=files=>files.map(file=>({path:file.path,bytes:Array.from(file.bytes)})).sort((a,b)=>a.path<b.path?-1:a.path>b.path?1:0);
function check(execution,item){
    const label='Native case '+item.kind;
    assert.equal(execution.error,undefined,label);
    assert.deepEqual(execution.result,{type:'string',value:item.expected},label);
    const base=item.arguments[1].toLowerCase(),prefix=base+'\\',state=execution.file_system_state;
    assert.deepEqual(sorted(state.files.filter(file=>file.path.toLowerCase().startsWith(prefix))),sorted(item.expected_files),label+' file bytes/paths');
    assert.deepEqual(state.folders.map(path=>path.toLowerCase()).filter(path=>path===base||path.startsWith(prefix)).sort(),item.expected_folder_keys,label+' folders');
    const kept=new Set(item.expected_files.map(file=>file.path));
    assert.deepEqual([...(execution.deleted_files||[])].sort(),item.files.filter(file=>!kept.has(file.path)).map(file=>file.path).sort(),label+' deletion outputs');
    const overrides=new Map((state.attributes||[]).map(entry=>[entry.path.toLowerCase(),entry.value]));
    for(const entry of item.expected_attributes){
        const folder=state.folders.some(path=>path.toLowerCase()===entry.path);
        assert.equal(overrides.get(entry.path) ?? (folder?16:32),entry.value,label+' attributes '+entry.path);
    }
    for(const entry of state.attributes||[]){
        const key=entry.path.toLowerCase();
        if(key===base||key.startsWith(prefix))assert.ok(item.expected_attributes.some(expected=>expected.path===key),label+' stale attributes');
    }
}
let direct=0;
for(const group of groups){
    const state={displayAlerts:false,directoryState:group.fixture.initial_directory_state};
    for(const item of group.fixture.cases){
        const execution=wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(group.bytes),group.source,item.procedure,item.arguments,0,'input.xlsm',options(group,item,state));
        check(execution,item);direct++;
    }
}
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
const runner=await import(pathToFileURL(path.join(root,'web/vba-runner.js')));let worker=0;
try {
    for(const group of groups){
        for(const item of group.fixture.cases){
            runner.closeVbaProject();
            const state=worker===0?{displayAlerts:false,directoryState:group.fixture.initial_directory_state}:undefined;
            const execution=await runner.runVbaInWorker(wasm.parse_spreadsheet(group.bytes),group.source,item.procedure,item.arguments,0,'input.xlsm',options(group,item,state),group.bytes);
            check(execution,item);worker++;
        }
    }
    const group=groups[0],item=group.fixture.cases.find(item=>item.kind===2);
    runner.closeVbaProject();
    const changed=await runner.runVbaInWorker(wasm.parse_spreadsheet(group.bytes),group.source,item.procedure,item.arguments,0,'input.xlsm',options(group,item),group.bytes);
    check(changed,item);
    const namespace=JSON.parse(JSON.stringify(changed.file_system_state));
    const storedPath=item.arguments[1]+'\\held.txt';
    const reader='Public Function StoredAttributes(ByVal path As String) As String\nStoredAttributes = CStr(GetAttr(path))\nEnd Function';
    const read=(state)=>wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(group.bytes),reader,'StoredAttributes',[storedPath],0,'input.xlsm',{errorLanguage:'ja-JP',fileSystemState:state});
    assert.deepEqual(read(namespace).result,{type:'string',value:'1'},'serialized namespace attributes');
    const legacy=structuredClone(namespace);delete legacy.attributes;
    assert.deepEqual(read(legacy).result,{type:'string',value:'32'},'legacy snapshots restore default attributes');
    const entry={path:storedPath,value:1};
    for(const attributes of [null,{},[null],[{...entry,extra:true}],[{value:1}],[{path:storedPath,value:'1'}],
        [{...entry,value:0.5}],[{...entry,value:-1}],[{...entry,value:64}],[{...entry,value:16}],
        [{path:storedPath+'missing',value:1}],[entry,{path:storedPath.toUpperCase(),value:0}]]){
        assert.throws(()=>read({...namespace,attributes}),/invalid VBA fileSystemState/,'reject malformed metadata');
    }
    changed.file_system_state.attributes[0].value=55;
    changed.file_system_state.attributes[0].path='C:\\mutated';
    changed.file_system_state.attributes.length=0;
    runner.closeVbaProject();
    const recreated=await runner.runVbaInWorker(wasm.parse_spreadsheet(group.bytes),reader,'StoredAttributes',[storedPath],0,'input.xlsm',{errorLanguage:'ja-JP'},group.bytes);
    assert.equal(recreated.error,undefined);
    assert.deepEqual(recreated.result,{type:'string',value:'1'},'project recreation/caller mutation isolation');
} finally {runner.closeVbaProject();globalThis.Worker=originalWorker;}
assert.equal(direct,64);assert.equal(worker,64);
console.log(JSON.stringify({native_cases:64,direct,worker,complete_results:true,post_file_bytes_and_folders:true,post_attributes:true,deletion_outputs:true,passed:true}));
