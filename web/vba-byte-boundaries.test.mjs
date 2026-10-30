// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.
import assert from 'node:assert/strict';
import {readFile,writeFile}from'node:fs/promises';
import path from'node:path';
import {fileURLToPath,pathToFileURL}from'node:url';
import {Worker as ThreadWorker}from'node:worker_threads';
const root=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
const pkg=path.resolve(process.argv[2]||path.join(root,'crates/oxidocs-wasm/pkg'));
const bindingUrl='data:text/javascript;base64,'+Buffer.from(await readFile(path.join(pkg,'oxidocs_wasm.js'))).toString('base64');
const wasm=await import(bindingUrl);await wasm.default({module_or_path:await readFile(path.join(pkg,'oxidocs_wasm_bg.wasm'))});
const fixture=path.join(root,'tests/fixtures/vba');
const source=await readFile(path.join(fixture,'worksheet_byte_ends.bas'),'utf8')+'\n'+await readFile(path.join(fixture,'worksheet_byte_persistence.bas'),'utf8');
const original=new Uint8Array(await readFile(path.join(root,'crates/oxicells-core/tests/fixtures/hidden_rows_cols.xlsx')));
const tables=new Map();
for(const[name,file]of[['ByteBoundaries','boundaries'],['ByteBoundaryControls','controls'],['ByteNumericLimits','limits'],['ByteLongLimits','long_limits'],['ByteEndScanning','ends']]){
  tables.set(name,(await readFile(path.join(fixture,'worksheet_byte_'+file+'.tsv'),'utf8')).split(/\r?\n/).filter(l=>l&&!l.startsWith('#')).map(l=>l.split('\t')));
}
const host=new Map((await readFile(path.join(fixture,'worksheet_byte_host.tsv'),'utf8')).split(/\r?\n/).filter(l=>l&&!l.startsWith('#')).map(l=>l.split('\t')));
const expressions=['LEFTB(J20,3)','MIDB(J20,2,1)','REPLACEB(J20,3,1,K20)','FINDB(L20,J20,1)','SEARCHB("?",J20,3)'];
function expected(name,version){
  if(name==='HostByteFunctions')return host.get(String(version));
  const samples=new Map();for(const[,units,,value]of tables.get(name).filter(r=>Number(r[0])===version)){if(!samples.has(units))samples.set(units,[]);samples.get(units).push(value);}
  return[...samples].map(([units,values])=>units+'='+values.join('|')+'|').join('#');
}
const persisted=version=>expressions.map(expression=>tables.get('ByteBoundaries').find(r=>Number(r[0])===version&&r[1]==='4:97:-10179:-8704:98'&&r[2]===expression)[3]).join('|');
const options={worksheetDbcsLanguage:false,hostAnsiCodePage:932,errorLanguage:'ja-JP'};
const seed=version=>{const book=wasm.parse_spreadsheet(original);book.calculation.compatibility_version=version;return book;};
const verify=(outcome,report,label)=>{assert.equal(outcome.error,undefined,label+': '+JSON.stringify(outcome.error));assert.deepEqual(outcome.result,{type:'string',value:report},label);return outcome.workbook;};
const cellReport=value=>{
  if(value.Number!==undefined)return'5:'+value.Number;
  const units=value.Utf16||Array.from({length:value.String.length},(_,i)=>value.String.charCodeAt(i));
  return'S:'+units.length+units.map(u=>':'+(u>=32768?u-65536:u)).join('');
};
const cached=book=>[12,13,14,15,16].map(col=>cellReport(book.sheets[0].rows.find(r=>r.index===20).cells.find(c=>c.col===col).value)).join('|');
const methods=[...tables.keys(),'HostByteFunctions'];
for(const version of[1,2]){
  for(const name of methods)verify(wasm.run_spreadsheet_vba_with_outcome(seed(version),source,name,[],0,'byte-boundaries.xlsx',options),expected(name,version),'direct '+name+' v'+version);
  const session=new wasm.SpreadsheetVbaSession();
  try{
    assert.deepEqual(session.compile(source,{errorLanguage:'ja-JP'}),{compiled:true});let book=seed(version);
    for(const name of methods)book=verify(session.run(book,source,name,[],0,'byte-boundaries.xlsx',options),expected(name,version),'session '+name+' v'+version);
    book=verify(session.run(book,source,'PersistByteFormulas',[],0,'byte-boundaries.xlsx',options),persisted(version),'formula caches v'+version);
    assert.equal(cached(book),persisted(version));
    const saved=wasm.edit_xlsx_from_workbook(original,book);
    if(process.argv[3])await writeFile(path.join(process.argv[3],'byte-mode-'+version+'.xlsx'),saved);
    const reopened=wasm.parse_spreadsheet(saved);assert.equal(reopened.calculation.dbcs_language,undefined);assert.equal(reopened.calculation.compatibility_version||1,version);
    assert.equal(cached(reopened),persisted(version),'stored caches before macro execution v'+version);
    verify(session.run(reopened,source,'ReadByteFormulas',[],0,'byte-boundaries.xlsx',options),persisted(version),'read saved formulas v'+version);
  }finally{session.free();}
}
const shim=`import init,{SpreadsheetVbaSession,spreadsheet_vba_project_name,spreadsheet_vba_source}from ${JSON.stringify(bindingUrl)};
import{readFile}from'node:fs/promises';await init({module_or_path:await readFile(${JSON.stringify(path.join(pkg,'oxidocs_wasm_bg.wasm'))})});
export{SpreadsheetVbaSession,spreadsheet_vba_project_name,spreadsheet_vba_source};export default function ready(){return Promise.resolve();}`;
const workerSource=(await readFile(path.join(root,'web/vba-worker.js'),'utf8')).replace("'./oxidocs_wasm.js'",JSON.stringify('data:text/javascript;base64,'+Buffer.from(shim).toString('base64')));
const bridge=`import{parentPort,workerData}from'node:worker_threads';const queued=[];let listener;globalThis.self={postMessage:value=>parentPort.postMessage(value),addEventListener:(type,handler)=>{listener=handler;for(const data of queued)listener({data});}};parentPort.on('message',data=>listener?listener({data}):queued.push(data));await import('data:text/javascript;base64,'+Buffer.from(workerData).toString('base64'));`;
class BrowserWorker{constructor(){this.thread=new ThreadWorker(new URL('data:text/javascript;base64,'+Buffer.from(bridge).toString('base64')),{workerData:workerSource});}addEventListener(type,handler){this.thread.on(type,value=>handler(type==='message'?{data:value}:value));}postMessage(value){this.thread.postMessage(value);}terminate(){return this.thread.terminate();}}
const previous=globalThis.Worker;globalThis.Worker=BrowserWorker;const runner=await import(pathToFileURL(path.join(root,'web/vba-runner.js')));
try{
  for(const version of[1,2]){
    runner.closeVbaProject();assert.deepEqual(await runner.compileVbaInWorker(source,options,seed(version)),{compiled:true});let book=seed(version);
    for(const name of methods)book=verify(await runner.runVbaInWorker(book,source,name,[],0,'byte-boundaries.xlsx',options),expected(name,version),'Worker '+name+' v'+version);
    verify(await runner.runVbaInWorker(book,source,'PersistByteFormulas',[],0,'byte-boundaries.xlsx',options),persisted(version),'Worker formulas v'+version);
  }
}finally{runner.closeVbaProject();globalThis.Worker=previous;}
console.log(JSON.stringify({passed:true,native_formula_cases:[...tables.values()].reduce((n,t)=>n+t.length,0),host_calls:16,compatibility_versions:2,direct:true,session:true,worker:true,save_reopen:true,preserved_formula_caches:10}));
