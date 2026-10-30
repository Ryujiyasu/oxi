// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.
import assert from 'node:assert/strict';
import { readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { Worker as ThreadWorker } from 'node:worker_threads';
const root=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
const pkg=path.resolve(process.argv[2]||path.join(root,'crates/oxidocs-wasm/pkg'));
const bindingUrl='data:text/javascript;base64,'+Buffer.from(await readFile(path.join(pkg,'oxidocs_wasm.js'))).toString('base64');
const wasm=await import(bindingUrl);await wasm.default({module_or_path:await readFile(path.join(pkg,'oxidocs_wasm_bg.wasm'))});
const source=await readFile(path.join(root,'tests/fixtures/vba/worksheet_text_extensions.bas'),'utf8');
const original=new Uint8Array(await readFile(path.join(root,'crates/oxicells-core/tests/fixtures/hidden_rows_cols.xlsx')));
const options={hostAnsiCodePage:932,errorLanguage:'ja-JP',worksheetDbcsLanguage:false};
const seed=version=>{const book=wasm.parse_spreadsheet(original);book.calculation.compatibility_version=version;return book;};
async function expected(file,version){
  const rows=(await readFile(path.join(root,'tests/fixtures/vba',file),'utf8')).split(/\r?\n/).filter(line=>line&&!line.startsWith('#')).map(line=>line.split('\t')).filter(row=>Number(row[0])===version);
  const samples=new Map();for(const [,sample,,value]of rows){if(!samples.has(sample))samples.set(sample,[]);samples.get(sample).push(value);}
  return [...samples].map(([sample,values])=>sample+'='+values.join('|')+'|').join('#');
}
const verify=(outcome,report,label)=>{assert.equal(outcome.error,undefined,label+': '+JSON.stringify(outcome.error));assert.deepEqual(outcome.result,{type:'string',value:report},label);return outcome.workbook;};
for(const version of [1,2]){
  const report=await expected('worksheet_text_extensions.tsv',version);
  let book=verify(wasm.run_spreadsheet_vba_with_outcome(seed(version),source,'TextExtensions',[],0,'extensions.xlsx',options),report,'direct v'+version);
  const session=new wasm.SpreadsheetVbaSession();
  try{
    assert.deepEqual(session.compile(source,{errorLanguage:'ja-JP'}),{compiled:true});
    book=verify(session.run(book,source,'TextExtensions',[],0,'extensions.xlsx',options),report,'session v'+version);
    if(version===2)book=verify(session.run(book,source,'ByteCases',[],0,'extensions.xlsx',options),await expected('worksheet_byte_extensions.tsv',2),'byte session');
    const saved=wasm.edit_xlsx_from_workbook(original,book);
    if(process.argv[3])await writeFile(path.join(process.argv[3],'extensions-mode-'+version+'.xlsx'),saved);
    const reopened=wasm.parse_spreadsheet(saved);
    assert.equal(reopened.calculation.dbcs_language,undefined);
    verify(session.run(reopened,source,'TextExtensions',[],0,'extensions.xlsx',options),report,'saved/reopened context v'+version);
  }finally{session.free();}
}
assert.throws(()=>wasm.run_spreadsheet_vba_with_outcome(seed(1),source,'TextExtensions',[],0,'extensions.xlsx',{worksheetDbcsLanguage:'false'}),/boolean|bool/i);
const languageSource='Public Function LanguageProbe() As String\nLanguageProbe = CStr(Application.Evaluate("LENB(UNICHAR(28450))"))\nEnd Function';
verify(wasm.run_spreadsheet_vba_with_outcome(seed(1),languageSource,'LanguageProbe',[],0,'extensions.xlsx',{worksheetDbcsLanguage:true}),'2','legacy DBCS');
verify(wasm.run_spreadsheet_vba_with_outcome(seed(1),languageSource,'LanguageProbe',[],0,'extensions.xlsx',{}),'2','legacy default');
verify(wasm.run_spreadsheet_vba_with_outcome(seed(1),languageSource,'LanguageProbe',[],0,'extensions.xlsx',options),'1','explicit SBCS independent of ANSI code page');
const shim=`import init,{SpreadsheetVbaSession,spreadsheet_vba_project_name,spreadsheet_vba_source}from ${JSON.stringify(bindingUrl)};
import{readFile}from'node:fs/promises';await init({module_or_path:await readFile(${JSON.stringify(path.join(pkg,'oxidocs_wasm_bg.wasm'))})});
export{SpreadsheetVbaSession,spreadsheet_vba_project_name,spreadsheet_vba_source};export default function ready(){return Promise.resolve();}`;
const workerSource=(await readFile(path.join(root,'web/vba-worker.js'),'utf8')).replace("'./oxidocs_wasm.js'",JSON.stringify('data:text/javascript;base64,'+Buffer.from(shim).toString('base64')));
const bridge=`import{parentPort,workerData}from'node:worker_threads';const queued=[];let listener;globalThis.self={postMessage:value=>parentPort.postMessage(value),addEventListener:(type,handler)=>{listener=handler;for(const data of queued)listener({data});}};parentPort.on('message',data=>listener?listener({data}):queued.push(data));await import('data:text/javascript;base64,'+Buffer.from(workerData).toString('base64'));`;
class BrowserWorker{constructor(){this.thread=new ThreadWorker(new URL('data:text/javascript;base64,'+Buffer.from(bridge).toString('base64')),{workerData:workerSource});}addEventListener(type,handler){this.thread.on(type,value=>handler(type==='message'?{data:value}:value));}postMessage(value){this.thread.postMessage(value);}terminate(){return this.thread.terminate();}}
const previous=globalThis.Worker;globalThis.Worker=BrowserWorker;const runner=await import(pathToFileURL(path.join(root,'web/vba-runner.js')));
try{
  for(const version of [1,2]){
    runner.closeVbaProject();assert.deepEqual(await runner.compileVbaInWorker(source,options,seed(version)),{compiled:true});
    const book=verify(await runner.runVbaInWorker(seed(version),source,'TextExtensions',[],0,'extensions.xlsx',options),await expected('worksheet_text_extensions.tsv',version),'Worker v'+version);
    if(version===2)verify(await runner.runVbaInWorker(book,source,'ByteCases',[],0,'extensions.xlsx',options),await expected('worksheet_byte_extensions.tsv',2),'Worker bytes');
  }
}finally{runner.closeVbaProject();globalThis.Worker=previous;}
console.log(JSON.stringify({passed:true,native_text_cases:440,native_byte_cases:112,direct:true,session:true,worker:true,save_reopen:true,explicit_calculation_language:true}));
