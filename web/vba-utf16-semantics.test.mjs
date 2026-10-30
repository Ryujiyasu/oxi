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
const fixture=path.join(root,'tests/fixtures/vba');
const table=async name=>(await readFile(path.join(fixture,name),'utf8')).split(/\r?\n/).filter(line=>line&&!line.startsWith('#')).map(line=>line.split('\t'));
const source=await readFile(path.join(fixture,'worksheet_utf16.bas'),'utf8');
const characterSource=await readFile(path.join(fixture,'character_utf16.bas'),'utf8');
const richSource=await readFile(path.join(fixture,'character_rich_utf16.bas'),'utf8');
const rich=new Map(await table('character_rich_utf16.tsv'));
const expressions=await table('worksheet_utf16.tsv');const modes=new Map(await table('worksheet_utf16_modes.tsv'));const characters=new Map(await table('character_utf16.tsv'));
const host=new Map((await table('worksheet_utf16_host.tsv')).map(([version,name,value])=>[version+':'+name,value]));
const original=new Uint8Array(await readFile(path.join(root,'crates/oxicells-core/tests/fixtures/hidden_rows_cols.xlsx')));
const options={hostAnsiCodePage:932,errorLanguage:'ja-JP'};
const seed=version=>{const book=wasm.parse_spreadsheet(original);book.calculation.compatibility_version=version;return book;};
const verify=(outcome,expected,label)=>{assert.equal(outcome.error,undefined,label+': '+JSON.stringify(outcome.error));assert.deepEqual(outcome.result,{type:'string',value:expected},label);return outcome.workbook;};
const outcomes=[];
for(const version of [1,2]){
  const formulaReport=expressions.filter(row=>Number(row[0])===version).map(row=>row[2]).join('|');
  const session=new wasm.SpreadsheetVbaSession();
  try{
    assert.deepEqual(session.compile(source,options),{compiled:true});
    let book=session.run(seed(version),source,'SetupCells',[],0,'utf16.xlsx',options).workbook;
    book=verify(session.run(book,source,'FormulaCases',[],0,'utf16.xlsx',options),formulaReport,'FormulaCases v'+version);
    book=verify(session.run(book,source,'Modes',[],0,'utf16.xlsx',options),modes.get(String(version)),'Modes v'+version);
    book=verify(session.run(book,source,'CharacterCases',[],0,'utf16.xlsx',options),characters.get('CharacterCases'),'CharacterCases v'+version);
    for(const name of ['AfterEdit','SupportedFunctions'])book=verify(session.run(book,source,name,[],0,'utf16.xlsx',options),host.get(version+':'+name),name+' v'+version);
    assert.deepEqual(session.compile(richSource,options),{compiled:true});
    book=verify(session.run(book,richSource,'RichCases',[],0,'utf16.xlsx',options),rich.get('RichCases'),'partial rich fonts v'+version);
    const saved=wasm.edit_xlsx_from_workbook(original,book);
    if(process.argv[3])await writeFile(path.join(process.argv[3],'oxi-mode-'+version+'.xlsx'),saved);
    const reopened=wasm.parse_spreadsheet(saved);assert.equal(reopened.calculation.compatibility_version||1,version);
    verify(session.run(reopened,richSource,'AfterEdit',[],0,'utf16.xlsx',options),host.get(version+':AfterEdit'),'saved fonts and formula values v'+version);
    verify(session.run(reopened,richSource,'ReadRich',[],0,'utf16.xlsx',options),rich.get('ReadRich'),'saved partial fonts v'+version);
    outcomes.push({version,book:reopened});
  }finally{session.free();}
  verify(wasm.run_spreadsheet_vba_with_outcome(seed(version),characterSource,'BoundaryCases',[],0,'utf16.xlsx',options),characters.get('BoundaryCases'),'160 character boundaries v'+version);
}
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
    let book=(await runner.runVbaInWorker(seed(version),source,'SetupCells',[],0,'utf16.xlsx',options)).workbook;
    book=verify(await runner.runVbaInWorker(book,source,'FormulaCases',[],0,'utf16.xlsx',options),expressions.filter(row=>Number(row[0])===version).map(row=>row[2]).join('|'),'Worker formula v'+version);
    book=verify(await runner.runVbaInWorker(book,source,'Modes',[],0,'utf16.xlsx',options),modes.get(String(version)),'Worker modes v'+version);
    book=verify(await runner.runVbaInWorker(book,source,'CharacterCases',[],0,'utf16.xlsx',options),characters.get('CharacterCases'),'Worker characters v'+version);
    for(const name of ['AfterEdit','SupportedFunctions'])book=verify(await runner.runVbaInWorker(book,source,name,[],0,'utf16.xlsx',options),host.get(version+':'+name),'Worker '+name+' v'+version);
    runner.closeVbaProject();assert.deepEqual(await runner.compileVbaInWorker(richSource,options,book),{compiled:true});
    verify(await runner.runVbaInWorker(book,richSource,'RichCases',[],0,'utf16.xlsx',options),rich.get('RichCases'),'Worker partial fonts v'+version);
    runner.closeVbaProject();assert.deepEqual(await runner.compileVbaInWorker(characterSource,options,seed(version)),{compiled:true});
    verify(await runner.runVbaInWorker(seed(version),characterSource,'BoundaryCases',[],0,'utf16.xlsx',options),characters.get('BoundaryCases'),'Worker boundaries v'+version);
  }
}finally{runner.closeVbaProject();globalThis.Worker=previous;}
console.log(JSON.stringify({passed:true,compatibility_versions:2,character_boundaries:160,native_formula_expressions:44,worksheetfunction_absent_members:5,worker:true,save_reopen:true,rich_text_font:true}));
