// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { Worker as ThreadWorker } from 'node:worker_threads';
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const pkg = path.resolve(process.argv[2] || path.join(root, 'crates/oxidocs-wasm/pkg'));
const bindingUrl = 'data:text/javascript;base64,' + Buffer.from(await readFile(path.join(pkg, 'oxidocs_wasm.js'))).toString('base64');
const wasm = await import(bindingUrl);
await wasm.default({ module_or_path: await readFile(path.join(pkg, 'oxidocs_wasm_bg.wasm')) });
const cases = JSON.parse(await readFile(path.join(root, 'tests/fixtures/vba/constant_types.json'), 'utf8'));
const emptyBook = () => ({ sheets: [{ name: 'Sheet1', rows: [], col_count: 0, col_widths: [], default_col_width: 0, default_row_height: 15, merge_cells: [] }] });
async function sequence(testCase,invoke,worker=false) {
  const source=await readFile(path.join(root,'tests/fixtures/vba',testCase.source),'utf8');
  let book=emptyBook();
  for (const step of testCase.outcome==='compile'?[null,null]:testCase.steps) {
    if (testCase.outcome==='compile' && worker) {
      await assert.rejects(invoke(book,source,testCase.procedure),error=>error.message.includes('コンパイル エラー: '+testCase.message),testCase.name);
      continue;
    }
    const result=await invoke(book,source,testCase.procedure);
    if (testCase.outcome==='compile') {
      assert.equal(result.error.kind,'compile',testCase.name);
      assert.ok(result.error.message.endsWith('コンパイル エラー: '+testCase.message),testCase.name);
    } else {
      assert.equal(result.error,undefined,testCase.name);
      assert.deepEqual(result.result,{type:'string',value:step.value},testCase.name);
    }
    book=result.workbook;
  }
}
for (const item of cases) {
  const session=new wasm.SpreadsheetVbaSession();
  try {await sequence(item,(book,source,procedure)=>session.run(book,source,procedure,[],0,'Compile.xlsm',{errorLanguage:'ja-JP'}));}
  finally {session.free();}
}
const shim = `import init, { SpreadsheetVbaSession, spreadsheet_vba_project_name, spreadsheet_vba_source } from ${JSON.stringify(bindingUrl)};
import { readFile } from 'node:fs/promises';
await init({ module_or_path: await readFile(${JSON.stringify(path.join(pkg, 'oxidocs_wasm_bg.wasm'))}) });
export { SpreadsheetVbaSession, spreadsheet_vba_project_name, spreadsheet_vba_source };
export default function ready() { return Promise.resolve(); }`;
const workerSource = (await readFile(path.join(root, 'web/vba-worker.js'), 'utf8')).replace("'./oxidocs_wasm.js'", JSON.stringify('data:text/javascript;base64,' + Buffer.from(shim).toString('base64')));
const bridge = `import { parentPort, workerData } from 'node:worker_threads';
const queued = []; let listener;
globalThis.self = { postMessage: value => parentPort.postMessage(value), addEventListener: (type, handler) => { listener = handler; for (const data of queued) listener({ data }); } };
parentPort.on('message', data => listener ? listener({ data }) : queued.push(data));
await import('data:text/javascript;base64,' + Buffer.from(workerData).toString('base64'));`;
class BrowserWorker {
  constructor() { this.thread = new ThreadWorker(new URL('data:text/javascript;base64,' + Buffer.from(bridge).toString('base64')), { workerData: workerSource }); }
  addEventListener(type, handler) { this.thread.on(type, value => handler(type === 'message' ? { data: value } : value)); }
  postMessage(value) { this.thread.postMessage(value); }
  terminate() { void this.thread.terminate(); }
}
const runner = await import(pathToFileURL(path.join(root, 'web/vba-runner.js')));
const priorWorker = globalThis.Worker;
try {
  globalThis.Worker = BrowserWorker;
  for (const item of cases) {
    runner.closeVbaProject();
    await sequence(item,(book,source,procedure)=>runner.runVbaInWorker(book,source,procedure,[],0,'Compile.xlsm',{applicationState:{},errorLanguage:'ja-JP'}),true);
  }
} finally { runner.closeVbaProject(); globalThis.Worker = priorWorker; }
console.log(JSON.stringify({excel_cases:cases.length,value_sequences:cases.filter(item=>item.outcome==='values').length,compile_cases:cases.filter(item=>item.outcome==='compile').length,value_steps:cases.filter(item=>item.outcome==='values').reduce((sum,item)=>sum+item.steps.length,0),direct_and_worker:true,passed:true}));
