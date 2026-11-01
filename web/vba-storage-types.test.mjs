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
const cases = JSON.parse(await readFile(path.join(root, 'tests/fixtures/vba/storage_types.json'), 'utf8'));
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

const resultBoundarySource = "Public Type BoundaryRecord\n SafeValue As LongLong\n LargeValue As LongLong\nEnd Type\nPrivate retainedBoundaryObject As Collection\nPublic Function BoundarySmall() As LongLong\n BoundarySmall = CLngLng(42)\nEnd Function\nPublic Function BoundarySafe() As LongLong\n BoundarySafe = CLngLng(4503599627370496#) * CLngLng(2) - CLngLng(1)\nEnd Function\nPublic Function BoundaryUnsafe() As LongLong\n BoundaryUnsafe = CLngLng(4503599627370496#) * CLngLng(2)\nEnd Function\nPublic Function BoundaryAdjacent() As LongLong\n BoundaryAdjacent = BoundaryUnsafe() + CLngLng(1)\nEnd Function\nPublic Function BoundaryNegativeSafe() As LongLong\n BoundaryNegativeSafe = -BoundarySafe()\nEnd Function\nPublic Function BoundaryNegativeUnsafe() As LongLong\n BoundaryNegativeUnsafe = -BoundaryUnsafe()\nEnd Function\nPublic Function BoundaryMinimum() As LongLong\n BoundaryMinimum = -CLngLng(4611686018427387904#) - CLngLng(4611686018427387904#)\nEnd Function\nPublic Function BoundaryMaximum() As LongLong\n BoundaryMaximum = CLngLng(4611686018427387904#) - CLngLng(1) + CLngLng(4611686018427387904#)\nEnd Function\nPublic Function BoundaryDouble() As Double\n BoundaryDouble = 9007199254740992#\nEnd Function\nPublic Function BoundaryWorkbook() As Object\n Set BoundaryWorkbook = ThisWorkbook\nEnd Function\nPublic Function BoundaryCollection() As Object\n Set BoundaryCollection = New Collection\nEnd Function\nPublic Function BoundaryRetained() As Object\n If retainedBoundaryObject Is Nothing Then Set retainedBoundaryObject = New Collection\n Set BoundaryRetained = retainedBoundaryObject\nEnd Function\nPublic Function BoundaryArray() As Variant\n Dim values(0 To 3) As Variant\n values(0) = BoundaryUnsafe()\n Set values(1) = New Collection\n Set values(2) = New Collection\n values(3) = CLngLng(42)\n BoundaryArray = values\nEnd Function\n\nPublic Function BoundaryRecordResult() As BoundaryRecord\n Dim entry As BoundaryRecord\n entry.SafeValue = CLngLng(42)\n entry.LargeValue = BoundaryUnsafe()\n BoundaryRecordResult = entry\nEnd Function\nPublic Function DescribeBoundaryRecord() As String\n Dim entry As BoundaryRecord\n entry = BoundaryRecordResult()\n DescribeBoundaryRecord = CStr(entry.SafeValue) & \"|\" & CStr(entry.LargeValue)\nEnd Function\n";
async function verifyResultBoundaries(invoke) {
  let book = emptyBook();
  async function run(procedure) {
    const execution = await invoke(book, resultBoundarySource, procedure);
    assert.equal(execution.error, undefined, procedure);
    book = execution.workbook;
    assert.deepEqual(JSON.parse(JSON.stringify(execution.result)), execution.result,
      procedure + ': result survives JSON transport without precision loss');
    return execution.result;
  }
  const integerCases = [
    ['BoundarySmall', 42],
    ['BoundarySafe', 9007199254740991],
    ['BoundaryUnsafe', '9007199254740992'],
    ['BoundaryAdjacent', '9007199254740993'],
    ['BoundaryNegativeSafe', -9007199254740991],
    ['BoundaryNegativeUnsafe', '-9007199254740992'],
    ['BoundaryMinimum', '-9223372036854775808'],
    ['BoundaryMaximum', '9223372036854775807'],
  ];
  for (const [procedure, value] of integerCases) {
    assert.deepEqual(await run(procedure), {type:'integer', value}, procedure);
  }
  assert.deepEqual(await run('BoundaryDouble'), {type:'double', value:9007199254740992});
  const workbook = await run('BoundaryWorkbook');
  assert.equal(workbook.type, 'object');
  assert.equal(workbook.value.kind, 'Workbook');
  assert.equal(typeof workbook.value.handle, 'number');
  function collection(value) {
    assert.equal(value.type, 'object');
    assert.equal(value.value.kind, 'Collection');
    assert.equal(typeof value.value.handle, 'string');
    assert.ok(BigInt(value.value.handle) >= (1n << 63n));
    return value.value.handle;
  }
  collection(await run('BoundaryCollection'));
  const retained = collection(await run('BoundaryRetained'));
  assert.equal(collection(await run('BoundaryRetained')), retained,
    'the retained VBA object keeps its identity between invocations');
  const array = await run('BoundaryArray');
  assert.equal(array.type, 'array');
  assert.equal(array.value.lower_bound, 0);
  assert.deepEqual(array.value.values[0], {type:'integer',value:'9007199254740992'});
  assert.notEqual(collection(array.value.values[1]), collection(array.value.values[2]),
    'distinct runtime objects keep distinct 64-bit handles');
  assert.deepEqual(array.value.values[3], {type:'integer',value:42});
  const record = await run('BoundaryRecordResult');
  assert.equal(record.type, 'record');
  assert.match(record.value.type_name, /(^|\.)BoundaryRecord$/i);
  const fields = new Map(record.value.fields.map(([name, value]) => [name.toLowerCase(), value]));
  assert.deepEqual(fields.get('safevalue'), {type:'integer',value:42});
  assert.deepEqual(fields.get('largevalue'), {type:'integer',value:'9007199254740992'});
  assert.deepEqual(await run('DescribeBoundaryRecord'), {type:'string',value:'42|9007199254740992'});
}
{
  const session = new wasm.SpreadsheetVbaSession();
  try {
    await verifyResultBoundaries((book, source, procedure) =>
      session.run(book, source, procedure, [], 0, 'ResultBoundary.xlsm', {errorLanguage:'ja-JP'}));
  } finally { session.free(); }
}
try {
  globalThis.Worker = BrowserWorker;
  runner.closeVbaProject();
  await verifyResultBoundaries((book, source, procedure) =>
    runner.runVbaInWorker(book, source, procedure, [], 0, 'ResultBoundary.xlsm',
      {applicationState:{},errorLanguage:'ja-JP'}));
} finally { runner.closeVbaProject(); globalThis.Worker = priorWorker; }

console.log(JSON.stringify({excel_cases:cases.length,value_sequences:cases.filter(item=>item.outcome==='values').length,compile_cases:cases.filter(item=>item.outcome==='compile').length,value_steps:cases.filter(item=>item.outcome==='values').reduce((sum,item)=>sum+item.steps.length,0),direct_and_worker:true,passed:true}));
