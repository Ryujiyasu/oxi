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
const source = await readFile(path.join(root, 'tests/fixtures/vba/project_state.bas'), 'utf8');
const fixture = JSON.parse(await readFile(path.join(root, 'tests/fixtures/vba/project_state.json'), 'utf8'));
const objectSource = await readFile(path.join(root, 'tests/fixtures/vba/project_objects.bas'), 'utf8');
const objects = JSON.parse(await readFile(path.join(root, 'tests/fixtures/vba/project_objects.json'), 'utf8'));
const emptyBook = () => ({ sheets: [{ name: 'Sheet1', rows: [], col_count: 0, col_widths: [], default_col_width: 0, default_row_height: 15, merge_cells: [] }] });
const cell = (book, column) => book.sheets[0].rows.find(row => row.index === 1).cells.find(cell => cell.col === column);
async function sequence(invoke, testCase) {
  const prime = await invoke(emptyBook(), 'Prime', [testCase.mode]);
  assert.equal(Boolean(prime.error), testCase.failed);
  assert.equal(cell(prime.workbook, 3).value.String, testCase.inside);
  if (testCase.returned != null) assert.deepEqual(prime.result, { type: 'string', value: testCase.returned });
  const after = await invoke(prime.workbook, 'ReadGlobals');
  assert.deepEqual(after.result, { type: 'string', value: testCase.after });
  cell(after.workbook, 0).value = { Number: 9 };
  const again = await invoke(after.workbook, 'ReadGlobals');
  assert.deepEqual(again.result, { type: 'string', value: testCase.again });
  return again;
}
for (const testCase of fixture) {
  const session = new wasm.SpreadsheetVbaSession();
  try { await sequence((book, name, args = []) => session.run(book, source, name, args, 0), testCase); }
  finally { session.free(); }
}
async function objectSequence(invoke) {
  const seeded = wasm.run_spreadsheet_vba_with_outcome(emptyBook(), 'Sub Main()\nRange("A1").Value = 5\nEnd Sub', 'Main', [], 0).workbook;
  let result = await invoke(seeded, 'Prime');
  result = await invoke(result.workbook, 'ReadState');
  assert.equal(result.result.value, objects.first);
  cell(result.workbook, 0).value = { Number: 9 };
  result = await invoke(result.workbook, 'ReadState');
  assert.equal(result.result.value, objects.second);
  result = await invoke(result.workbook, 'ReadFile');
  assert.equal(result.result.value, objects.file);
  for (const expected of objects.random) {
    result = await invoke(result.workbook, 'NextRandom');
    assert.equal(result.result.value, expected);
  }
  result = await invoke(result.workbook, 'Release');
  assert.equal(result.result.value, objects.release);
}
const graphSession = new wasm.SpreadsheetVbaSession();
try { await objectSequence((book, procedure) => graphSession.run(book, objectSource, procedure, [], 0)); }
finally { graphSession.free(); }
const first = new wasm.SpreadsheetVbaSession(), other = new wasm.SpreadsheetVbaSession();
try {
  const prime = first.run(emptyBook(), source, 'Prime', [0], 0);
  assert.throws(() => first.run(prime.workbook, '"unterminated string', 'Prime', [], 0));
  assert.throws(() => first.run(prime.workbook, source, 'Prime', [], 5), /active sheet/);
  assert.throws(() => first.run(prime.workbook, source, 'Prime', [], 0, undefined, { applicationState: { calculation: 999 } }), /invalid VBA/);
  assert.deepEqual(first.run(prime.workbook, source, 'ReadGlobals', [], 0).result, { type: 'string', value: fixture[0].after });
  assert.deepEqual(other.run(prime.workbook, source, 'ReadGlobals', [], 0).result, { type: 'string', value: '0|True|1|none' });
} finally { first.free(); other.free(); }
const udfSource = `Private factor As Long
Sub Configure()
factor = 7
Application.Calculation = xlCalculationManual
Range("A1").Value = 2
Range("B1").Formula = "=Scale(A1)"
End Sub
Function Scale(ByVal value As Double) As Double
Scale = factor * value
End Function`;
const saveSession = new wasm.SpreadsheetVbaSession();
try {
  const configured = saveSession.run(emptyBook(), udfSource, 'Configure', [], 0);
  assert.equal(configured.error, undefined);
  const saved = saveSession.prepare_save(configured.workbook, udfSource, 0);
  assert.equal(saved.error, undefined);
  assert.equal(cell(saved.workbook, 1).value.Number, 14);
} finally { saveSession.free(); }
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
let workers = 0;
class BrowserWorker {
  constructor() { workers++; this.thread = new ThreadWorker(new URL('data:text/javascript;base64,' + Buffer.from(bridge).toString('base64')), { workerData: workerSource }); }
  addEventListener(type, handler) { this.thread.on(type, value => handler(type === 'message' ? { data: value } : value)); }
  postMessage(value) { this.thread.postMessage(value); }
  terminate() { void this.thread.terminate(); }
}
const priorWorker = globalThis.Worker;
const runner = await import(pathToFileURL(path.join(root, 'web/vba-runner.js')));
try {
  globalThis.Worker = BrowserWorker;
  for (const testCase of fixture) {
    runner.closeVbaProject();
    const countBefore = workers;
    const invoke = async (book, procedure, args = []) => {
      try { return await runner.runVbaInWorker(book, source, procedure, args, 0, 'State.xlsm', { applicationState: {} }); }
      catch (error) { if (error.execution) return error.execution; throw error; }
    };
    const again = await sequence(invoke, testCase);
    assert.equal(workers, countBefore + 1, 'all three invocations use one loaded project');
    await assert.rejects(runner.runVbaInWorker(again.workbook, source, 'ReadGlobals', [Symbol('clone failure')], 0), /could not be cloned/);
    const next = await invoke(again.workbook, 'ReadGlobals');
    assert.equal(next.result.value.split('|')[2], String(Number(testCase.again.split('|')[2]) + 1), 'structured clone failure retains project state');
    runner.closeVbaProject();
    assert.deepEqual((await invoke(again.workbook, 'ReadGlobals')).result, { type: 'string', value: '0|True|1|none' });
  }
  runner.closeVbaProject();
  await objectSequence((book, procedure) => runner.runVbaInWorker(book, objectSource, procedure, [], 0, 'Graph.xlsm'));
  runner.closeVbaProject();
  const configured = await runner.runVbaInWorker(emptyBook(), udfSource, 'Configure', [], 0);
  const saved = await runner.prepareVbaSaveInWorker(configured.workbook, udfSource, 0);
  assert.equal(cell(saved.workbook, 1).value.Number, 14);
} finally { runner.closeVbaProject(); globalThis.Worker = priorWorker; }
console.log(JSON.stringify({ excel_project_sequences: fixture.length, direct_session_and_worker: true, live_range: true, udf_globals_on_save: true, passed: true }));
