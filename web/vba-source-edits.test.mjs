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
const emptyBook = () => ({ sheets: [{ name: 'Sheet1', rows: [], col_count: 0, col_widths: [], default_col_width: 0, default_row_height: 15, merge_cells: [] }] });
const fixtures = [];
for (const name of ['standard', 'classes', 'controls']) {
  fixtures.push({ name, source: await readFile(path.join(root, 'tests/fixtures/vba', name === 'classes' ? 'source_edit_classes.bas' : 'source_edit_standard.bas'), 'utf8'),
    cases: JSON.parse(await readFile(path.join(root, `tests/fixtures/vba/source_edit_${name}.json`), 'utf8')), read: name === 'classes' ? 'ReadState' : 'ReadGlobals' });
}
const standard = fixtures[0].source;
const unicodePrefixes = [
  'VERSION 5.00\nBegin VB.UserForm Form1\n Caption = "日本語の文字列"\nEnd\nAttribute VB_Name = "Form1"\n',
  '#If False Then\n日本語の文字列 = "unterminated\n#End If\n',
];
async function unicodeSequence(invoke, prefix) {
  const source = prefix + 'Private stored As Long\nSub Prime()\nstored = 7\nEnd Sub\nFunction Value() As Long\nValue = stored + 1\nEnd Function\n';
  const prime = await invoke(emptyBook(), source, 'Prime', [], { applicationState: {} });
  assert.equal(prime.error, undefined);
  const before = await invoke(prime.workbook, source, 'Value', []);
  assert.equal(before.result.value, 8);
  const after = await invoke(before.workbook, source.replace('stored + 1', 'stored + 2'), 'Value', []);
  assert.equal(after.error, undefined);
  assert.equal(after.result.value, 9);
}
function updateInput(book) {
  const cell = book.sheets[0].rows.find(row => row.index === 1)?.cells.find(cell => cell.col === 0);
  if (cell) cell.value = { Number: 9 };
}
async function sequence(invoke, fixture, testCase, beforeEdit) {
  const prime = await invoke(emptyBook(), fixture.source, 'Prime', fixture.read === 'ReadGlobals' ? [0] : [], { applicationState: {} });
  assert.equal(prime.error, undefined);
  const before = await invoke(prime.workbook, fixture.source, fixture.read, []);
  assert.deepEqual(before.result, { type: 'string', value: testCase.before });
  updateInput(before.workbook);
  if (beforeEdit) beforeEdit();
  const result = await invoke(before.workbook, testCase.source, fixture.read, [], beforeEdit ? {} : { sourceEdits: testCase.edits });
  assert.equal(result.error, undefined, testCase.name);
  assert.deepEqual(result.result, { type: 'string', value: testCase.after }, testCase.name);
}
for (const fixture of fixtures) {
  for (const testCase of fixture.cases) {
    const session = new wasm.SpreadsheetVbaSession();
    try { await sequence((book, source, procedure, args, options) => session.run(book, source, procedure, args, 0, undefined, options), fixture, testCase); }
    finally { session.free(); }
  }
}
for (const prefix of unicodePrefixes) {
  const session = new wasm.SpreadsheetVbaSession();
  try { await unicodeSequence((book, source, procedure, args, options) => session.run(book, source, procedure, args, 0, undefined, options), prefix); }
  finally { session.free(); }
}
// Final source alone identifies body edits and insertions that do not touch
// persistent declarations. Explicit editing history covers the other cases.
for (const [edited, expected] of [
  ["' standalone comment\n" + standard, '123|False|4|A1:9'],
  [standard.replace('Counter = stored', 'Counter = stored + 10'), '123|False|14|A1:9'],
  [standard.replace('Static stored As Long', 'Static stored As Long\nDim temporary As Long'), '123|False|4|A1:9'],
]) {
  const session = new wasm.SpreadsheetVbaSession();
  try {
    const prime = session.run(emptyBook(), standard, 'Prime', [0], 0);
    const before = session.run(prime.workbook, standard, 'ReadGlobals', [], 0);
    updateInput(before.workbook);
    assert.equal(session.run(before.workbook, edited, 'ReadGlobals', [], 0).result.value, expected);
  } finally { session.free(); }
}
const invalidSession = new wasm.SpreadsheetVbaSession();
try {
  let previous = invalidSession.run(emptyBook(), standard, 'Prime', [0], 0);
  let count = 2;
  for (const sourceEdits of [42, {}, [null], [[]], [{ line: 0, deleteLines: 1, text: '' }],
    [{ line: 999999, deleteLines: 0, text: '' }], [{ line: 1.5, deleteLines: 1, text: '' }],
    [{ line: -1, deleteLines: 1, text: '' }], [{ line: 1, deleteLines: -1, text: '' }],
    [{ line: 1, deleteLines: 1, text: '', extra: true }], [{ line: 1, deleteLines: 1, text: 3 }],
    [{ line: 1, deleteLines: 1, text: 'Private wrong As Long' }]]) {
    assert.throws(() => invalidSession.run(previous.workbook, standard, 'ReadGlobals', [], 0, undefined, { sourceEdits }), /invalid VBA/);
    previous = invalidSession.run(previous.workbook, standard, 'ReadGlobals', [], 0);
    assert.equal(previous.result.value, `123|False|${++count}|A1:5`);
  }
  const sameModule = fixtures[2].cases.find(row => row.name === 'same_module');
  assert.throws(() => invalidSession.run(previous.workbook, standard, 'ReadGlobals', [], 0, undefined,
    { sourceEdits: sameModule.edits, applicationState: { calculation: 999 } }), /invalid VBA/);
  previous = invalidSession.run(previous.workbook, standard, 'ReadGlobals', [], 0);
  assert.equal(previous.result.value, `123|False|${++count}|A1:5`);
  const prepared = invalidSession.prepare_save(previous.workbook, '"unterminated unused source', 0, undefined,
    { applicationState: { calculation: -4135, calculateBeforeSave: false } });
  previous = invalidSession.run(prepared.workbook, standard, 'ReadGlobals', [], 0);
  assert.equal(previous.result.value, `123|False|${++count}|A1:5`);
} finally { invalidSession.free(); }

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
  const invoke = (book, source, procedure, args, options = {}) => runner.runVbaInWorker(book, source, procedure, args, 0, undefined, options);
  for (const prefix of unicodePrefixes) {
    runner.closeVbaProject();
    await unicodeSequence(invoke, prefix);
  }
  for (const fixture of fixtures) {
    for (const testCase of fixture.cases) {
      runner.closeVbaProject();
      await sequence(invoke, fixture, testCase);
    }
  }
  for (const testCase of fixtures[2].cases) {
    runner.closeVbaProject();
    await sequence(invoke, fixtures[2], testCase, () => {
      const start = testCase.old ? standard.indexOf(testCase.old) : undefined;
      const selection = start !== undefined ? { start, end: start + testCase.old.length } : undefined;
      assert.equal(runner.recordVbaSourceChange(standard, testCase.source, selection), true, testCase.name);
    });
  }
  for (const declaration of [true, false]) {
    runner.closeVbaProject();
    const prime = await invoke(emptyBook(), standard, 'Prime', [0], { applicationState: {} });
    const before = await invoke(prime.workbook, standard, 'ReadGlobals', []);
    const temporary = declaration ? standard.replace('moduleValue As Long', 'moduleValue As Double')
      : standard.replace('Counter = stored', 'Counter = Len("');
    assert.equal(runner.recordVbaSourceChange(standard, temporary), true);
    assert.equal(runner.recordVbaSourceChange(temporary, standard), true);
    const after = await invoke(before.workbook, standard, 'ReadGlobals', []);
    assert.equal(after.result.value, declaration ? '0|True|1|none' : '123|False|4|A1:5');
  }
  runner.closeVbaProject();
  const udf = `Private factor As Long
Sub Configure()
factor = 7
Application.Calculation = xlCalculationManual
Range("A1").Value = 2
Range("B1").Formula = "=Scale(A1)"
End Sub
Function Scale(ByVal value As Double) As Double
Scale = factor * value
End Function`;
  const configured = await invoke(emptyBook(), udf, 'Configure', [], { applicationState: {} });
  const edited = udf.replace('Scale = factor * value', 'Scale = factor * value + 1');
  assert.equal(runner.recordVbaSourceChange(udf, edited), true);
  const saved = await runner.prepareVbaSaveInWorker(configured.workbook, edited, 0);
  assert.equal(saved.error, undefined);
  assert.equal(saved.workbook.sheets[0].rows.find(row => row.index === 1).cells.find(cell => cell.col === 1).value.Number, 15);
} finally { runner.closeVbaProject(); globalThis.Worker = priorWorker; }
console.log(JSON.stringify({ excel_edit_cases: fixtures.reduce((sum, fixture) => sum + fixture.cases.length, 0), direct_and_worker: true, editor_history: true, invalid_input_preserves_project: true, edited_udf_save: true, passed: true }));
