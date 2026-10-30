// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

// Exercise the real suite WASM, worker transport and workbook save after failure.
// node web/vba-failure-state.test.mjs <suite WASM package directory>
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { Worker as ThreadWorker } from 'node:worker_threads';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const pkg = path.resolve(process.argv[2] || path.join(root, 'crates/oxidocs-wasm/pkg'));
const bindings = await readFile(path.join(pkg, 'oxidocs_wasm.js'), 'utf8');
const bindingUrl = 'data:text/javascript;base64,' + Buffer.from(bindings).toString('base64');
const wasm = await import(bindingUrl);
await wasm.default({ module_or_path: await readFile(path.join(pkg, 'oxidocs_wasm_bg.wasm')) });
const source = await readFile(path.join(root, 'tests/fixtures/vba/failure_state.bas'), 'utf8');
const fixture = JSON.parse(await readFile(path.join(root, 'tests/fixtures/vba/failure_state.json'), 'utf8'));
const sheet = name => ({ name, rows: [], col_count: 0, col_widths: [], default_col_width: 0, default_row_height: 15, merge_cells: [] });
const seed = () => wasm.run_spreadsheet_vba_with_outcome({ sheets: [sheet('Sheet1'), sheet('Sheet2')] }, source, 'Seed', [], 0).workbook;
const cell = (sheet, row, col) => sheet?.rows.find(one => one.index === row)?.cells.find(one => one.col === col);
const value = cell => cell?.value === 'Empty' ? null : cell ? Object.values(cell.value)[0] : null;
function snapshot(execution) {
  const { workbook, active_sheet } = execution;
  const first = workbook.sheets[0];
  const a1 = cell(first, 1, 0), b1 = cell(first, 1, 1);
  return {
    a1: value(a1), b1: value(b1), b1_formula: '=' + b1.formula.replace(/^=/, ''),
    c1: value(cell(first, 1, 2)), bold: a1.style.bold,
    number_format: a1.style.number_format ?? 'General',
    sheets: workbook.sheets.map(sheet => sheet.name), active_sheet,
    row2_hidden: first.rows.some(row => row.index === 2 && row.hidden),
    created_d1: value(cell(workbook.sheets.find(sheet => sheet.name === 'Created'), 1, 3)),
    names: Object.fromEntries((workbook.defined_names || []).map(([name, formula]) => [name, '=' + formula.replace(/^=/, '')])),
  };
}
function check(execution, testCase) {
  const expected = { ...testCase.state };
  const { saved, calculation, status_bar } = expected;
  delete expected.saved; delete expected.calculation; delete expected.status_bar;
  assert.deepEqual(snapshot(execution), expected, testCase.procedure);
  assert.equal(Boolean(execution.error), testCase.failed, testCase.procedure);
  assert.equal(execution.workbook_changed, !saved, testCase.procedure);
  if (status_bar !== undefined) assert.equal(execution.status_bar, status_bar, testCase.procedure);
}
const results = new Map();
for (const testCase of fixture.cases) {
  const workbook = seed(), before = structuredClone(workbook);
  const execution = wasm.run_spreadsheet_vba_with_outcome(workbook, source, testCase.procedure, [], 0, 'Failure.xlsm', { errorLanguage: 'ja-JP' });
  check(execution, testCase);
  assert.deepEqual(workbook, before, 'input object is not mutated by the binding');
  results.set(testCase.procedure, execution);
  if (testCase.failed) {
    assert.throws(() => wasm.run_spreadsheet_vba_with_options(workbook, source, testCase.procedure, [], 0, undefined, { errorLanguage: 'ja-JP' }), error => error.message === execution.error.message, 'legacy API still throws the same error');
  }
}
const legacy = wasm.run_spreadsheet_vba(seed(), source, 'Success', [], 0);
assert.deepEqual(Object.keys(legacy).sort(), ['workbook', 'result', 'debug_output', 'messages', 'status_bar'].sort(), 'legacy success shape');
assert.equal(results.get('AbortFailure').error.kind, 'terminated');
assert.deepEqual(results.get('NestedFailure').debug_output, ['before', 'inner']);
for (const [procedure, changed, debug] of [['CompileEntryFailure', false, []], ['CompileCalleeFailure', true, ['before']]]) {
  const result = wasm.run_spreadsheet_vba_with_outcome(seed(), source, procedure, [], 0);
  assert.equal(result.error.kind, 'compile');
  assert.equal(result.workbook_changed, changed);
  assert.deepEqual(result.debug_output, debug);
}
assert.throws(() => wasm.run_spreadsheet_vba_with_outcome(seed(), source, 'Success', {}, 0), /invalid VBA arguments/);
assert.throws(() => wasm.run_spreadsheet_vba_with_outcome(seed(), source, 'Success', [], 0, undefined, { typo: true }), /invalid VBA options/);

// Saving is checked independently of the runtime: serialize the returned IR
// into an existing macro package, reopen it, then inspect cached values/styles.
const original = await readFile(path.join(root, 'tests/fixtures/vba/named_project_ascii.xlsm'));
for (const testCase of fixture.cases) {
  const execution = results.get(testCase.procedure);
  const bytes = wasm.edit_xlsx_from_workbook(original, execution.workbook);
  const reopened = wasm.parse_spreadsheet(bytes);
  assert.deepEqual(snapshot({ workbook: reopened, active_sheet: execution.active_sheet }), snapshot(execution), `save/reopen ${testCase.procedure}`);
  assert.equal(wasm.spreadsheet_vba_project_name(bytes), 'OxiNamedProject');
}

const shim = `import init, { SpreadsheetVbaSession, run_spreadsheet_vba_with_outcome, prepare_spreadsheet_vba_save, spreadsheet_vba_project_name, spreadsheet_vba_source } from ${JSON.stringify(bindingUrl)};
import { readFile } from 'node:fs/promises';
await init({ module_or_path: await readFile(${JSON.stringify(path.join(pkg, 'oxidocs_wasm_bg.wasm'))}) });
export { SpreadsheetVbaSession, run_spreadsheet_vba_with_outcome, prepare_spreadsheet_vba_save, spreadsheet_vba_project_name, spreadsheet_vba_source };
export default function ready() { return Promise.resolve(); }`;
const workerSource = (await readFile(path.join(root, 'web/vba-worker.js'), 'utf8')).replace("'./oxidocs_wasm.js'", JSON.stringify('data:text/javascript;base64,' + Buffer.from(shim).toString('base64')));
const bridge = `import { parentPort, workerData } from 'node:worker_threads';
const queued = []; let listener;
globalThis.self = { postMessage: value => parentPort.postMessage(value),
addEventListener: (type, handler) => { listener = handler; for (const data of queued) listener({ data }); } };
parentPort.on('message', data => listener ? listener({ data }) : queued.push(data));
await import('data:text/javascript;base64,' + Buffer.from(workerData).toString('base64'));`;
class BrowserWorker {
  constructor() { this.thread = new ThreadWorker(new URL('data:text/javascript;base64,' + Buffer.from(bridge).toString('base64')), { workerData: workerSource }); }
  addEventListener(type, handler) { this.thread.on(type, value => handler(type === 'message' ? { data: value } : value)); }
  postMessage(value) { this.thread.postMessage(value); }
  terminate() { void this.thread.terminate(); }
}
const priorWorker = globalThis.Worker;
try {
  globalThis.Worker = BrowserWorker;
  const { runVbaInWorker, closeVbaProject } = await import(pathToFileURL(path.join(root, 'web/vba-runner.js')));
  for (const testCase of fixture.cases) {
    closeVbaProject();
    let execution;
    try {
      // Each oracle row was measured in an independent application context.
      execution = await runVbaInWorker(seed(), source, testCase.procedure, [], 0, 'Failure.xlsm', { errorLanguage: 'ja-JP', applicationState: {} });
      assert.equal(testCase.failed, false);
    } catch (error) {
      assert.equal(error.name, 'VbaExecutionError');
      execution = error.execution;
      assert.equal(error.message, execution.error.message);
    }
    check(execution, testCase);
  }
} finally { (await import(pathToFileURL(path.join(root, 'web/vba-runner.js')))).closeVbaProject(); globalThis.Worker = priorWorker; }
console.log(JSON.stringify({ excel_states: fixture.cases.length, save_reopen: fixture.cases.length, worker_states: fixture.cases.length, passed: true }));
