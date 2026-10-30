// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.
import assert from 'node:assert/strict';
import { readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { Worker as ThreadWorker } from 'node:worker_threads';
import vm from 'node:vm';
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const pkg = path.resolve(process.argv[2] || path.join(root, 'crates/oxidocs-wasm/pkg'));
const bindingUrl = 'data:text/javascript;base64,' + Buffer.from(await readFile(path.join(pkg, 'oxidocs_wasm.js'))).toString('base64');
const wasm = await import(bindingUrl);
await wasm.default({ module_or_path: await readFile(path.join(pkg, 'oxidocs_wasm_bg.wasm')) });
const source = await readFile(path.join(root, 'tests/fixtures/vba/cell_utf16.bas'), 'utf8');
const expected = new Map((await readFile(path.join(root, 'tests/fixtures/vba/cell_utf16.tsv'), 'utf8')).split(/\r?\n/).filter(line => line && !line.startsWith('#')).map(line => line.split('\t')));
const original = new Uint8Array(await readFile(path.join(root, 'crates/oxicells-core/tests/fixtures/hidden_rows_cols.xlsx')));
const options = { hostAnsiCodePage: 932, errorLanguage: 'ja-JP' };
const verify = (outcome, procedure) => {
  assert.equal(outcome.error, undefined, procedure + ': ' + JSON.stringify(outcome.error));
  assert.deepEqual(outcome.result, { type: 'string', value: expected.get(procedure) });
  return outcome.workbook;
};
let book = wasm.parse_spreadsheet(original);
book = verify(wasm.run_spreadsheet_vba_with_outcome(book, source, 'SetupCells', [], 0, 'utf16.xlsx', options), 'SetupCells');
const session = new wasm.SpreadsheetVbaSession();
try {
  assert.deepEqual(session.compile(source, options), { compiled: true });
  for (let call = 0; call < 3; call++) book = verify(session.run(book, source, 'ReadCells', [], 0, 'utf16.xlsx', options), 'ReadCells');
} finally { session.free(); }
const saved = wasm.edit_xlsx_from_workbook(original, book);
if (process.argv[3]) await writeFile(process.argv[3], saved);
const reopened = wasm.parse_spreadsheet(saved);
verify(wasm.run_spreadsheet_vba_with_outcome(reopened, source, 'ReadCells', [], 0, 'utf16.xlsx', options), 'ReadCells');
assert.deepEqual(reopened.sheets[0].rows[0].cells[0].value, { Utf16: [0xD83D] });
assert.deepEqual(reopened.sheets[0].rows[1].cells[0].value, { Utf16: [0xDE00] });

const index = await readFile(path.join(root, 'web/index.html'), 'utf8');
const extract = name => index.slice(index.indexOf('function ' + name + '('), index.indexOf('\n}', index.indexOf('function ' + name + '(')) + 2);
const ui = vm.createContext({});
vm.runInContext(extract('cellStringValue') + '\n' + extract('cellUtf16Text') + '\n' + extract('formatCellValue'), ui);
for (const units of [[0xD83D], [0xDE00], [0xD83D, 0xDE00], [65, 0, 0xD83D], Array(20000).fill(0xD83D)]) {
  ui.units = units;
  vm.runInContext('text = cellUtf16Text(units); stored = cellStringValue(text); shown = formatCellValue(stored)', ui);
  assert.deepEqual(Array.from({ length: ui.shown.length }, (_, at) => ui.shown.charCodeAt(at)), units);
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
  terminate() { return this.thread.terminate(); }
}
const previous = globalThis.Worker;
globalThis.Worker = BrowserWorker;
const runner = await import(pathToFileURL(path.join(root, 'web/vba-runner.js')));
try {
  assert.deepEqual(await runner.compileVbaInWorker(source, options, reopened), { compiled: true });
  verify(await runner.runVbaInWorker(reopened, source, 'ReadCells', [], 0, 'utf16.xlsx', options), 'ReadCells');
  let workerBook = wasm.parse_spreadsheet(original);
  workerBook = verify(await runner.runVbaInWorker(workerBook, source, 'SetupCells', [], 0, 'utf16.xlsx', options), 'SetupCells');
  verify(await runner.runVbaInWorker(workerBook, source, 'ReadCells', [], 0, 'utf16.xlsx', options), 'ReadCells');
} finally { runner.closeVbaProject(); globalThis.Worker = previous; }
console.log(JSON.stringify({ passed: true, native_cells: 60, transport: ['free', 'session', 'save-reopen', 'worker', 'editor'], retained_unpaired_units: true }));
