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
const fixtures = path.join(root, 'tests/fixtures/vba');
const cases = JSON.parse(await readFile(path.join(fixtures, 'callbyname_binding.json'), 'utf8'));
const book = item => ({ sheets: [{ name: item.host_context.sheets[0].name, rows: [], col_count: 0, col_widths: [],
  default_col_width: 0, default_row_height: 15, merge_cells: [] }] });
const options = item => ({ errorLanguage: 'ja-JP', hostAnsiCodePage: 932, documentModules: item.document_modules });
const check = (result, item) => {
  assert.equal(result.error, undefined, item.name);
  assert.deepEqual(result.result, { type: 'string', value: item.expected }, item.name);
};
for (const item of cases) {
  const source = await readFile(path.join(fixtures, item.source), 'utf8');
  const session = new wasm.SpreadsheetVbaSession();
  try {
    let workbook = book(item);
    for (let call = 0; call < 3; call++) {
      const settings = call ? { errorLanguage: 'en-US', documentModules: item.document_modules } : options(item);
      const result = session.run(workbook, source, item.procedure, [], 0, item.host_context.book_name, settings);
      check(result, item); workbook = result.workbook;
    }
    check(wasm.run_spreadsheet_vba_with_outcome(book(item), source, item.procedure, [], 0,
      item.host_context.book_name, options(item)), item);
    const fresh = new wasm.SpreadsheetVbaSession();
    try { check(fresh.run(book(item), source, item.procedure, [], 0, item.host_context.book_name, options(item)), item); }
    finally { fresh.free(); }
  } finally { session.free(); }
}
const shim = `import init, { SpreadsheetVbaSession, spreadsheet_vba_project_name, spreadsheet_vba_source } from ${JSON.stringify(bindingUrl)};
import { readFile } from 'node:fs/promises';
await init({ module_or_path: await readFile(${JSON.stringify(path.join(pkg, 'oxidocs_wasm_bg.wasm'))}) });
export { SpreadsheetVbaSession, spreadsheet_vba_project_name, spreadsheet_vba_source };
export default function ready() { return Promise.resolve(); }`;
const workerSource = (await readFile(path.join(root, 'web/vba-worker.js'), 'utf8'))
  .replace("'./oxidocs_wasm.js'", JSON.stringify('data:text/javascript;base64,' + Buffer.from(shim).toString('base64')));
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
const previous = globalThis.Worker;
try {
  globalThis.Worker = BrowserWorker;
  for (const item of cases) {
    runner.closeVbaProject();
    const source = await readFile(path.join(fixtures, item.source), 'utf8');
    assert.deepEqual(await runner.compileVbaInWorker(source, { ...item.compilation_options, hostAnsiCodePage: 932 }, book(item)), { compiled: true });
    let workbook = book(item);
    for (let call = 0; call < 3; call++) {
      const settings = call ? { errorLanguage: 'en-US', documentModules: item.document_modules } : options(item);
      const result = await runner.runVbaInWorker(workbook, source, item.procedure, [], 0, item.host_context.book_name, settings);
      check(result, item); workbook = result.workbook;
    }
    const saved = await runner.prepareVbaSaveInWorker(workbook, source, 0, item.host_context.book_name,
      { documentModules: item.document_modules });
    assert.equal(saved.error, undefined, item.name);
  }
} finally { runner.closeVbaProject(); globalThis.Worker = previous; }
console.log(JSON.stringify({ callbyname_projects: cases.length, binding_cases: cases.reduce((n, item) => n + item.measured_cases, 0), direct_outputs: cases.length * 3,
  worker_outputs: cases.length * 3, free_outputs: cases.length, fresh_outputs: cases.length,
  worker_compile_cases: cases.length, passed: true }));
