// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { Worker as ThreadWorker } from 'node:worker_threads';
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const fixtures = path.join(root, 'tests/fixtures/vba');
const pkg = path.resolve(process.argv[2] || path.join(root, 'crates/oxidocs-wasm/pkg'));
const bindingUrl = 'data:text/javascript;base64,' + Buffer.from(await readFile(path.join(pkg, 'oxidocs_wasm.js'))).toString('base64');
const wasm = await import(bindingUrl);
await wasm.default({ module_or_path: await readFile(path.join(pkg, 'oxidocs_wasm_bg.wasm')) });
const cases = JSON.parse(await readFile(path.join(fixtures, 'document_runtime.json'), 'utf8'));
const positives = cases.filter(item => item.outcome === 'values');
const bookFor = item => ({ sheets: item.host_context.sheets.map((sheet, index) => ({
  name: sheet.name, origin_id: index, rows: [], col_count: 0, col_widths: [],
  default_col_width: 0, default_row_height: 15, merge_cells: [],
})) });
const check = (result, step, name) => {
  assert.equal(result.error, undefined, name);
  assert.deepEqual(result.result, { type: 'string', value: step.value }, name);
};
for (const item of positives) {
  const source = await readFile(path.join(fixtures, item.source), 'utf8');
  const options = { errorLanguage: 'ja-JP', ...item.runtime_options };
  check(wasm.run_spreadsheet_vba_with_outcome(bookFor(item), source, item.procedure,
    [], item.host_context.active_sheet, item.host_context.book_name, options), item.steps[0], item.name);
  const session = new wasm.SpreadsheetVbaSession();
  try {
    let book = bookFor(item), active = item.host_context.active_sheet;
    for (const [index, step] of item.steps.entries()) {
      if (index) {
        assert.deepEqual(session.compile(source, item.compilation_options), { compiled: true });
        if (item.rename_after_first) book.sheets[0].name = item.rename_after_first;
      }
      const result = session.run(book, source, item.procedure, [], active,
        item.host_context.book_name, index ? { errorLanguage: 'ja-JP' } : options);
      check(result, step, item.name);
      book = result.workbook; active = result.active_sheet;
      if (!index && item.name === 'sheet_ticks') {
        assert.throws(() => session.run(book, source, item.procedure, [], active,
          item.host_context.book_name, { documentModules: [{ name: 'Ledger', baseType: 'Excel.Worksheet', sheetIndex: 99 }] }));
      }
    }
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
const priorWorker = globalThis.Worker;
try {
  globalThis.Worker = BrowserWorker;
  for (const item of positives) {
    runner.closeVbaProject();
    const source = await readFile(path.join(fixtures, item.source), 'utf8');
    let book = bookFor(item), active = item.host_context.active_sheet;
    for (const [index, step] of item.steps.entries()) {
      if (index) {
        assert.deepEqual(await runner.compileVbaInWorker(source, item.compilation_options), { compiled: true });
        if (item.rename_after_first) book.sheets[0].name = item.rename_after_first;
      }
      const result = await runner.runVbaInWorker(book, source, item.procedure, [], active,
        item.host_context.book_name, index ? { errorLanguage: 'ja-JP' } : { errorLanguage: 'ja-JP', ...item.runtime_options });
      check(result, step, item.name);
      book = result.workbook; active = result.active_sheet;
    }
  }
  for (const item of positives.filter(item => item.package_fixture)) {
    runner.closeVbaProject();
    const bytes = new Uint8Array(await readFile(path.join(fixtures, item.package_fixture)));
    const source = wasm.spreadsheet_vba_source(bytes);
    const descriptors = wasm.spreadsheet_vba_document_modules(bytes);
    assert.ok(descriptors.some(document => document.name === item.document_name));
    assert.equal(descriptors.find(document => document.name === item.document_name).codeName,
      item.runtime_options.documentModules[0].codeName);
    const session = new wasm.SpreadsheetVbaSession();
    try { assert.deepEqual(session.document_modules(bytes), descriptors); }
    finally { session.free(); }
    let book = wasm.parse_spreadsheet(bytes);
    for (const step of item.steps) {
      const result = await runner.runVbaInWorker(book, source, item.procedure, [], 0,
        item.host_context.book_name, { errorLanguage: 'ja-JP' }, bytes);
      check(result, step, item.name);
      book = result.workbook;
    }
  }
} finally { runner.closeVbaProject(); globalThis.Worker = priorWorker; }
console.log(JSON.stringify({ document_cases: cases.length, runtime_sequences: positives.length,
  runtime_outputs: positives.reduce((count, item) => count + item.steps.length, 0),
  free_session_worker: true, package_cases: 2, invalid_context_preserves_state: true, passed: true }));
