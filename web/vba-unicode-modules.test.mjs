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
const cases = JSON.parse(await readFile(path.join(fixtures, 'unicode_modules.json'), 'utf8'));
const sorted = values => [...values].sort((a, b) => a.name < b.name ? -1 : a.name > b.name ? 1 : 0);
const check = (result, item) => {
  assert.equal(result.error, undefined, item.name);
  assert.deepEqual(result.result, { type: 'string', value: item.expected }, item.name);
};
for (const item of cases) {
  const bytes = new Uint8Array(await readFile(path.join(fixtures, item.package_fixture)));
  const source = wasm.spreadsheet_vba_source(bytes, item.source_code_page);
  assert.ok(source, item.name);
  const session = new wasm.SpreadsheetVbaSession();
  try {
    const documents = session.document_modules(bytes);
    assert.deepEqual(sorted(documents), sorted(item.document_modules), item.name);
    const options = { errorLanguage: 'ja-JP', documentModules: documents };
    let book = wasm.parse_spreadsheet(bytes);
    for (let call = 0; call < 3; call++) {
      const result = session.run(book, source, item.procedure, [], item.host_context.active_sheet, item.host_context.book_name, options);
      check(result, item); book = result.workbook;
    }
    check(wasm.run_spreadsheet_vba_with_outcome(book, source, item.procedure, [], item.host_context.active_sheet, item.host_context.book_name, options), item);
    const fresh = new wasm.SpreadsheetVbaSession();
    try {
      check(fresh.run(wasm.parse_spreadsheet(bytes), wasm.spreadsheet_vba_source(bytes, item.source_code_page), item.procedure,
        [], item.host_context.active_sheet, item.host_context.book_name, { errorLanguage: 'ja-JP', documentModules: fresh.document_modules(bytes) }), item);
    } finally { fresh.free(); }
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
    const bytes = new Uint8Array(await readFile(path.join(fixtures, item.package_fixture)));
    const source = wasm.spreadsheet_vba_source(bytes, item.source_code_page);
    let book = wasm.parse_spreadsheet(bytes);
    for (let call = 0; call < 3; call++) {
      const result = await runner.runVbaInWorker(book, source, item.procedure, [], item.host_context.active_sheet,
        item.host_context.book_name, { errorLanguage: 'ja-JP' }, bytes);
      check(result, item); book = result.workbook;
    }
    const saved = await runner.prepareVbaSaveInWorker(book, '', item.host_context.active_sheet, item.host_context.book_name,
      { projectSourceCodePage: item.source_code_page }, bytes);
    assert.equal(saved.error, undefined, item.name);
  }
} finally { runner.closeVbaProject(); globalThis.Worker = previous; }
console.log(JSON.stringify({ unicode_module_cases: cases.length, excel_observations: cases.length * 3,
  runtime_outputs: cases.length * 3, worker_outputs: cases.length * 3, free_outputs: cases.length,
  fresh_outputs: cases.length, metadata_cases: cases.length, passed: true }));
