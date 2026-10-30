// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.
import assert from 'node:assert/strict';
import { readFile, writeFile, mkdir } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { Worker as ThreadWorker } from 'node:worker_threads';
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const fixtures = path.join(root, 'tests/fixtures/vba');
const pkg = path.resolve(process.argv[2] || path.join(root, 'crates/oxidocs-wasm/pkg'));
const outputDirectory = process.argv[3];
if (outputDirectory) await mkdir(outputDirectory, { recursive: true });
const bindingUrl = 'data:text/javascript;base64,' + Buffer.from(await readFile(path.join(pkg, 'oxidocs_wasm.js'))).toString('base64');
const wasm = await import(bindingUrl);
await wasm.default({ module_or_path: await readFile(path.join(pkg, 'oxidocs_wasm_bg.wasm')) });
const cases = JSON.parse(await readFile(path.join(fixtures, 'chart_persistence.json'), 'utf8'));
const check = (result, item, index) => {
  assert.equal(result.error, undefined, item.name);
  assert.deepEqual(result.result, { type: 'string', value: item.steps[index].value }, item.name + ':' + item.steps[index].stage);
};
for (const item of cases) {
  const bytes = new Uint8Array(await readFile(path.join(fixtures, item.package_fixture)));
  const source = await readFile(path.join(fixtures, item.source), 'utf8');
  const session = new wasm.SpreadsheetVbaSession();
  let saved;
  try {
    const options = { errorLanguage: 'ja-JP', documentModules: session.document_modules(bytes) };
    let book = wasm.parse_spreadsheet(bytes), active = item.host_context.active_sheet;
    const run = procedure => {
      const result = session.run(book, source, procedure, [], active, item.host_context.book_name, options);
      book = result.workbook; active = result.active_sheet; return result;
    };
    check(run(item.procedure), item, 0);
    assert.equal(run(item.change_procedure).error, undefined, item.name);
    check(run(item.procedure), item, 1);
    const prepared = session.prepare_save(book, source, active, item.host_context.book_name, options);
    assert.equal(prepared.error, undefined, item.name);
    saved = wasm.edit_xlsx_from_workbook(bytes, prepared.workbook);
  } finally { session.free(); }
  // Read with a new project and native objects, not the writer's cached state.
  for (const packageBytes of [saved, new Uint8Array(await readFile(path.join(fixtures, item.package_refreshed)))]) {
    const reader = new wasm.SpreadsheetVbaSession();
    try {
      const book = wasm.parse_spreadsheet(packageBytes), options = { errorLanguage: 'ja-JP', documentModules: reader.document_modules(packageBytes) };
      check(reader.run(book, source, item.procedure, [], item.host_context.active_sheet, item.host_context.book_name, options), item, 2);
      check(wasm.run_spreadsheet_vba_with_outcome(book, source, item.procedure, [], item.host_context.active_sheet, item.host_context.book_name, options), item, 2);
    } finally { reader.free(); }
  }
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
  for (const item of cases) {
    runner.closeVbaProject();
    const bytes = new Uint8Array(await readFile(path.join(fixtures, item.package_fixture)));
    const source = wasm.spreadsheet_vba_source(bytes, item.source_code_page);
    let book = wasm.parse_spreadsheet(bytes), active = item.host_context.active_sheet;
    const run = async procedure => {
      const result = await runner.runVbaInWorker(book, source, procedure, [], active, item.host_context.book_name, { errorLanguage: 'ja-JP' }, bytes);
      book = result.workbook; active = result.active_sheet; return result;
    };
    check(await run(item.procedure), item, 0);
    assert.equal((await run(item.change_procedure)).error, undefined, item.name);
    check(await run(item.procedure), item, 1);
    const prepared = await runner.prepareVbaSaveInWorker(book, '', active, item.host_context.book_name, { projectSourceCodePage: item.source_code_page }, bytes);
    assert.equal(prepared.error, undefined, item.name);
    const saved = wasm.edit_xlsx_from_workbook(bytes, prepared.workbook);
    if (outputDirectory) await writeFile(path.join(outputDirectory, item.name + '.xlsm'), saved);
    runner.closeVbaProject();
    check(await runner.runVbaInWorker(wasm.parse_spreadsheet(saved), wasm.spreadsheet_vba_source(saved, item.source_code_page), item.procedure,
      [], item.host_context.active_sheet, item.host_context.book_name, { errorLanguage: 'ja-JP' }, saved), item, 2);
  }
} finally { runner.closeVbaProject(); globalThis.Worker = priorWorker; }
console.log(JSON.stringify({ persistence_cases: cases.length, excel_observations: cases.reduce((count, item) => count + item.steps.length, 0),
  fresh_session_reopen: true, free_session_worker: true, native_reference_import: true, passed: true }));
