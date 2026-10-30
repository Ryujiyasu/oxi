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
const cases = JSON.parse(await readFile(path.join(fixtures, 'chart_documents.json'), 'utf8'));
const positives = cases.filter(item => item.outcome === 'values');
const executionCases = positives.flatMap(item => [item, { ...item, without_origins: true }]);
const templateBytes = new Uint8Array(await readFile(path.join(fixtures, 'chart_document_save_refresh.xlsm')));
const bookFor = item => {
  const book = wasm.parse_spreadsheet(templateBytes);
  if (item.without_origins) for (const sheet of book.sheets) sheet.origin_id = null;
  return book;
};
const compileOptions = item => ({ errorLanguage: 'ja-JP', documentModules: item.documents.map(document => ({
  name: document.name, baseType: document.kind === 'chart' ? 'Excel.Chart' : document.kind === 'sheet' ? 'Excel.Worksheet' : 'Excel.Workbook',
})) });
for (const item of cases.filter(item => item.outcome === 'compile')) {
  const source = await readFile(path.join(fixtures, item.source), 'utf8');
  const session = new wasm.SpreadsheetVbaSession();
  try {
    const compiled = session.compile(source, compileOptions(item));
    assert.deepEqual(compiled, wasm.compile_spreadsheet_vba(source, compileOptions(item)));
    assert.equal(compiled.compiled, false, item.name);
    assert.ok(compiled.error.message.endsWith(item.message), item.name);
  } finally { session.free(); }
}
const failures = [];
const check = (result, step, name) => {
  if (result.error || result.result?.type !== 'string' || result.result.value !== step.value)
    failures.push({ name, expected: step.value, result: result.result, error: result.error });
  if (JSON.stringify(result.workbook.sheets.map(sheet => sheet.name)) !== JSON.stringify(step.sheet_names_after))
    failures.push({ name, expected_sheet_names: step.sheet_names_after, sheet_names: result.workbook.sheets.map(sheet => sheet.name) });
};
const optionsAt = (item, index) => {
  if (index && item.name === 'omit_metadata') return { errorLanguage: 'ja-JP' };
  const documents = structuredClone(index && item.name !== 'move_original_context' ? item.steps[index].metadata : item.metadata_before);
  if (index && item.name === 'metadata_order') documents.reverse();
  if (index && item.name === 'metadata_alias') for (const document of documents) document.baseType = document.baseType.split('.').pop();
  if (index && item.name === 'metadata_case') for (const document of documents) { document.name = document.name.toUpperCase(); document.baseType = document.baseType.toUpperCase(); }
  return { errorLanguage: 'ja-JP', documentModules: documents };
};
for (const item of executionCases) {
  const source = await readFile(path.join(fixtures, item.source), 'utf8');
  const options = optionsAt(item, 0);
  check(wasm.run_spreadsheet_vba_with_outcome(bookFor(item), source, item.procedure,
    [], item.host_context.active_sheet, item.host_context.book_name, options), item.steps[0], item.name);
  const session = new wasm.SpreadsheetVbaSession();
  try {
    let book = bookFor(item), active = item.host_context.active_sheet;
    for (const [index, step] of item.steps.entries()) {
      if (index) {
        if (index === 1 && item.after_first.save) {
          const saved = session.prepare_save(book, source, active, item.host_context.book_name, {});
          assert.equal(saved.error, undefined, item.name); book = saved.workbook; active = saved.active_sheet;
        }
        assert.deepEqual(session.compile(source, { documentModules: item.metadata_before }), { compiled: true });
        if (index === 1 && item.after_first.move) book.sheets.reverse();
      }
      const result = session.run(book, source, item.procedure, [], active,
        item.host_context.book_name, optionsAt(item, index));
      check(result, step, item.name);
      book = result.workbook; active = result.active_sheet;

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
  for (const item of cases.filter(item => item.outcome === 'compile')) {
    const source = await readFile(path.join(fixtures, item.source), 'utf8');
    await assert.rejects(runner.compileVbaInWorker(source, compileOptions(item)), error => error.name === 'VbaCompileError' && error.message.endsWith(item.message), item.name);
  }
  for (const item of executionCases) {
    runner.closeVbaProject();
    const source = await readFile(path.join(fixtures, item.source), 'utf8');
    let book = bookFor(item), active = item.host_context.active_sheet;
    for (const [index, step] of item.steps.entries()) {
      if (index) {
        if (index === 1 && item.after_first.save) {
          const saved = await runner.prepareVbaSaveInWorker(book, source, active, item.host_context.book_name, {});
          assert.equal(saved.error, undefined, item.name); book = saved.workbook; active = saved.active_sheet;
        }
        assert.deepEqual(await runner.compileVbaInWorker(source, { documentModules: item.metadata_before }), { compiled: true });
        if (index === 1 && item.after_first.move) book.sheets.reverse();
      }
      const result = await runner.runVbaInWorker(book, source, item.procedure, [], active,
        item.host_context.book_name, optionsAt(item, index));
      check(result, step, item.name);
      book = result.workbook; active = result.active_sheet;
    }
  }
  for (const item of positives.filter(item => item.package_fixture).flatMap(item => [item, { ...item, reparse: true }])) {
    runner.closeVbaProject();
    const bytes = new Uint8Array(await readFile(path.join(fixtures, item.package_fixture)));
    const source = wasm.spreadsheet_vba_source(bytes);
    const descriptors = wasm.spreadsheet_vba_document_modules(bytes);
    assert.ok(descriptors.some(document => document.name === "PlotHost" && document.baseType === "Excel.Chart"));
    assert.ok(descriptors.some(document => document.name === item.documents[0].name));
    assert.equal(descriptors.find(document => document.name === item.documents[0].name).codeName,
      item.metadata_before.find(document => document.name === item.documents[0].name).codeName);
    for (const descriptor of descriptors) {
      assert.equal(descriptor.sheetIndex, item.metadata_before.find(document => document.name === descriptor.name).sheetIndex, item.name);
    }
    const session = new wasm.SpreadsheetVbaSession();
    try { assert.deepEqual(session.document_modules(bytes), descriptors); }
    finally { session.free(); }
    let book = wasm.parse_spreadsheet(bytes), currentBytes = bytes, active = item.host_context.active_sheet;
    for (const [index, step] of item.steps.entries()) {
      if (index === 1) {
        const saved = await runner.prepareVbaSaveInWorker(book, source, active, item.host_context.book_name, {}, currentBytes);
        assert.equal(saved.error, undefined, item.name); book = saved.workbook; active = saved.active_sheet;
        currentBytes = wasm.edit_xlsx_from_workbook(bytes, book);
        if (item.reparse) book = wasm.parse_spreadsheet(currentBytes);
        const referenceBytes = new Uint8Array(await readFile(path.join(fixtures, item.package_refreshed)));
        assert.deepEqual(wasm.spreadsheet_vba_document_modules(currentBytes), wasm.spreadsheet_vba_document_modules(referenceBytes), item.name);
      }
      const result = await runner.runVbaInWorker(book, source, item.procedure, [], active,
        item.host_context.book_name, { errorLanguage: 'ja-JP' }, currentBytes);
      check(result, step, 'package:' + item.name);
      book = result.workbook; active = result.active_sheet;
    }
  }
} finally { runner.closeVbaProject(); globalThis.Worker = priorWorker; }
console.log(JSON.stringify({ chart_cases: cases.length, runtime_outputs: positives.reduce((count, item) => count + item.steps.length, 0),
  free_session_worker: true, representation_variants: 2, package_save_refresh_sequences: 4, mismatches: failures, passed: failures.length === 0 }));
assert.deepEqual(failures, []);
