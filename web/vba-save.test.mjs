// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { Worker as ThreadWorker } from 'node:worker_threads';
import { inflateRawSync } from 'node:zlib';
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const pkg = path.resolve(process.argv[2] || path.join(root, 'crates/oxidocs-wasm/pkg'));
const bindingUrl = 'data:text/javascript;base64,' + Buffer.from(await readFile(path.join(pkg, 'oxidocs_wasm.js'))).toString('base64');
const wasm = await import(bindingUrl);
await wasm.default({ module_or_path: await readFile(path.join(pkg, 'oxidocs_wasm_bg.wasm')) });
const source = await readFile(path.join(root, 'tests/fixtures/vba/save_calculation.bas'), 'utf8');
const fixture = JSON.parse(await readFile(path.join(root, 'tests/fixtures/vba/save_calculation.json'), 'utf8'));
const original = await readFile(path.join(root, 'tests/fixtures/vba/save_calculation.xlsm'));
function zipPart(bytes, target) {
  const zip = Buffer.from(bytes), end = zip.lastIndexOf(Buffer.from([80, 75, 5, 6]));
  let at = zip.readUInt32LE(end + 16);
  for (let i = 0; i < zip.readUInt16LE(end + 10); i++) {
    const length = zip.readUInt16LE(at + 28);
    if (zip.subarray(at + 46, at + 46 + length).toString() === target) {
      const local = zip.readUInt32LE(at + 42);
      const start = local + 30 + zip.readUInt16LE(local + 26) + zip.readUInt16LE(local + 28);
      const data = zip.subarray(start, start + zip.readUInt32LE(at + 20));
      return zip.readUInt16LE(at + 10) === 8 ? inflateRawSync(data) : data;
    }
    at += 46 + length + zip.readUInt16LE(at + 30) + zip.readUInt16LE(at + 32);
  }
  assert.fail(target);
}
const b1 = workbook => workbook.sheets[0].rows.find(row => row.index === 1).cells.find(cell => cell.col === 1).value.Number;
const fresh = () => ({ workbook: wasm.parse_spreadsheet(original) });
function invoke(previous, procedure, args = []) {
  const result = wasm.run_spreadsheet_vba_with_outcome(previous.workbook, source, procedure, args, 0, undefined, { applicationState: previous.application_state });
  assert.equal(result.error, undefined);
  return result;
}
function checkSaved(execution, testCase) {
  assert.equal(b1(execution.workbook), testCase.after);
  const bytes = wasm.edit_xlsx_from_workbook(original, execution.workbook);
  const reopened = wasm.parse_spreadsheet(bytes);
  assert.equal(b1(reopened), testCase.cached);
  const sheetXml = zipPart(bytes, 'xl/worksheets/sheet1.xml').toString('utf8');
  const cellXml = sheetXml.match(/<c\b(?=[^>]*\br="B1")[^>]*>([\s\S]*?)<\/c>/)[1];
  assert.equal(Number(cellXml.match(/<v>([^<]*)<\/v>/)[1]), testCase.cached, 'stored cache is present without parser recalculation');
  assert.equal(cellXml.match(/<f\b[^>]*>([^<]*)<\/f>/)[1], testCase.formula);
  const xml = zipPart(bytes, 'xl/workbook.xml').toString('utf8');
  const tag = xml.match(/<(?:\w+:)?calcPr\b[^>]*\/?\s*>/)[0];
  const attributes = Object.fromEntries([...tag.matchAll(/(\w+)="([^"]*)"/g)].map(([, key, value]) => [key, value]));
  delete attributes.calcId;
  assert.deepEqual(attributes, testCase.calcPr, 'actual serialized calcPr attributes match Excel');
  assert.deepEqual(zipPart(bytes, 'xl/vbaProject.bin'), zipPart(original, 'xl/vbaProject.bin'), 'original VBA source and compiled streams survive Save');
}
for (const testCase of fixture.saves) {
  const reset = invoke(fresh(), 'Reset', [testCase.udf]);
  const configured = invoke(reset, 'Configure', [testCase.initial, testCase.target]);
  assert.equal(b1(configured.workbook), testCase.before);
  const saved = wasm.prepare_spreadsheet_vba_save(configured.workbook, source, 0, undefined, { applicationState: configured.application_state });
  assert.equal(saved.error, undefined);
  checkSaved(saved, testCase);
}
for (const [rows, procedure] of [[fixture.modes, 'ModeOnly'], [fixture.on_save, 'OnSaveOnly']]) {
  for (const testCase of rows) {
    let execution = invoke(fresh(), 'Reset', [false]);
    execution = invoke(execution, procedure, [testCase.initial]);
    execution = invoke(execution, procedure, [testCase.target]);
    assert.equal(execution.workbook_changed, !testCase.saved, `${procedure} ${testCase.initial}->${testCase.target}`);
  }
}
const importedSource = wasm.spreadsheet_vba_source(original);
assert.ok(importedSource.includes('Function Twice'));
assert.ok(importedSource.includes('VERSION 1.0 CLASS'));
assert.deepEqual(Buffer.from(wasm.save_spreadsheet_vba_source(original, importedSource)), original);
const editedSource = importedSource.replace('Twice = amount * 2', 'Twice = amount * 3');
const sourceSaved = wasm.save_spreadsheet_vba_source(original, editedSource);
const sourceReloaded = wasm.spreadsheet_vba_source(sourceSaved);
assert.ok(sourceReloaded.includes('Twice = amount * 3'));
assert.equal(wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(sourceSaved), sourceReloaded, 'Twice', [7]).result.value, 21);
assert.deepEqual(wasm.spreadsheet_vba_document_modules(sourceSaved), wasm.spreadsheet_vba_document_modules(original));
const unicodeCases = JSON.parse(await readFile(path.join(root, 'tests/fixtures/vba/unicode_modules.json'), 'utf8'));
for (const item of unicodeCases) {
  const bytes = await readFile(path.join(root, 'tests/fixtures/vba', item.package_fixture));
  const text = wasm.spreadsheet_vba_source(bytes, item.source_code_page);
  assert.deepEqual(Buffer.from(wasm.save_spreadsheet_vba_source(bytes, text, item.source_code_page)), bytes, item.name);
}
const shim = `import init, { SpreadsheetVbaSession, run_spreadsheet_vba_with_outcome, prepare_spreadsheet_vba_save, spreadsheet_vba_project_name, spreadsheet_vba_source } from ${JSON.stringify(bindingUrl)};
import { readFile } from 'node:fs/promises';
await init({ module_or_path: await readFile(${JSON.stringify(path.join(pkg, 'oxidocs_wasm_bg.wasm'))}) });
export { SpreadsheetVbaSession, run_spreadsheet_vba_with_outcome, prepare_spreadsheet_vba_save, spreadsheet_vba_project_name, spreadsheet_vba_source };
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
const priorWorker = globalThis.Worker;
try {
  globalThis.Worker = BrowserWorker;
  const { runVbaInWorker, prepareVbaSaveInWorker } = await import(pathToFileURL(path.join(root, 'web/vba-runner.js')));
  for (const testCase of fixture.saves) {
    const reset = await runVbaInWorker(fresh().workbook, source, 'Reset', [testCase.udf], 0, 'Save.xlsm', { applicationState: null }, original);
    const configured = await runVbaInWorker(reset.workbook, source, 'Configure', [testCase.initial, testCase.target], 0, 'Save.xlsm', {}, original);
    assert.equal(b1(configured.workbook), testCase.before);
    // Empty caller source must import the workbook's real modules for UDFs.
    const saved = await prepareVbaSaveInWorker(configured.workbook, '', 0, 'Save.xlsm', {}, original);
    checkSaved(saved, testCase);
  }
  const manual = invoke(invoke(fresh(), 'Reset', [true]), 'Configure', [-4135, true]);
  const loadedBytes = wasm.edit_xlsx_from_workbook(original, manual.workbook);
  const loaded = wasm.parse_spreadsheet(loadedBytes);
  assert.equal(loaded.calculation.mode, 'manual');
  const prepared = await prepareVbaSaveInWorker(loaded, '', 0, 'Imported.xlsm', { applicationState: null }, loadedBytes);
  assert.equal(b1(prepared.workbook), 14, 'first application state is read from imported calculation metadata');
  const noCalc = invoke(invoke(fresh(), 'Reset', [true]), 'Configure', [-4135, false]);
  const uncompiled = await prepareVbaSaveInWorker(noCalc.workbook, 'invalid source (', 0, 'Save.xlsm', { applicationState: noCalc.application_state }, original);
  assert.equal(b1(uncompiled.workbook), 2, 'manual Save without calculation does not compile unused source');
} finally { (await import(pathToFileURL(path.join(root, 'web/vba-runner.js')))).closeVbaProject(); globalThis.Worker = priorWorker; }
console.log(JSON.stringify({ save_cases: fixture.saves.length, mode_assignments: fixture.modes.length, on_save_assignments: fixture.on_save.length, wasm_and_worker: true, passed: true }));
