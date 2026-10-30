// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { Worker as ThreadWorker } from 'node:worker_threads';
import { bindVbaEditorSettings } from './vba-editor-settings.js';
import * as runner from './vba-runner.js';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const pkg = path.resolve(process.argv[2] || path.join(root, 'crates/oxidocs-wasm/pkg'));
const bindingUrl = 'data:text/javascript;base64,' + Buffer.from(await readFile(path.join(pkg, 'oxidocs_wasm.js'))).toString('base64');
const wasm = await import(bindingUrl);
await wasm.default({ module_or_path: await readFile(path.join(pkg, 'oxidocs_wasm_bg.wasm')) });
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
let requests = 0;
class BrowserWorker {
  constructor() { this.thread = new ThreadWorker(new URL('data:text/javascript;base64,' + Buffer.from(bridge).toString('base64')), { workerData: workerSource }); }
  addEventListener(type, handler) { this.thread.on(type, value => handler(type === 'message' ? { data: value } : value)); }
  postMessage(value) { requests++; this.thread.postMessage(value); }
  terminate() { void this.thread.terminate(); }
}
class Control extends EventTarget {
  value = ''; disabled = false;
  fire(type) { this.dispatchEvent(new Event(type)); }
}
const fields = () => Object.fromEntries(['hostCodePage', 'errorLanguage', 'helpFile', 'sourceInput', 'loadSourceButton'].map(key => [key, new Control()]));
const storage = new Map();
const previousStorage = Object.getOwnPropertyDescriptor(globalThis, 'localStorage');
const previousWorker = globalThis.Worker;
const previousNavigator = Object.getOwnPropertyDescriptor(globalThis, 'navigator');
Object.defineProperty(globalThis, 'navigator', { configurable: true, value: { language: 'ja-JP' } });
Object.defineProperty(globalThis, 'localStorage', { configurable: true, value: { getItem: key => storage.get(key), setItem: (key, value) => storage.set(key, value) } });
globalThis.Worker = BrowserWorker;
try {
  const controls = fields();
  const packageBytes = new Uint8Array(await readFile(path.join(root, 'tests/fixtures/vba/named_project_ascii.xlsm')));
  let readMode = 'project', reads = 0, resets = 0, lastReport, sourceCodePage;
  const controller = bindVbaEditorSettings({ ...controls,
    readProjectSource: codePage => {
      sourceCodePage = codePage;
      reads++;
      if (readMode === 'none') return '';
      return wasm.spreadsheet_vba_source(readMode === 'corrupt' ? new Uint8Array([0, 1, 2]) : packageBytes, codePage);
    },
    onSourceLoaded: (before, after) => runner.recordVbaSourceChange(before, after),
    onProfileChange: () => { resets++; runner.closeVbaProject(); },
    report: (text, bad) => { lastReport = { text, bad }; },
    messages: { noSource: () => 'No project', loaded: () => 'Loaded' },
  });
  assert.deepEqual(controller.options(), { errorLanguage: 'en-US' });
  controls.sourceInput.value = 'edited draft';
  controls.loadSourceButton.fire('click');
  assert.equal(controls.sourceInput.value, wasm.spreadsheet_vba_source(packageBytes));
  assert.equal(lastReport.bad, false);
  assert.equal(requests, 0, 'reading source must not start compilation or execution');
  for (const mode of ['none', 'corrupt']) {
    readMode = mode; controls.sourceInput.value = 'keep edits';
    controls.loadSourceButton.fire('click');
    assert.equal(controls.sourceInput.value, 'keep edits');
    assert.equal(lastReport.bad, mode === 'corrupt');
    assert.equal(requests, 0);
  }
  const endFirst = controller.beginOperation(), endSecond = controller.beginOperation();
  const beforeReads = reads, beforeResets = resets;
  controls.loadSourceButton.fire('click'); controls.hostCodePage.fire('change');
  assert.equal(reads, beforeReads); assert.equal(resets, beforeResets);
  endFirst(); endFirst();
  assert.equal(controls.hostCodePage.disabled, true, 'the other operation is still running');
  endSecond();
  assert.equal(controls.loadSourceButton.disabled, false);

  const fixtures = path.join(root, 'tests/fixtures/vba');
  const item = JSON.parse(await readFile(path.join(fixtures, 'automation_names.json'), 'utf8')).find(row => row.name === 'ansi_scope_callbyname_field');
  controls.sourceInput.value = await readFile(path.join(fixtures, item.source), 'utf8');
  const workbook = () => ({ sheets: [{ name: item.host_context.sheets[0].name, rows: [], col_count: 0, col_widths: [], default_col_width: 0, default_row_height: 15, merge_cells: [] }] });
  for (const [page, expected] of [['', '438|'], ['932', '0|field'], ['1252', '438|'], ['932', '0|field'], ['', '438|']]) {
    const before = controls.sourceInput.value;
    controls.hostCodePage.value = page; controls.hostCodePage.fire('change');
    assert.equal(controls.sourceInput.value, before);
    const result = await runner.runVbaInWorker(workbook(), before, item.procedure, [], 0, 'Editor.xlsm',
      { ...controller.options(), documentModules: item.document_modules });
    assert.equal(result.error, undefined); assert.equal(result.result.value, expected);
  }
  controls.sourceInput.value = `Public Function EditorError() As String
On Error Resume Next
Err.Raise 13
EditorError = Err.Description & "|" & Err.HelpFile
End Function`;
  controls.hostCodePage.value = '932'; controls.hostCodePage.fire('change');
  controls.loadSourceButton.fire('click');
  assert.equal(sourceCodePage, 932, 'source import uses the same host encoding as execution and saving');
  controls.sourceInput.value = `Public Function EditorError() As String
On Error Resume Next
Err.Raise 13
EditorError = Err.Description & "|" & Err.HelpFile
End Function`;
  controls.helpFile.value = 'editor-help.chm'; controls.helpFile.fire('change');
  for (const [language, expected] of [['en-US', 'Type mismatch|editor-help.chm'], ['ja-JP', '型が一致しません。|editor-help.chm'], ['', 'Type mismatch|']]) {
    controls.errorLanguage.value = language;
    if (!language) controls.helpFile.value = '';
    controls.errorLanguage.fire('change');
    assert.equal(controller.options().hostAnsiCodePage, 932, 'error language must not choose the Windows character set');
    const result = await runner.runVbaInWorker(workbook(), controls.sourceInput.value, 'EditorError', [], 0, 'Editor.xlsm', controller.options());
    assert.equal(result.error, undefined); assert.equal(result.result.value, expected);
  }
  const restored = fields();
  const restoring = bindVbaEditorSettings({ ...restored, readProjectSource() { assert.fail('restoring settings must not read source'); },
    onSourceLoaded() {}, onProfileChange() { assert.fail('restoring settings must not reset an active project'); }, report() {}, messages: {} });
  assert.deepEqual(restoring.options(), { hostAnsiCodePage: 932, errorLanguage: 'en-US' });
  for (const key of storage.keys()) storage.set(key, JSON.stringify({ hostCodePage: '65001', errorLanguage: 'ja-JP', helpFile: '' }));
  const invalid = fields();
  const fallback = bindVbaEditorSettings({ ...invalid, readProjectSource() {}, onSourceLoaded() {}, onProfileChange() {}, report() {}, messages: {} });
  assert.deepEqual(fallback.options(), { errorLanguage: 'en-US' }, 'invalid stored profiles must leave the default controls intact');
  console.log(JSON.stringify({ source_import_cases: 3, overlapping_operation_cases: 1, profile_transitions: 5, error_environment_cases: 3, preference_cases: 2, production_worker: true, passed: true }));
} finally {
  runner.closeVbaProject(); globalThis.Worker = previousWorker;
  if (previousNavigator) Object.defineProperty(globalThis, 'navigator', previousNavigator);
  else delete globalThis.navigator;
  if (previousStorage) Object.defineProperty(globalThis, 'localStorage', previousStorage);
  else delete globalThis.localStorage;
}
