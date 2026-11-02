// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

// Main editor macro/error/save handlers with real suite WASM and a small DOM.
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import vm from 'node:vm';
import {Worker as ThreadWorker} from 'node:worker_threads';
import { bindVbaEditorSettings } from './vba-editor-settings.js';
import { closeVbaProject, getVbaSourceForSave, recordVbaSourceChange, runVbaInWorker, prepareVbaSaveInWorker, compileVbaInWorker, getVbaObjectReferences, releaseAllVbaObjectsInWorker, bindVbaWorkbookPicker, vbaWorkbookExtension } from './vba-runner.js';
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const pkg = path.resolve(process.argv[2] || path.join(root, 'crates/oxidocs-wasm/pkg'));
const wasm = await import('data:text/javascript;base64,' + Buffer.from(await readFile(path.join(pkg, 'oxidocs_wasm.js'))).toString('base64'));
await wasm.default({ module_or_path: await readFile(path.join(pkg, 'oxidocs_wasm_bg.wasm')) });
const page = await readFile(path.join(root, 'web/index.html'), 'utf8');
function lift(name, source = page) {
  if (typeof source !== 'string') source = page;
  const at = source.indexOf(`function ${name}(`);
  assert.ok(at >= 0, name);
  let depth = 0;
  for (let i = source.indexOf('{', at); i < source.length; i++) {
    if (source[i] === '{') depth++;
    if (source[i] === '}' && --depth === 0) {
      return (source.slice(at - 6, at) === 'async ' ? 'async ' : '') + source.slice(at, i + 1);
    }
  }
  assert.fail(`unclosed ${name}`);
}
const functions = ['xlsxEditKey', 'collectXlsxEdits', 'defaultCellStyle', 'cellStringValue', 'cellUtf16Text', 'setWorkbookCellString',
  'syncPendingXlsxEditsToIr', 'workbookCellMap', 'typedXlsxEdit', 'recordVbaWorkbookChanges',
  'captureVbaWorkbookView', 'applyVbaWorkbookView', 'clearClosedVbaWorkbookView', 'formatVbaResult', 'formatVbaExecution', 'updateVbaReferenceButton', 'applyVbaExecution', 'executeVbaFromDialog', 'releaseVbaResultsFromDialog', 'compileVbaFromDialog', 'downloadEdited'];
const macroSource = await readFile(path.join(root, 'tests/fixtures/vba/failure_state.bas'), 'utf8');
const applicationSource = await readFile(path.join(root, 'tests/fixtures/vba/application_state.bas'), 'utf8');
const applicationFixture = JSON.parse(await readFile(path.join(root, 'tests/fixtures/vba/application_state.json'), 'utf8'));
const bindingUrl='data:text/javascript;base64,'+Buffer.from(await readFile(path.join(pkg,'oxidocs_wasm.js'))).toString('base64');
const shim=`import init, {SpreadsheetVbaSession,spreadsheet_vba_source,spreadsheet_vba_project_name} from ${JSON.stringify(bindingUrl)};
import {readFile} from 'node:fs/promises';await init({module_or_path:await readFile(${JSON.stringify(path.join(pkg,'oxidocs_wasm_bg.wasm'))})});
export {SpreadsheetVbaSession,spreadsheet_vba_source,spreadsheet_vba_project_name};export default function ready(){return Promise.resolve();}`;
const workerSource=(await readFile(path.join(root,'web/vba-worker.js'),'utf8')).replace("'./oxidocs_wasm.js'",JSON.stringify('data:text/javascript;base64,'+Buffer.from(shim).toString('base64')));
const bridge=`import {parentPort,workerData} from 'node:worker_threads';const queued=[];let listener;
globalThis.self={postMessage:value=>parentPort.postMessage(value),addEventListener:(type,handler)=>{listener=handler;for(const data of queued)listener({data});}};
parentPort.on('message',data=>listener?listener({data}):queued.push(data));await import('data:text/javascript;base64,'+Buffer.from(workerData).toString('base64'));`;
globalThis.Worker = class BrowserWorker {
    constructor(){this.thread=new ThreadWorker(new URL('data:text/javascript;base64,'+Buffer.from(bridge).toString('base64')),{workerData:workerSource});}
    addEventListener(type,handler){this.thread.on(type,value=>handler(type==='message'?{data:value}:value));}
    postMessage(value){this.thread.postMessage(value);}
    terminate(){void this.thread.terminate();}
};

const bytes = new Uint8Array(await readFile(path.join(root, 'tests/fixtures/vba/named_project_ascii.xlsm')));
const nodes = new Map();
const node = id => {
  if (!nodes.has(id)) nodes.set(id, { value: '', textContent: '', disabled: false, style: {},
    listeners: new Map(),
    addEventListener(type, handler) { if (!this.listeners.has(type)) this.listeners.set(type, []); this.listeners.get(type).push(handler); },
    fire(type) { return Promise.all((this.listeners.get(type) || []).map(handler => handler({ target: this }))); },
    replaceChildren(...children) { this.children = children; },
    ownerDocument: {createElement: () => ({})},
    classList: {
    values: new Set(), add(name) { this.values.add(name); }, remove(name) { this.values.delete(name); },
    toggle(name, force) { if (force) this.values.add(name); else this.values.delete(name); },
  } });
  return nodes.get(id);
};
const written = [], statuses = [], downloads = [];
let pending = [];
const context = vm.createContext({
  ...wasm, Uint8Array, Map, Set, console: { error() {} },
  document: { getElementById: node, querySelectorAll: () => pending,
    createElement: () => ({ click() { downloads.push(this.download); } }) },
  requestAnimationFrame: resolve => resolve(),
  showStatus: (...args) => statuses.push(args),
  renderEditableXlsx() { pending = []; },
  parseVbaArguments: text => text.trim() ? JSON.parse(text) : [],
  formatCellValue: value => typeof value === 'object' ? String(Object.values(value)[0]) : '',
  Blob: class { constructor(parts, options) { this.parts = parts; this.type = options.type; } },
  URL: { createObjectURL: blob => { written.push(blob); return 'blob:test'; }, revokeObjectURL() {} },
  t: text => text, I18N: { en: { downloaded: count => `saved ${count}` } },
  runVbaInWorker, getVbaObjectReferences, releaseAllVbaObjectsInWorker,
  prepareVbaSaveInWorker,
  compileVbaInWorker,
  closeVbaProject, bindVbaEditorSettings, bindVbaWorkbookPicker, vbaWorkbookExtension,
  recordVbaSourceChange, getVbaSourceForSave, setTimeout() { return 0; }, clearTimeout() {},
});
vm.runInContext(`let currentDocIR, currentFileBytes, currentXlsxSheetIndex = 0, wasmReady = true;
let currentFileType = 'xlsx', currentFileName = 'Failure.xlsm', currentLang = 'en';
let vbaWorkbookPicker; let xlsxWorkbookChanged = false, currentXlsxVbaSource = '';
let xlsxVbaSourceBound = false, xlsxVbaSourceBaseline = '', xlsxVbaSourceDirty = false, xlsxVbaSourceRevision = 0;
let currentXlsxSavedCalculation = { mode: 'auto', on_save: true }; const xlsxVbaEdits = new Map();
const MIME_TYPES = { xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' };
${functions.map(name => lift(name)).join('\n')}`, context);
const editorSettings = page.slice(page.indexOf('const vbaEditorSettings ='), page.indexOf('\nfunction refreshVbaProcedures('));
assert.ok(editorSettings.includes('bindVbaEditorSettings'));
vm.runInContext(editorSettings, context);
const sourceHooks = page.slice(page.indexOf('let vbaSourceBeforeInput;'), page.indexOf("document.addEventListener('keydown'", page.indexOf('let vbaSourceBeforeInput;')));
assert.ok(sourceHooks.includes("addEventListener('beforeinput'"));
vm.runInContext(`let vbaProcedureRefreshTimer; function refreshVbaProcedures() {}\n${sourceHooks}`, context);
node('vbaSource').value = macroSource;
async function prepare() {
  closeVbaProject();
  const workbook = wasm.parse_spreadsheet(bytes);
  const seeded = (await runVbaInWorker(workbook, applicationSource, 'ResetState', [], 0, undefined, { applicationState: {} })).workbook;
  context.seeded = seeded; context.packageBytes = bytes;
  vm.runInContext('currentDocIR = seeded; currentFileBytes = packageBytes; currentXlsxSheetIndex = 0; xlsxWorkbookChanged = false; currentXlsxVbaSource = ""; xlsxVbaSourceBound = false; xlsxVbaSourceBaseline = ""; xlsxVbaSourceDirty = false; xlsxVbaSourceRevision = 0; currentXlsxSavedCalculation = { mode: "auto", on_save: true }; xlsxVbaEdits.clear();', context);
  written.length = 0; statuses.length = 0; pending = [];
}
function savedCell(book, col) { return book.sheets[0].rows.find(row => row.index === 1)?.cells.find(cell => cell.col === col); }
for (const [procedure, b1] of [['DivideFailure', 14], ['ManualFailure', 14], ['FormatFailure', 2], ['CompileCalleeFailure', 14]]) {
  await prepare();
  node('vbaProcedure').value = procedure;
  await vm.runInContext('executeVbaFromDialog()', context);
  assert.equal(node('vbaRun').disabled, false);
  assert.ok(node('vbaResult').classList.values.has('error'));
  assert.ok(!node('vbaResult').textContent.includes('Completed.'));
  assert.equal(vm.runInContext('xlsxWorkbookChanged', context), true);
  await vm.runInContext('downloadEdited()', context);
  assert.equal(written.length, 1, procedure);
  const saved = wasm.parse_spreadsheet(written[0].parts[0]);
  assert.equal(savedCell(saved, 1).value.Number, b1, procedure);
  assert.equal(savedCell(saved, 0).style.bold, true, procedure);
  assert.equal(downloads.at(-1), 'edited.xlsm');
}
await prepare();
node('vbaProcedure').value = 'AddSheetFailure';
await vm.runInContext('executeVbaFromDialog()', context);
assert.equal(vm.runInContext('currentXlsxSheetIndex', context), 1, 'an added sheet stays active after failure');
await vm.runInContext('downloadEdited()', context);
assert.equal(written.length, 1);
const addedBook = wasm.parse_spreadsheet(written[0].parts[0]);
assert.equal(addedBook.sheets[1].name, 'Created');
assert.equal(addedBook.sheets[1].rows.find(row => row.index === 1).cells.find(cell => cell.col === 3).value.String, 'kept');

await prepare();
context.seeded.sheets.push({ ...structuredClone(context.seeded.sheets[0]), name: 'Sheet2', rows: [], origin_id: undefined });
node('vbaProcedure').value = 'ActivateFailure';
await vm.runInContext('executeVbaFromDialog()', context);
assert.equal(vm.runInContext('currentXlsxSheetIndex', context), 1, 'activation before failure reaches the editor');
assert.equal(vm.runInContext('xlsxWorkbookChanged', context), false, 'activation alone leaves workbook content clean');

await prepare();
node('vbaProcedure').value = 'CompileEntryFailure';
await vm.runInContext('executeVbaFromDialog()', context);
await vm.runInContext('downloadEdited()', context);
assert.equal(written.length, 0, 'compile before entry has no workbook to save');

await prepare();
// A pending cell edit must survive a failing macro that makes no workbook edit.
pending = [{ dataset: { sheet: '0', row: '3', col: '0', original: '' }, textContent: 'pending' }];
node('vbaProcedure').value = 'StatusFailure';
await vm.runInContext('executeVbaFromDialog()', context);
assert.equal(vm.runInContext('xlsxWorkbookChanged', context), false);
assert.ok(node('vbaResult').textContent.includes('status-only'));
await vm.runInContext('downloadEdited()', context);
assert.equal(written.length, 1);
const saved = wasm.parse_spreadsheet(written[0].parts[0]);
assert.equal(saved.sheets[0].rows.find(row => row.index === 3).cells.find(cell => cell.col === 0).value.String, 'pending');
await prepare();
pending = [{ dataset: { sheet: '0', row: '3', col: '0', original: '' }, textContent: '\uD83D' }];
node('vbaProcedure').value = 'StatusFailure';
await vm.runInContext('executeVbaFromDialog()', context);
assert.ok(node('vbaResult').textContent.includes('status-only'));
await vm.runInContext('downloadEdited()', context);
assert.equal(written.length, 1);
const wideSaved = wasm.parse_spreadsheet(written[0].parts[0]);
assert.deepEqual(wideSaved.sheets[0].rows.find(row => row.index === 3).cells.find(cell => cell.col === 0).value, { Utf16: [0xD83D] });
await prepare();
node('vbaSource').value = applicationSource;
node('vbaProcedure').value = 'Configure';
node('vbaArguments').value = '[1]';
await vm.runInContext('executeVbaFromDialog()', context);
const expectedState = applicationFixture.cases.find(row => row.mode === 1);
for (const [procedure, column] of [['ReadState', 'after'], ['WriteMore', 'after_write'], ['Automatic', 'automatic']]) {
  node('vbaProcedure').value = procedure;
  node('vbaArguments').value = '';
  await vm.runInContext('executeVbaFromDialog()', context);
  assert.ok(node('vbaResult').textContent.includes(expectedState[column]), procedure);
}
const saveSource = await readFile(path.join(root, 'tests/fixtures/vba/save_calculation.bas'), 'utf8');
const saveBytes = new Uint8Array(await readFile(path.join(root, 'tests/fixtures/vba/save_calculation.xlsm')));
async function prepareSaveFixture(udf) {
  const reset = await runVbaInWorker(wasm.parse_spreadsheet(saveBytes), saveSource, 'Reset', [udf], 0, 'Save.xlsm', { applicationState: null }, saveBytes);
  const baseline = wasm.edit_xlsx_from_workbook(saveBytes, reset.workbook);
  context.seeded = wasm.parse_spreadsheet(baseline);
  context.packageBytes = baseline;
  vm.runInContext('currentDocIR = seeded; currentFileBytes = packageBytes; currentFileName = "Save.xlsm"; currentXlsxSheetIndex = 0; xlsxWorkbookChanged = false; currentXlsxVbaSource = ""; currentXlsxSavedCalculation = { mode: "auto", on_save: true }; xlsxVbaEdits.clear();', context);
  node('vbaSource').value = saveSource;
  written.length = 0; pending = [];
}
for (const udf of [false, true]) {
  for (const onSave of [true, false]) {
    await prepareSaveFixture(udf);
    node('vbaProcedure').value = 'Configure';
    node('vbaArguments').value = JSON.stringify([-4135, onSave]);
    await vm.runInContext('executeVbaFromDialog()', context);
    await vm.runInContext('downloadEdited()', context);
    assert.equal(written.length, 1);
    const saved = wasm.parse_spreadsheet(written[0].parts[0]);
    assert.equal(savedCell(saved, 1).value.Number, onSave ? 14 : 2, `main Save UDF=${udf} onSave=${onSave}`);
    assert.equal(saved.calculation.mode, 'manual');
    assert.equal(saved.calculation.on_save, onSave);
    assert.equal(vm.runInContext('currentDocIR.sheets[0].rows.find(row => row.index === 1).cells.find(cell => cell.col === 1).value.Number', context), onSave ? 14 : 2, 'Save updates the live workbook too');
  }
}
await prepareSaveFixture(false);
node('vbaProcedure').value = 'ModeOnly'; node('vbaArguments').value = '[-4105]';
await vm.runInContext('executeVbaFromDialog()', context);
assert.equal(vm.runInContext('xlsxWorkbookChanged', context), true, 'same-mode assignment dirties the main editor');
await vm.runInContext('downloadEdited()', context);
assert.equal(written.length, 1);
await prepareSaveFixture(false);
node('vbaProcedure').value = 'OnSaveOnly'; node('vbaArguments').value = '[false]';
await vm.runInContext('executeVbaFromDialog()', context);
assert.equal(vm.runInContext('xlsxWorkbookChanged', context), false, 'on-save metadata alone is not a workbook content edit');
await vm.runInContext('downloadEdited()', context);
assert.equal(written.length, 1, 'metadata-only settings remain saveable');
assert.equal(wasm.parse_spreadsheet(written[0].parts[0]).calculation.on_save, false);
const editSource = await readFile(path.join(root, 'tests/fixtures/vba/source_edit_standard.bas'), 'utf8');
const editControls = JSON.parse(await readFile(path.join(root, 'tests/fixtures/vba/source_edit_controls.json'), 'utf8'));
for (const name of ['same_static', 'same_body', 'insert_dynamic', 'insert_static']) {
  await prepare();
  node('vbaSource').value = editSource;
  node('vbaProcedure').value = 'Prime'; node('vbaArguments').value = '[0]';
  await vm.runInContext('executeVbaFromDialog()', context);
  node('vbaProcedure').value = 'ReadGlobals'; node('vbaArguments').value = '';
  await vm.runInContext('executeVbaFromDialog()', context);
  const testCase = editControls.find(row => row.name === name);
  const input = node('vbaSource');
  input.selectionStart = testCase.old ? editSource.indexOf(testCase.old) : editSource.indexOf(testCase.insert_after) + testCase.insert_after.length + 1;
  input.selectionEnd = testCase.old ? input.selectionStart + testCase.old.length : input.selectionStart;
  input.fire('beforeinput'); input.value = testCase.source; input.fire('input');
  await vm.runInContext('executeVbaFromDialog()', context);
  assert.ok(node('vbaResult').textContent.includes(testCase.after.replace('A1:9', 'A1:5')), name);
}
await prepareSaveFixture(true);
node('vbaProcedure').value = 'Configure'; node('vbaArguments').value = '[-4135,true]';
await vm.runInContext('executeVbaFromDialog()', context);
node('vbaSource').selectionStart = saveSource.indexOf('Twice = amount * 2');
node('vbaSource').selectionEnd = node('vbaSource').selectionStart + 'Twice = amount * 2'.length;
node('vbaSource').fire('beforeinput');
node('vbaSource').value = saveSource.replace('Twice = amount * 2', 'Twice = amount * 3');
node('vbaSource').fire('input');
await vm.runInContext('downloadEdited()', context);
assert.equal(savedCell(wasm.parse_spreadsheet(written[0].parts[0]), 1).value.Number, 21, 'Save uses the pending editor body without rerunning Configure');
for (const [file, succeeds] of [['project_compile_valid.bas', true], ['project_compile_uncalled_undefined.bas', false]]) {
  await prepare();
  node('vbaSource').value = await readFile(path.join(root, 'tests/fixtures/vba', file), 'utf8');
  const before = vm.runInContext('JSON.stringify(currentDocIR)', context);
  await vm.runInContext('compileVbaFromDialog()', context);
  assert.equal(node('vbaCompile').disabled, false);
  assert.equal(node('vbaRun').disabled, false);
  assert.equal(node('vbaResult').classList.values.has('error'), !succeeds);
  if (succeeds) assert.equal(node('vbaResult').textContent, 'Compilation succeeded.');
  assert.equal(vm.runInContext('JSON.stringify(currentDocIR)', context), before);
  assert.equal(vm.runInContext('xlsxWorkbookChanged', context), false);
}
for (const kind of ['sheet', 'book']) {
  await prepare();
  const packageBytes = new Uint8Array(await readFile(path.join(root, 'tests/fixtures/vba', `document_context_${kind}.xlsm`)));
  context.seeded = wasm.parse_spreadsheet(packageBytes); context.packageBytes = packageBytes;
  vm.runInContext('currentDocIR = seeded; currentFileBytes = packageBytes; xlsxWorkbookChanged = false;', context);
  const source = wasm.spreadsheet_vba_source(packageBytes);
  assert.notEqual(source.replace('.Helper(7)', '.Helper()'), source, 'the invalid control must edit the saved source');
  for (const valid of [true, false]) {
    node('vbaSource').value = valid ? source : source.replace('.Helper(7)', '.Helper()');
    const before = vm.runInContext('JSON.stringify(currentDocIR)', context);
    await vm.runInContext('compileVbaFromDialog()', context);
    assert.equal(node('vbaResult').classList.values.has('error'), !valid, `${kind} package compilation`);
    if (valid) assert.equal(node('vbaResult').textContent, 'Compilation succeeded.');
    assert.equal(vm.runInContext('JSON.stringify(currentDocIR)', context), before);
    assert.equal(vm.runInContext('xlsxWorkbookChanged', context), false);
  }
}
// Exercise the main grid with the actual sparse workbook returned by VBA and
// again after save/reopen; blank cells do not have their own cell records.
const gridContainer = makeElement('div');
function makeElement(tag) {
  return {tag, children: [], style: {}, dataset: {}, className: '', textContent: '',
    classList: {add() {}, remove() {}},
    appendChild(child) { this.children.push(child); return child; },
    addEventListener() {}, querySelectorAll() { return []; }, querySelector() { return null; }};
}
const gridContext = vm.createContext({Map, Set, Math,
  document: {getElementById() { return gridContainer; }, createElement: makeElement, addEventListener() {}},
  formatCellValue: () => '', syncPendingXlsxEditsToIr() {},
});
vm.runInContext(`let currentXlsxSheetIndex = 0; ${['colName', 'buildMergeInfo', 'getColWidthPx', 'getRowHeightPx', 'renderEditableXlsx'].map(lift).join('\n')}`, gridContext);
const bandPackage = new Uint8Array(await readFile(path.join(root, 'tests/fixtures/vba/row_height_state.xlsm')));
const bandSource = `Public Sub FormatBands()
Rows(1).Interior.Color = 255: Rows(2).Insert
Columns(1).Interior.Color = 65280: Columns(2).Insert
End Sub`;
const bandRun = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(bandPackage), bandSource, 'FormatBands', [], 0, 'bands.xlsm', {});
assert.equal(bandRun.error, undefined);
const bandBook = JSON.parse(JSON.stringify(bandRun.workbook));
const bandReopened = wasm.parse_spreadsheet(wasm.edit_xlsx_from_workbook(bandPackage, bandBook));
for (const workbook of [bandBook, bandReopened]) {
  gridContainer.children.length = 0;
  gridContext.workbook = workbook;
  vm.runInContext('renderEditableXlsx(workbook)', gridContext);
  const wrapper = gridContainer.children.find(child => child.className === 'spreadsheet-wrapper');
  const cells = [];
  const visit = element => { if (element.className === 'edit-cell') cells.push(element); for (const child of element.children) visit(child); };
  visit(wrapper);
  const at = (row, col) => cells.find(cell => Number(cell.dataset.row) === row && Number(cell.dataset.col) === col);
  assert.equal(at(2, 25).style.background, '#FF0000', 'blank Z2 inherits row colour');
  assert.equal(at(40, 1).style.background, '#00FF00', 'blank B40 inherits column colour');
}
context.editBandBook = bandBook;
vm.runInContext('setWorkbookCellString(editBandBook, 0, 2, 25, "row"); setWorkbookCellString(editBandBook, 0, 40, 1, "column")', context);
const editedBands = wasm.parse_spreadsheet(wasm.edit_xlsx_from_workbook(bandPackage, bandBook));
const checkBands = wasm.run_spreadsheet_vba_with_outcome(editedBands, `Public Function Check() As String
Check = CStr(Range("Z2").Interior.Color) & "|" & CStr(Range("B40").Interior.Color)
End Function`, 'Check', [], 0, 'bands.xlsm', {});
assert.equal(checkBands.error, undefined);
assert.deepEqual(checkBands.result, {type: 'string', value: '255|65280'}, 'typed cells keep their inherited colours after save');

// The main dialog owns returned references until the user releases them.
// Releasing uses the loaded source even when the editor has pending invalid text.
const leaseUiSource = `Attribute VB_Name = "LeaseUiDriver"
Option Explicit
Public Held As Object
Public Function MakeUiLease() As Object
 Range("Z1").Value = 0
 Set MakeUiLease = New LeaseUiClass
End Function
Public Function EchoUiLease(ByVal target As Object) As Object
 Set EchoUiLease = target
End Function
Public Function MakeHeldUiLease() As Object
 Range("Z1").Value = 0
 Set Held = New LeaseUiClass
 Set MakeHeldUiLease = Held
End Function
Public Sub DropHeld()
 Set Held = Nothing
End Sub
VERSION 1.0 CLASS
Attribute VB_Name = "LeaseUiClass"
Attribute VB_Exposed = True
Private Sub Class_Terminate()
 Range("Z1").Value = Range("Z1").Value + 1
 Debug.Print "terminated"
End Sub
`;
await prepare();
node('vbaSource').value=leaseUiSource;node('vbaProcedure').value='MakeUiLease';node('vbaArguments').value='';
await vm.runInContext('executeVbaFromDialog()',context);
assert.equal(getVbaObjectReferences().length,1);
assert.equal(node('vbaRelease').disabled,false);
const firstUiLease=getVbaObjectReferences()[0];
node('vbaProcedure').value='EchoUiLease';
node('vbaArguments').value=JSON.stringify([{type:'object',value:{reference:firstUiLease}}]);
await vm.runInContext('executeVbaFromDialog()',context);
assert.equal(getVbaObjectReferences().length,2,'alias results own independent leases');
assert.match(node('vbaRelease').textContent,/\(2\)/);
const invalidLeaseUiSource=leaseUiSource.replace('End Function','');
assert.equal(recordVbaSourceChange(leaseUiSource,invalidLeaseUiSource),true);
node('vbaSource').value=invalidLeaseUiSource;
const releasedUiSource=vm.runInContext('currentXlsxVbaSource',context);
vm.runInContext('xlsxWorkbookChanged = false',context);
await vm.runInContext('releaseVbaResultsFromDialog()',context);
assert.equal(node('vbaResult').classList.values.has('error'),false);
assert.equal(getVbaObjectReferences().length,0);
assert.equal(node('vbaRelease').disabled,true);
assert.equal(node('vbaSource').value,invalidLeaseUiSource,'release never edits the editor');
assert.equal(getVbaSourceForSave(),invalidLeaseUiSource,'release preserves pending editing operations');
assert.equal(vm.runInContext('currentXlsxVbaSource',context),releasedUiSource,'release never commits pending source');
assert.equal(savedCell(vm.runInContext('currentDocIR',context),25).value.Number,1,'last alias release runs termination once');
assert.match(node('vbaResult').textContent,/terminated/);
assert.equal(vm.runInContext('xlsxWorkbookChanged',context),true);
assert.equal(node('vbaRun').disabled,false);assert.equal(node('vbaCompile').disabled,false);
await prepare();
node('vbaSource').value=leaseUiSource;node('vbaProcedure').value='MakeHeldUiLease';node('vbaArguments').value='';
await vm.runInContext('executeVbaFromDialog()',context);
assert.equal(getVbaObjectReferences().length,1);
await vm.runInContext('releaseVbaResultsFromDialog()',context);
assert.equal(getVbaObjectReferences().length,0);
assert.equal(savedCell(vm.runInContext('currentDocIR',context),25).value.Number,0,'VBA global still holds the object');
node('vbaProcedure').value='DropHeld';
await vm.runInContext('executeVbaFromDialog()',context);
assert.equal(savedCell(vm.runInContext('currentDocIR',context),25).value.Number,1,'VBA releases its own last reference');
closeVbaProject();
assert.equal(getVbaObjectReferences().length,0);
assert.ok(page.includes("document.getElementById('vbaRelease').addEventListener('click', releaseVbaResultsFromDialog)"));


// The real main-page picker synchronizes pending DOM cells before moving its
// invoking owner. Saving B must retain B's original project metadata, not A's.
await prepare();
const selectionCode = "Attribute VB_Name = \"Driver\"\nPrivate count As Long\nPublic Sub Seed()\n count = 7\n ThisWorkbook.Worksheets(1).Range(\"A1\").Value = \"before\"\n Dim peer As Workbook\n Set peer = Workbooks.Open(\"C:\\B.xlsm\")\nEnd Sub\nPublic Function ReadPeer(ByVal peer As String) As String\n count = count + 1\n ReadPeer = CStr(count) & \"|\" & CStr(Workbooks(peer).Worksheets(1).Range(\"A1\").Value)\nEnd Function\nPublic Sub AddTemplate()\n Dim created As Workbook\n Set created = Workbooks.Add(\"C:\\T.xlsm\")\nEnd Sub\n";
const packetA = wasm.save_spreadsheet_vba_source(bytes,selectionCode);
const mixedSeed = new Uint8Array(await readFile(path.join(root,'tests/fixtures/vba/named_project_mixed.xlsm')));
const packetB = wasm.save_spreadsheet_vba_source(mixedSeed,selectionCode+"\n' B package\n");
assert.notEqual(wasm.spreadsheet_vba_project_name(packetA),wasm.spreadsheet_vba_project_name(packetB));
closeVbaProject();
const selectedSeed = await runVbaInWorker(wasm.parse_spreadsheet(packetA),selectionCode,'Driver.Seed',[],0,'A.xlsm',
 {fileContext:{packageBytes:packetA,files:[{path:'C:\\B.xlsm',bytes:packetB},{path:'C:\\T.xlsm',bytes:packetB}]}});
context.selectedSeed=selectedSeed;context.selectionBytes=packetA;
vm.runInContext('currentDocIR=selectedSeed.workbook;currentFileBytes=selectionBytes;currentFileName="A.xlsm";currentXlsxSheetIndex=0;xlsxVbaSourceBound=true;xlsxVbaSourceBaseline="";vbaWorkbookPicker.reset(currentFileName);applyVbaExecution(selectedSeed,currentDocIR);',context);
node('vbaSource').value=selectionCode;
context.selectionCode=selectionCode;
vm.runInContext('currentXlsxVbaSource=selectionCode;xlsxVbaSourceBaseline=selectionCode;',context);
node('vbaProcedure').value='Driver.ReadPeer';node('vbaArguments').value='["A.xlsm"]';
pending=[{dataset:{sheet:'0',row:'1',col:'0',original:'before'},textContent:'edited-A'}];
node('vbaWorkbook').value='1';await node('vbaWorkbook').fire('change');
assert.equal(vm.runInContext('currentFileName',context),'B.xlsm');
assert.equal(node('vbaWorkbook').value,'1');
assert.deepEqual(Array.from(vm.runInContext('new Uint8Array(currentFileBytes)',context)),Array.from(packetB),'selected B backing bytes');
node('vbaProcedure').value='Driver.ReadPeer';node('vbaArguments').value='["A.xlsm"]';
await vm.runInContext('executeVbaFromDialog()',context);
assert.match(node('vbaResult').textContent,/1\|edited-A/);
pending=[{dataset:{sheet:'0',row:'3',col:'0',original:''},textContent:'edited-B'}];
written.length=0;
await vm.runInContext('downloadEdited()',context);
assert.equal(written.length,1);
const savedB=written[0].parts[0];
assert.equal(wasm.spreadsheet_vba_project_name(savedB),wasm.spreadsheet_vba_project_name(packetB),'B project metadata survives save');
assert.equal(wasm.parse_spreadsheet(savedB).sheets[0].rows.find(row=>row.index===3).cells.find(cell=>cell.col===0).value.String,'edited-B');
node('vbaWorkbook').value='0';await node('vbaWorkbook').fire('change');
assert.equal(vm.runInContext('currentFileName',context),'A.xlsm');
assert.equal(savedCell(vm.runInContext('currentDocIR',context),0).value.String,'edited-A');
await vm.runInContext('executeVbaFromDialog()',context);
assert.match(node('vbaResult').textContent,/8\|edited-A/,'returning to A preserves its private counter');
assert.equal(node('vbaWorkbook').disabled,false);
node('vbaProcedure').value='Driver.AddTemplate';node('vbaArguments').value='[]';
await vm.runInContext('executeVbaFromDialog()',context);
assert.ok(!node('vbaResult').classList.values.has('error'),node('vbaResult').textContent);
node('vbaWorkbook').value='2';await node('vbaWorkbook').fire('change');
assert.equal(vm.runInContext('currentFileName',context),'T1');
written.length=0;await vm.runInContext('downloadEdited()',context);
assert.equal(written.length,1);
assert.equal(downloads.at(-1),'edited.xlsm','untitled VBA template download preserves macro extension');
assert.equal(written[0].type,'application/vnd.ms-excel.sheet.macroEnabled.12');
assert.equal(wasm.spreadsheet_vba_project_name(written[0].parts[0]),wasm.spreadsheet_vba_project_name(packetB));
assert.equal(vbaWorkbookExtension('Book1',{sheets:[{kind:'macro_sheet'}]},false),'xlsm');
assert.equal(vbaWorkbookExtension('Book1',{sheets:[{kind:'international_macro_sheet'}]},false),'xlsm');
assert.equal(vbaWorkbookExtension('Book1',{sheets:[{}]},false),'xlsx');
closeVbaProject();

// The production worker and page handlers switch the backing package and
// source together, then clear the editor when the surviving workbook closes.
const closeSource = "Attribute VB_Name = \"Driver\"\nPublic Sub Seed()\n Dim peer As Workbook\n Set peer = Workbooks.Open(\"C:\\CloseUI-B.xlsm\")\n peer.Worksheets(1).Range(\"A1\").Value = \"survivor\"\nEnd Sub\nPublic Sub Finish()\n ThisWorkbook.Close False\nEnd Sub\n\nPublic Sub ClosePeer()\n Workbooks(\"CloseUI-B.xlsm\").Close False\nEnd Sub\n";
const closePacketA = wasm.save_spreadsheet_vba_source(bytes,closeSource);
const closePacketB = wasm.save_spreadsheet_vba_source(mixedSeed,closeSource);
closeVbaProject();
const closeSeed = await runVbaInWorker(wasm.parse_spreadsheet(closePacketA),closeSource,'Driver.Seed',[],0,'CloseUI-A.xlsm',
 {applicationState:{},fileContext:{fullName:'C:\\CloseUI-A.xlsm',files:[{path:'C:\\CloseUI-B.xlsm',bytes:closePacketB}]}},closePacketA);
context.closeSeed=closeSeed;context.closePacketA=closePacketA;
vm.runInContext('currentDocIR=closeSeed.workbook;currentFileBytes=closePacketA;currentFileName="CloseUI-A.xlsm";currentXlsxSheetIndex=0;currentXlsxVbaSource="";vbaWorkbookPicker.reset(currentFileName);applyVbaExecution(closeSeed,currentDocIR);',context);
node('vbaSource').value=closeSource;
node('vbaRun').disabled=node('vbaCompile').disabled=false;
node('vbaProcedure').value='Driver.Finish';node('vbaArguments').value='[]';
await vm.runInContext('executeVbaFromDialog()',context);
assert.equal(vm.runInContext('currentFileName',context),'CloseUI-B.xlsm');
assert.equal(savedCell(vm.runInContext('currentDocIR',context),0).value.String,'survivor');
assert.deepEqual(Array.from(vm.runInContext('new Uint8Array(currentFileBytes)',context)),Array.from(closePacketB));
assert.equal(node('vbaSource').value.replace(/\r\n/g,'\n').trim(),wasm.spreadsheet_vba_source(closePacketB).replace(/\r\n/g,'\n').trim());
assert.deepEqual(node('vbaWorkbook').children.map(child=>child.value),['1']);
assert.equal(node('vbaRun').disabled,false);
assert.equal(node('vbaResult').classList.values.has('error'),false);
node('vbaProcedure').value='Driver.Finish';
await vm.runInContext('executeVbaFromDialog()',context);
assert.equal(vm.runInContext('currentDocIR',context),null);
assert.equal(vm.runInContext('currentFileBytes',context),null);
assert.equal(node('vbaSource').value,'');
assert.deepEqual(node('vbaProcedureList').children,[]);
assert.deepEqual(node('vbaWorkbook').children,[]);
assert.equal(node('vbaWorkbook').disabled,true);
assert.equal(node('vbaRun').disabled,true);
assert.equal(node('vbaCompile').disabled,true);
assert.equal(node('btnDownload').disabled,true);
assert.equal(node('vbaResult').classList.values.has('error'),false);
assert.match(node('pages-container').textContent,/No workbook/);
closeVbaProject();

// Closing only a peer keeps the current source editable while removing
// the peer from the live picker.
const peerSeed = await runVbaInWorker(wasm.parse_spreadsheet(closePacketA),closeSource,'Driver.Seed',[],0,'CloseUI-A.xlsm',
 {applicationState:{},fileContext:{fullName:'C:\\CloseUI-A.xlsm',files:[{path:'C:\\CloseUI-B.xlsm',bytes:closePacketB}]}},closePacketA);
context.peerSeed=peerSeed;
vm.runInContext('currentDocIR=peerSeed.workbook;currentFileBytes=closePacketA;currentFileName="CloseUI-A.xlsm";currentXlsxSheetIndex=0;vbaWorkbookPicker.reset(currentFileName);applyVbaExecution(peerSeed,currentDocIR);',context);
node('vbaSource').value=closeSource;node('vbaProcedure').value='Driver.ClosePeer';
node('vbaRun').disabled=node('vbaCompile').disabled=false;
await vm.runInContext('executeVbaFromDialog()',context);
assert.equal(vm.runInContext('currentFileName',context),'CloseUI-A.xlsm');
assert.notEqual(vm.runInContext('currentDocIR',context),null);
assert.deepEqual(node('vbaWorkbook').children.map(child=>child.value),['0']);
assert.equal(node('vbaRun').disabled,false);
assert.equal(node('vbaResult').classList.values.has('error'),false);
closeVbaProject();

// The canvas editor uses the same production worker, with its own real
// execution/view handlers and empty-grid cleanup.
const demoPage = await readFile(path.join(root,'web/xlsx-demo.html'),'utf8');
const demoNode = id => node('demo-'+id);
let demoPaints = 0;
const demo = vm.createContext({
 ...wasm, Uint8Array, Map, Set, parseVbaArguments: JSON.parse, runVbaInWorker, bindVbaWorkbookPicker,
 document:{getElementById:demoNode}, measureDigit:()=>7,
 weighCalc(){}, markConditional(){}, markFrozen(){}, markSteps(){}, describe(){}, sayFaces(){},
 goTo(){}, listMacroProcedures(){}, showMacroFiles(){}, draw(){demoPaints++;},
 sayMacro(message,bad){demoNode('result').textContent=message;demoNode('result').classList.toggle('error',bad);},
 closeSeed, closePacketA,
});
const demoFunctions=['captureMacroWorkbookView','applyMacroWorkbookView','clearClosedMacroWorkbookView',
 'applyMacroExecution','reportMacro','calculationPendingSave','markSaved'];
vm.runInContext('let book=closeSeed.workbook, source=closePacketA, fileName="CloseUI-A.xlsm", shown=0, geom={};\n'
 +'let macroProjectSource="",macroSourceBound=false,macroSourceBaseline="",macroSourceDirty=false,macroSourceRevision=0,beyondValues=false,digitWidth=7,editor=null;\n'
 +'let savedCalculation={mode:"auto",on_save:true};const at={row:1,col:0},changes=new Map(),pristine=new Map(),conditional=new Map(),undone=[],redone=[];\n'
 +'const macroSource=document.getElementById("source"),macroProc=document.getElementById("proc"),macroArgs=document.getElementById("args");\n'
 +'const macroRun=document.getElementById("run"),macroCompile=document.getElementById("compile"),saveButton=document.getElementById("save"),macroProcList=document.getElementById("list");\n'
 +'const tabsBox=document.getElementById("tabs"),factsBox=document.getElementById("facts"),cellBox=document.getElementById("cell"),missingBox=document.getElementById("missing"),stateBox=document.getElementById("state");\n'
 +'const formulaBox=document.getElementById("formula"),whereBox=document.getElementById("where"),widthBox=document.getElementById("width");\n'
 +'const t=key=>key==="noWorkbook"?"No workbook is open.":key; const macroEditorSettings={options:()=>({}),beginOperation:()=>()=>{}};\n'
 +demoFunctions.map(name=>lift(name,demoPage)).join('\n')
 +'\nconst macroWorkbookPicker=bindVbaWorkbookPicker(document.getElementById("picker"),{capture:captureMacroWorkbookView,apply:applyMacroWorkbookView,busy:()=>false,report:sayMacro});\n',demo);
const handlerStart=demoPage.indexOf('macroRun.onclick = async () => {');
vm.runInContext(demoPage.slice(handlerStart,demoPage.indexOf('macroCompile.onclick =',handlerStart)),demo);
closeVbaProject();
const demoSeed=await runVbaInWorker(wasm.parse_spreadsheet(closePacketA),closeSource,'Driver.Seed',[],0,'CloseUI-A.xlsm',
 {applicationState:{},fileContext:{fullName:'C:\\CloseUI-A.xlsm',files:[{path:'C:\\CloseUI-B.xlsm',bytes:closePacketB}]}},closePacketA);
demo.demoSeed=demoSeed;
vm.runInContext('book=demoSeed.workbook;macroWorkbookPicker.reset(fileName);applyMacroExecution(demoSeed);',demo);
demoNode('source').value=closeSource;demoNode('proc').value='Driver.Finish';demoNode('args').value='[]';
await vm.runInContext('macroRun.onclick()',demo);
assert.equal(vm.runInContext('fileName',demo),'CloseUI-B.xlsm');
assert.equal(savedCell(vm.runInContext('book',demo),0).value.String,'survivor');
assert.deepEqual(Array.from(vm.runInContext('new Uint8Array(source)',demo)),Array.from(closePacketB));
assert.deepEqual(demoNode('picker').children.map(child=>child.value),['1']);
assert.ok(demoPaints>0);
demoNode('proc').value='Driver.Finish';
demoNode('list').children=[{value:'Driver.Finish'}];
await vm.runInContext('macroRun.onclick()',demo);
assert.equal(vm.runInContext('book',demo),null);
assert.equal(vm.runInContext('source',demo),null);
assert.equal(vm.runInContext('geom',demo),null);
assert.deepEqual(demoNode('picker').children,[]);
assert.deepEqual(demoNode('list').children,[]);
assert.equal(demoNode('sheetInk').width,0);
assert.equal(demoNode('save').disabled,true);
assert.equal(demoNode('run').disabled,true);
assert.equal(demoNode('compile').disabled,true);
assert.equal(demoNode('result').classList.values.has('error'),false);
closeVbaProject();


async function seedCollectionCloseUI(route,cancel){
 closeVbaProject();
 const prefix='CollectionUI-'+route+'-'+cancel;
 const a=prefix+'-A.xlsm',b=prefix+'-B.xlsm';
 const source='Attribute VB_Name = "Driver"\nPublic Sub Seed()\nApplication.DisplayAlerts = False\nDim peer As Workbook\nSet peer = Workbooks.Open("C:\\'+b+'")\nThisWorkbook.Worksheets(1).Range("A1").Value = "current"\n'
  +(cancel?'ThisWorkbook.Worksheets(1).Range("A2").Value = "cancel"\n':'')
  +'End Sub\nPublic Sub FinishAll()\nWorkbooks.Close\nEnd Sub\nVERSION 1.0 CLASS\nAttribute VB_Name = "ThisWorkbook"\nPrivate Sub Workbook_BeforeClose(Cancel As Boolean)\nIf CStr(ThisWorkbook.Worksheets(1).Range("A2").Value) = "cancel" Then Cancel = True\nEnd Sub\n';
 const packet=wasm.save_spreadsheet_vba_source(bytes,source);
 const seeded=await runVbaInWorker(wasm.parse_spreadsheet(packet),source,'Driver.Seed',[],0,a,
  {documentModules:wasm.spreadsheet_vba_document_modules(packet),fileContext:{fullName:'C:\\'+a,files:[{path:'C:\\'+b,bytes:packet}]}},packet);
 return {seeded,packet,source,a};
}
for(const cancel of [true,false]){
 try{
  const fixture=await seedCollectionCloseUI('main',cancel);
  context.collectionUI=fixture;
  vm.runInContext('currentDocIR=collectionUI.seeded.workbook;currentFileBytes=collectionUI.packet;currentFileName=collectionUI.a;currentXlsxSheetIndex=0;currentXlsxVbaSource="";vbaWorkbookPicker.reset(currentFileName);applyVbaExecution(collectionUI.seeded,currentDocIR);',context);
  node('vbaSource').value=fixture.source;
  node('vbaProcedure').value='Driver.FinishAll';node('vbaArguments').value='[]';
  node('vbaRun').disabled=node('vbaCompile').disabled=node('btnDownload').disabled=false;
  await vm.runInContext('executeVbaFromDialog()',context);
  assert.equal(node('vbaResult').classList.values.has('error'),false,node('vbaResult').textContent);
  if(cancel){
   assert.equal(vm.runInContext('currentFileName',context),fixture.a);
   assert.equal(savedCell(vm.runInContext('currentDocIR',context),0).value.String,'current');
   assert.deepEqual(node('vbaWorkbook').children.map(child=>child.value),['0','1']);
   assert.equal(node('vbaRun').disabled,false);
   assert.equal(node('btnDownload').disabled,false);
   assert.match(node('vbaSource').value,/FinishAll/);
  }else{
   assert.equal(vm.runInContext('currentDocIR',context),null);
   assert.equal(vm.runInContext('currentFileBytes',context),null);
   assert.equal(node('vbaSource').value,'');
   assert.deepEqual(node('vbaWorkbook').children,[]);
   assert.deepEqual(node('vbaProcedureList').children,[]);
   assert.equal(node('vbaRun').disabled,true);
   assert.equal(node('btnDownload').disabled,true);
  }
 }finally{closeVbaProject();}
}
for(const cancel of [true,false]){
 try{
  const fixture=await seedCollectionCloseUI('canvas',cancel);
  demo.collectionUI=fixture;
  vm.runInContext('book=collectionUI.seeded.workbook;source=collectionUI.packet;fileName=collectionUI.a;shown=0;geom={};macroProjectSource="";macroSourceBound=false;macroSourceDirty=false;macroWorkbookPicker.reset(fileName);applyMacroExecution(collectionUI.seeded);',demo);
  demoNode('source').value=fixture.source;demoNode('proc').value='Driver.FinishAll';demoNode('args').value='[]';
  demoNode('run').disabled=demoNode('compile').disabled=demoNode('save').disabled=false;
  await vm.runInContext('macroRun.onclick()',demo);
  assert.equal(demoNode('result').classList.values.has('error'),false,demoNode('result').textContent);
  if(cancel){
   assert.equal(vm.runInContext('fileName',demo),fixture.a);
   assert.equal(savedCell(vm.runInContext('book',demo),0).value.String,'current');
   assert.deepEqual(demoNode('picker').children.map(child=>child.value),['0','1']);
   assert.equal(demoNode('run').disabled,false);
   assert.equal(demoNode('save').disabled,false);
  }else{
   assert.equal(vm.runInContext('book',demo),null);
   assert.equal(vm.runInContext('source',demo),null);
   assert.deepEqual(demoNode('picker').children,[]);
   assert.equal(demoNode('sheetInk').width,0);
   assert.equal(demoNode('run').disabled,true);
   assert.equal(demoNode('save').disabled,true);
  }
 }finally{closeVbaProject();}
}

console.log(JSON.stringify({ main_editor_failure_save_cases: 8, application_sequence: 1, calculation_save_cases: 6, source_edit_ui_cases: 5, project_compilation_ui_cases: 6, band_grid_roundtrips: 2, typed_band_cells: 2, reference_release_ui_cases: 2, workbook_selection_ui_cases: 1, workbook_close_ui_cases: 3, canvas_workbook_close_ui_cases: 2, workbooks_collection_close_main_ui_cases: 2, workbooks_collection_close_canvas_ui_cases: 2, real_worker: true, selected_package_save: true, untitled_template_macro_save: true, pending_cells_cross_book: true, passed: true }));
