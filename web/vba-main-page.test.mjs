// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

// Main editor macro/error/save handlers with real suite WASM and a small DOM.
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import vm from 'node:vm';
import { bindVbaEditorSettings } from './vba-editor-settings.js';
import { closeVbaProject, getVbaSourceForSave, recordVbaSourceChange, runVbaInWorker, prepareVbaSaveInWorker, compileVbaInWorker } from './vba-runner.js';
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const pkg = path.resolve(process.argv[2] || path.join(root, 'crates/oxidocs-wasm/pkg'));
const wasm = await import('data:text/javascript;base64,' + Buffer.from(await readFile(path.join(pkg, 'oxidocs_wasm.js'))).toString('base64'));
await wasm.default({ module_or_path: await readFile(path.join(pkg, 'oxidocs_wasm_bg.wasm')) });
const page = await readFile(path.join(root, 'web/index.html'), 'utf8');
function lift(name) {
  const at = page.indexOf(`function ${name}(`);
  assert.ok(at >= 0, name);
  let depth = 0;
  for (let i = page.indexOf('{', at); i < page.length; i++) {
    if (page[i] === '{') depth++;
    if (page[i] === '}' && --depth === 0) {
      return (page.slice(at - 6, at) === 'async ' ? 'async ' : '') + page.slice(at, i + 1);
    }
  }
  assert.fail(`unclosed ${name}`);
}
const functions = ['xlsxEditKey', 'collectXlsxEdits', 'defaultCellStyle', 'cellStringValue', 'cellUtf16Text', 'setWorkbookCellString',
  'syncPendingXlsxEditsToIr', 'workbookCellMap', 'typedXlsxEdit', 'recordVbaWorkbookChanges',
  'formatVbaResult', 'formatVbaExecution', 'applyVbaExecution', 'executeVbaFromDialog', 'compileVbaFromDialog', 'downloadEdited'];
const macroSource = await readFile(path.join(root, 'tests/fixtures/vba/failure_state.bas'), 'utf8');
const applicationSource = await readFile(path.join(root, 'tests/fixtures/vba/application_state.bas'), 'utf8');
const applicationFixture = JSON.parse(await readFile(path.join(root, 'tests/fixtures/vba/application_state.json'), 'utf8'));
globalThis.Worker = class {
  session = new wasm.SpreadsheetVbaSession();
  handlers = new Map();
  addEventListener(type, handler) { this.handlers.set(type, handler); }
  terminate() { this.session.free(); }
  postMessage(data) {
    queueMicrotask(() => {
      try {
        if (data.action === 'compile') {
          const compilation = this.session.compile(data.source, { ...data.options,
            documentModules: [...(data.packageBytes ? this.session.document_modules(data.packageBytes) : []),
              ...(data.options?.documentModules || [])] });
          this.handlers.get('message')({ data: { ok: compilation.compiled, compilation, error: compilation.error?.message } });
          return;
        }
        const execution = data.action === 'save'
          ? this.session.prepare_save(data.workbook, data.source || (data.packageBytes ? wasm.spreadsheet_vba_source(data.packageBytes) || '' : ''), data.activeSheet, data.fileName, data.options)
          : this.session.run(data.workbook, data.source, data.procedure, data.args, data.activeSheet, data.fileName, data.options);
        this.handlers.get('message')({ data: { ok: !execution.error, execution, error: execution.error?.message } });
      } catch (error) {
        this.handlers.get('message')({ data: { ok: false, error: error.message } });
      }
    });
  }
};
const bytes = new Uint8Array(await readFile(path.join(root, 'tests/fixtures/vba/named_project_ascii.xlsm')));
const nodes = new Map();
const node = id => {
  if (!nodes.has(id)) nodes.set(id, { value: '', textContent: '', disabled: false,
    listeners: new Map(),
    addEventListener(type, handler) { if (!this.listeners.has(type)) this.listeners.set(type, []); this.listeners.get(type).push(handler); },
    fire(type) { for (const handler of this.listeners.get(type) || []) handler({ target: this }); },
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
  runVbaInWorker,
  prepareVbaSaveInWorker,
  compileVbaInWorker,
  closeVbaProject, bindVbaEditorSettings,
  recordVbaSourceChange, getVbaSourceForSave, setTimeout() { return 0; }, clearTimeout() {},
});
vm.runInContext(`let currentDocIR, currentFileBytes, currentXlsxSheetIndex = 0, wasmReady = true;
let currentFileType = 'xlsx', currentFileName = 'Failure.xlsm', currentLang = 'en';
let xlsxWorkbookChanged = false, currentXlsxVbaSource = '';
let xlsxVbaSourceBound = false, xlsxVbaSourceBaseline = '', xlsxVbaSourceDirty = false, xlsxVbaSourceRevision = 0;
let currentXlsxSavedCalculation = { mode: 'auto', on_save: true }; const xlsxVbaEdits = new Map();
const MIME_TYPES = { xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' };
${functions.map(lift).join('\n')}`, context);
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
console.log(JSON.stringify({ main_editor_failure_save_cases: 8, application_sequence: 1, calculation_save_cases: 6, source_edit_ui_cases: 5, project_compilation_ui_cases: 6, band_grid_roundtrips: 2, typed_band_cells: 2, passed: true }));
