// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

// Run after a suite WASM build: node web/vba-errors.test.mjs <package directory>
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { Worker as ThreadWorker } from 'node:worker_threads';
import { inflateRawSync } from 'node:zlib';

// Read the known small Excel fixtures independently of Oxi's ZIP implementation.
function zipPart(bytes, target) {
  const zip = Buffer.from(bytes);
  const end = zip.lastIndexOf(Buffer.from([0x50, 0x4b, 0x05, 0x06]));
  assert.ok(end >= 0);
  let at = zip.readUInt32LE(end + 16);
  for (let entry = 0; entry < zip.readUInt16LE(end + 10); entry++) {
    assert.equal(zip.readUInt32LE(at), 0x02014b50);
    const nameLength = zip.readUInt16LE(at + 28);
    const name = zip.subarray(at + 46, at + 46 + nameLength).toString('utf8');
    if (name === target) {
      const local = zip.readUInt32LE(at + 42);
      assert.equal(zip.readUInt32LE(local), 0x04034b50);
      const start = local + 30 + zip.readUInt16LE(local + 26) + zip.readUInt16LE(local + 28);
      const data = zip.subarray(start, start + zip.readUInt32LE(at + 20));
      const method = zip.readUInt16LE(at + 10);
      assert.ok(method === 0 || method === 8);
      return method === 8 ? inflateRawSync(data) : data;
    }
    at += 46 + nameLength + zip.readUInt16LE(at + 30) + zip.readUInt16LE(at + 32);
  }
  assert.fail(`missing fixture part ${target}`);
}

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const pkg = path.resolve(process.argv[2] || path.join(root, 'crates/oxidocs-wasm/pkg'));
const bindings = await readFile(path.join(pkg, 'oxidocs_wasm.js'), 'utf8');
const wasm = await import('data:text/javascript;base64,' + Buffer.from(bindings).toString('base64'));
await wasm.default({ module_or_path: await readFile(path.join(pkg, 'oxidocs_wasm_bg.wasm')) });
assert.equal(typeof wasm.run_spreadsheet_vba, 'function', 'build with --features suite');

const source = await readFile(path.join(root, 'tests/fixtures/vba/err_properties.bas'), 'utf8');
const table = await readFile(path.join(root, 'tests/fixtures/vba/err_properties.tsv'), 'utf8');
const cases = table.split(/\r?\n/).filter(line => line.length > 0).map(line => {
  const tab = line.indexOf('\t');
  assert.ok(tab > 0, 'each case has a procedure and an expected Excel result');
  return [line.slice(0, tab), line.slice(tab + 1)];
});
assert.ok(cases.length > 0);
assert.equal(new Set(cases.map(([name]) => name)).size, cases.length, 'unique procedures');

const failures = [];
for (const [name, expected] of cases) {
  const workbook = { sheets: [{
    name: 'Sheet1', rows: [], col_count: 0, col_widths: [],
    default_col_width: 0, default_row_height: 15, merge_cells: [],
  }] };
  try {
    const execution = wasm.run_spreadsheet_vba(workbook, source, name, [], 0, 'ErrCompatibility.xlsm');
    assert.deepEqual(execution.result, { type: 'string', value: expected }, name);
  } catch (error) {
    failures.push({ procedure: name, error: error.message });
  }
}
console.log(JSON.stringify({ total: cases.length, passed: cases.length - failures.length, failures }, null, 2));
assert.equal(failures.length, 0, 'WASM results must match the Excel comparison fixture');

assert.equal(typeof wasm.run_spreadsheet_vba_with_options, 'function');
assert.equal(typeof wasm.spreadsheet_vba_project_name, 'function');
const helpFile = 'C:\\Oxi\\VBA\\1041\\VbLR6.chm';
const localizedSource = await readFile(path.join(root, 'tests/fixtures/vba/localized_errors_ja.bas'), 'utf8');
const localizedTable = await readFile(path.join(root, 'tests/fixtures/vba/localized_errors_ja.tsv'), 'utf8');
const localizedCases = localizedTable.trimEnd().split(/\r?\n/).map(line => {
  const tab = line.indexOf('\t');
  return [line.slice(0, tab), line.slice(tab + 1).replaceAll('{VBA_HELP}', helpFile)];
});
const emptyWorkbook = () => ({ sheets: [{
  name: 'Sheet1', rows: [], col_count: 0, col_widths: [],
  default_col_width: 0, default_row_height: 15, merge_cells: [],
}] });
for (const [name, expected] of localizedCases) {
  const execution = wasm.run_spreadsheet_vba_with_options(emptyWorkbook(), localizedSource, name, [], 0, undefined, {
    errorLanguage: 'ja-JP', vbaHelpFile: helpFile,
  });
  assert.deepEqual(execution.result, { type: 'string', value: expected }, name);
}
const projectTable = await readFile(path.join(root, 'tests/fixtures/vba/project_sources_ja.tsv'), 'utf8');
const projectCases = projectTable.trimEnd().split(/\r?\n/).map(line => {
  const first = line.indexOf('\t'), second = line.indexOf('\t', first + 1);
  return [line.slice(0, first), line.slice(first + 1, second), line.slice(second + 1).replaceAll('{VBA_HELP}', helpFile)];
});
const projectPackages = new Map();
for (const [label, expected] of [['ascii', 'OxiNamedProject'], ['mixed', 'Another_Project42'], ['unicode', '日本語検証']]) {
  const bytes = await readFile(path.join(root, 'tests/fixtures/vba', `named_project_${label}.xlsm`));
  assert.equal(wasm.spreadsheet_vba_project_name(bytes, 932), expected);
  if (label !== 'unicode') assert.equal(wasm.spreadsheet_vba_project_name(bytes), expected);
  projectPackages.set(expected, bytes);
}
assert.equal(wasm.spreadsheet_vba_project_name(await readFile(path.join(root, 'tests/fixtures/basic_test.xlsx'))), undefined);
for (const [projectName, procedure, expected] of projectCases) {
  const workbook = wasm.parse_spreadsheet(projectPackages.get(projectName));
  const execution = wasm.run_spreadsheet_vba_with_options(workbook, localizedSource, procedure, [], 0, undefined, { errorLanguage: 'ja-JP', vbaHelpFile: helpFile, projectName });
  assert.deepEqual(execution.result, { type: 'string', value: expected }, `${projectName}.${procedure}`);
}
assert.throws(() => wasm.run_spreadsheet_vba_with_options(emptyWorkbook(), localizedSource, 'RaiseBuiltin', [], 0, undefined, { projectName: 42 }), /invalid VBA options/);
const unhandledCustomSource = 'Sub Main()\nErr.Raise 513, "VBA", "before" & Chr(1) & "after"\nEnd Sub';
assert.throws(() => wasm.run_spreadsheet_vba_with_options(emptyWorkbook(), unhandledCustomSource, 'Main', [], 0, undefined, { errorLanguage: 'ja-JP' }), error => error.message.includes('before\u0001after'));

const specialSource = await readFile(path.join(root, 'tests/fixtures/vba/special_errors.bas'), 'utf8');
const specialCases = (await readFile(path.join(root, 'tests/fixtures/vba/special_errors_ja.tsv'), 'utf8')).trimEnd().split(/\r?\n/).map(line => line.split('\t'));
for (const [procedure, number, mode, outcome, , expected] of specialCases) {
  const args = [Number(number)];
  if (mode !== 'omitted') args.push(mode === 'empty' ? '' : mode);
  const execute = () => wasm.run_spreadsheet_vba_with_options(emptyWorkbook(), specialSource, procedure, args, 0, undefined, { errorLanguage: 'ja-JP', vbaHelpFile: helpFile });
  if (outcome === 'terminate') {
    assert.throws(execute, error => /line \d+:/.test(error.message), `${procedure}(${number})`);
  } else {
    assert.deepEqual(execute().result, { type: 'string', value: expected.replaceAll('{VBA_HELP}', helpFile) }, `${procedure}(${number}, ${mode})`);
  }
}
const sweepCases = (await readFile(path.join(root, 'tests/fixtures/vba/special_error_sweeps_ja.tsv'), 'utf8')).trimEnd().split(/\r?\n/).map(line => line.split('\t'));
const dispatchSource = await readFile(path.join(root, 'tests/fixtures/vba/err_dispatch.bas'), 'utf8');
const dispatchCases = (await readFile(path.join(root, 'tests/fixtures/vba/err_dispatch_ja.tsv'), 'utf8')).trimEnd().split(/\r?\n/).map(line => line.split('\t'));
for (const [member, kind, argc, expected] of dispatchCases) {
  const execution = wasm.run_spreadsheet_vba_with_options(emptyWorkbook(), dispatchSource, 'DispatchMatrix', [member, Number(kind), Number(argc)], 0, undefined, { errorLanguage: 'ja-JP', vbaHelpFile: helpFile });
  assert.deepEqual(execution.result, { type: 'string', value: expected.replaceAll('{VBA_HELP}', helpFile) }, `${member} ${kind} ${argc}`);
}
const collectionSource = await readFile(path.join(root, 'tests/fixtures/vba/collection_dispatch.bas'), 'utf8');
const collectionCases = (await readFile(path.join(root, 'tests/fixtures/vba/collection_dispatch_ja.tsv'), 'utf8')).trimEnd().split(/\r?\n/).map(line => line.split('\t'));
for (const [member, kind, argc, expected] of collectionCases) {
  const execution = wasm.run_spreadsheet_vba_with_options(emptyWorkbook(), collectionSource, 'DispatchMatrix', [member, Number(kind), Number(argc)], 0, undefined, { errorLanguage: 'ja-JP', vbaHelpFile: helpFile });
  assert.deepEqual(execution.result, { type: 'string', value: expected.replaceAll('{VBA_HELP}', helpFile) }, `Collection ${member} ${kind} ${argc}`);
}
const compileFixtures = path.join(root, 'tests/fixtures/vba/compile');
const compileCases = (await readFile(path.join(compileFixtures, 'cases.tsv'), 'utf8')).split(/\r?\n/).filter(Boolean).map(line => line.split('\t'));
const compileSources = new Map();
for (const [name, main, other, outcome, expected] of compileCases) {
  let source = 'Attribute VB_Name = "EntryProbe"\n' + await readFile(path.join(compileFixtures, main), 'utf8');
  if (other) source += '\nAttribute VB_Name = "OtherProbe"\n' + await readFile(path.join(compileFixtures, other), 'utf8');
  compileSources.set(name, source);
  const execute = () => wasm.run_spreadsheet_vba_with_options(emptyWorkbook(), source, 'Probe', [], 0, undefined, { errorLanguage: 'ja-JP' });
  if (outcome === 'compile') assert.throws(execute, error => error.message.includes('コンパイル エラー: ' + expected), name);
  else assert.deepEqual(execute().result, { type: 'string', value: expected }, name);
}
for (const [kind, useRaise] of [['statement', false], ['raise', true]]) {
  const expected = new Map(sweepCases.filter(([name]) => name === kind).map(([, number, snapshot]) => [number, snapshot]));
  const execution = wasm.run_spreadsheet_vba_with_options(emptyWorkbook(), specialSource, 'Sweep', [useRaise], 0, undefined, { errorLanguage: 'ja-JP' });
  assert.equal(execution.result.type, 'array');
  assert.equal(execution.result.value.values.length, 65536);
  for (let number = 0; number <= 65535; number++) {
    const snapshot = (expected.get(String(number)) ?? expected.get('*')).replaceAll('{NUMBER}', String(number));
    assert.deepEqual(execution.result.value.values[number], { type: 'string', value: snapshot }, `${kind} ${number}`);
  }
}

const catalogTable = await readFile(path.join(root, 'tests/fixtures/vba/error_messages_ja.tsv'), 'utf8');
const catalog = new Map(catalogTable.trimEnd().split(/\r?\n/).map(line => {
  const tab = line.indexOf('\t');
  return [Number(line.slice(0, tab)), line.slice(tab + 1)];
}));
assert.equal(catalog.size, 89);
const catalogSource = 'Public Function Describe(ByVal code As Long) As String\nDescribe = Error$(code)\nEnd Function';
for (let code = 0; code <= 65535; code++) {
  const execution = wasm.run_spreadsheet_vba_with_options(emptyWorkbook(), catalogSource, 'Describe', [code], 0, undefined, { errorLanguage: 'ja-JP' });
  assert.deepEqual(execution.result, { type: 'string', value: catalog.get(code) ?? catalog.get(95) }, `Error$(${code})`);
}
for (const options of [undefined, null, {}, { errorLanguage: 'en-US' }]) {
  assert.deepEqual(wasm.run_spreadsheet_vba_with_options(emptyWorkbook(), catalogSource, 'Describe', [6], 0, undefined, options).result, { type: 'string', value: 'Overflow' });
}
assert.throws(() => wasm.run_spreadsheet_vba_with_options(emptyWorkbook(), catalogSource, 'Describe', [6], 0, undefined, { errorLanguage: 'fr-FR' }), /unsupported VBA error language/);
assert.throws(() => wasm.run_spreadsheet_vba_with_options(emptyWorkbook(), catalogSource, 'Describe', [6], 0, undefined, { errorLanguage: 42 }), /invalid VBA options/);
assert.throws(() => wasm.run_spreadsheet_vba_with_options(emptyWorkbook(), catalogSource, 'Describe', [6], 0, undefined, { vbaHelpFiel: 'typo' }), /invalid VBA options/);
for (const options of [[], 42, 'ja-JP', { vbaHelpFile: 42 }]) {
  assert.throws(() => wasm.run_spreadsheet_vba_with_options(emptyWorkbook(), catalogSource, 'Describe', [6], 0, undefined, options), /invalid VBA options/);
}

// Run the actual worker message handler and runner, using Node threads for the
// browser Worker transport. The same rebuilt WASM serves both entry points.
const bindingUrl = 'data:text/javascript;base64,' + Buffer.from(bindings).toString('base64');
const shim = `import init, { SpreadsheetVbaSession, run_spreadsheet_vba_with_options, run_spreadsheet_vba_with_outcome, prepare_spreadsheet_vba_save, spreadsheet_vba_project_name, spreadsheet_vba_source } from ${JSON.stringify(bindingUrl)};
import { readFile } from 'node:fs/promises';
await init({ module_or_path: await readFile(${JSON.stringify(path.join(pkg, 'oxidocs_wasm_bg.wasm'))}) });
export { SpreadsheetVbaSession, run_spreadsheet_vba_with_options, run_spreadsheet_vba_with_outcome, prepare_spreadsheet_vba_save, spreadsheet_vba_project_name, spreadsheet_vba_source };
export default function ready() { return Promise.resolve(); }`;
const shimUrl = 'data:text/javascript;base64,' + Buffer.from(shim).toString('base64');
const workerSource = (await readFile(path.join(root, 'web/vba-worker.js'), 'utf8')).replace("'./oxidocs_wasm.js'", JSON.stringify(shimUrl));
const bridgeSource = `import { parentPort, workerData } from 'node:worker_threads';
const queued = [];
let listener;
globalThis.self = {
  postMessage: value => parentPort.postMessage(value),
  addEventListener: (type, handler) => {
    if (type !== 'message') throw new Error(type);
    listener = handler;
    for (const data of queued) listener({ data });
  },
};
parentPort.on('message', data => listener ? listener({ data }) : queued.push(data));
await import('data:text/javascript;base64,' + Buffer.from(workerData).toString('base64'));`;
class BrowserWorker {
  constructor(url) {
    assert.equal(path.basename(fileURLToPath(url)), 'vba-worker.js');
    this.thread = new ThreadWorker(new URL('data:text/javascript;base64,' + Buffer.from(bridgeSource).toString('base64')), { workerData: workerSource });
  }
  addEventListener(type, handler) {
    this.thread.on(type, value => handler(type === 'message' ? { data: value } : value));
  }
  postMessage(value) { this.thread.postMessage(value); }
  terminate() { void this.thread.terminate(); }
}
const priorWorker = globalThis.Worker;
const priorNavigator = Object.getOwnPropertyDescriptor(globalThis, 'navigator');
try {
  globalThis.Worker = BrowserWorker;
  Object.defineProperty(globalThis, 'navigator', { configurable: true, value: { language: 'ja-JP' } });
  const { runVbaInWorker } = await import(pathToFileURL(path.join(root, 'web/vba-runner.js')).href);
  const execution = await runVbaInWorker(emptyWorkbook(), localizedSource, 'HostWorksheetFunction', [], 0, 'Localized.xlsm', { vbaHelpFile: helpFile });
  assert.deepEqual(execution.result, { type: 'string', value: localizedCases.find(([name]) => name === 'HostWorksheetFunction')[1] });
  const languageError = await runVbaInWorker(emptyWorkbook(), localizedSource, 'RaiseBuiltin', [], 0, undefined, { vbaHelpFile: helpFile });
  assert.deepEqual(languageError.result, { type: 'string', value: localizedCases.find(([name]) => name === 'RaiseBuiltin')[1] });
  const customSource = await runVbaInWorker(emptyWorkbook(), localizedSource, 'SourceVba', [], 0, undefined, { vbaHelpFile: helpFile });
  assert.deepEqual(customSource.result, { type: 'string', value: localizedCases.find(([name]) => name === 'SourceVba')[1] });
  const overridden = await runVbaInWorker(emptyWorkbook(), catalogSource, 'Describe', [6], 0, undefined, { errorLanguage: 'en-US' });
  assert.deepEqual(overridden.result, { type: 'string', value: 'Overflow' });
  for (const [projectName, bytes] of projectPackages) {
    const expected = projectCases.find(([name, procedure]) => name === projectName && procedure === 'RaiseBuiltin')[2];
    const execution = await runVbaInWorker(wasm.parse_spreadsheet(bytes), localizedSource, 'RaiseBuiltin', [], 0, 'Named.xlsm', { vbaHelpFile: helpFile, projectNameCodePage: 932 }, bytes);
    assert.deepEqual(execution.result, { type: 'string', value: expected });
    const patched = wasm.edit_xlsx_from_workbook(bytes, execution.workbook);
    assert.equal(wasm.spreadsheet_vba_project_name(patched, 932), projectName);
    assert.deepEqual(zipPart(patched, 'xl/vbaProject.bin'), zipPart(bytes, 'xl/vbaProject.bin'), 'macro project bytes survive workbook save');
  }
  const explicit = await runVbaInWorker(emptyWorkbook(), localizedSource, 'RaiseBuiltin', [], 0, undefined, { projectName: 'ExplicitProject', vbaHelpFile: helpFile }, projectPackages.get('OxiNamedProject'));
  assert.deepEqual(explicit.result, { type: 'string', value: localizedCases.find(([name]) => name === 'RaiseBuiltin')[1].replace('VBAProject', 'ExplicitProject') });
  await assert.rejects(runVbaInWorker(emptyWorkbook(), specialSource, 'StatementNested', [40040], 0, undefined), /line \d+:/);
  const mapped = await runVbaInWorker(emptyWorkbook(), specialSource, 'RaiseResumeNext', [40040], 0, undefined, { vbaHelpFile: helpFile });
  assert.deepEqual(mapped.result, { type: 'string', value: specialCases.find(([name, number, mode]) => name === 'RaiseResumeNext' && number === '40040' && mode === 'omitted')[5].replaceAll('{VBA_HELP}', helpFile) });
  await assert.rejects(runVbaInWorker(emptyWorkbook(), compileSources.get('dead_branch'), 'Probe', [], 0, undefined), /コンパイル エラー: 型が一致しません。/);
  const dynamicTarget = await runVbaInWorker(emptyWorkbook(), compileSources.get('variant_long'), 'Probe', [], 0, undefined);
  assert.deepEqual(dynamicTarget.result, { type: 'string', value: '424|Empty|' });
} finally {
  (await import(pathToFileURL(path.join(root, 'web/vba-runner.js')))).closeVbaProject();
  globalThis.Worker = priorWorker;
  if (priorNavigator) Object.defineProperty(globalThis, 'navigator', priorNavigator);
  else delete globalThis.navigator;
}
console.log(JSON.stringify({ localized_snapshots: localizedCases.length, japanese_catalog: 65536, worker_cases: 12, project_cases: projectCases.length, special_cases: specialCases.length, dispatch_cases: dispatchCases.length, collection_cases: collectionCases.length, compile_cases: compileCases.length, error_sweep_cases: 131072, passed: true }));
