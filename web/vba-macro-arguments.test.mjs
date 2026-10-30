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
const cases = JSON.parse(await readFile(path.join(fixtures, 'macro_arguments.json'), 'utf8'));
const book = item => ({ sheets: [{ name: item.host_context.sheets[0].name, rows: [], col_count: 0, col_widths: [],
  default_col_width: 0, default_row_height: 15, merge_cells: [] }] });
const options = item => ({ errorLanguage: 'ja-JP', hostAnsiCodePage: 932, documentModules: item.document_modules });
const check = (result, item) => {
  assert.equal(result.error, undefined, item.name);
  assert.deepEqual(result.result, { type: 'string', value: item.expected }, item.name);
};
const optionalDefaults = [];
for (const stem of ['optional_default_context', 'optional_default_types', 'optional_default_reference_context']) {
  const source = await readFile(path.join(fixtures, stem + '.bas'), 'utf8');
  const rows = (await readFile(path.join(fixtures, stem + '.tsv'), 'utf8')).split(/\r?\n/)
    .filter(line => line && !line.startsWith('#') && !line.startsWith('procedure\t'));
  for (const line of rows) {
    const [procedure, kind, text] = line.split('\t');
    optionalDefaults.push({ source, procedure, expected: kind === 'string' ? text : kind === 'bool' ? text === 'true' : Number(text) });
  }
}
const defaultBook = () => book(cases[0]);
const defaultOptions = { errorLanguage: 'ja-JP', hostAnsiCodePage: 932 };
const checkDefault = (result, item) => {
  assert.equal(result.error, undefined, item.procedure);
  assert.equal(result.result?.value, item.expected, item.procedure);
};
const stringCases = [];
for (const stem of ['string_boundaries', 'string_text_boundaries', 'string_error_boundaries', 'string_literal_boundaries']) {
  const source = await readFile(path.join(fixtures, stem + '.bas'), 'utf8');
  for (const line of (await readFile(path.join(fixtures, stem + '.tsv'), 'utf8')).split(/\r?\n/)) {
    if (!line || line.startsWith('#') || line.startsWith('procedure\t')) continue;
    const [procedure, expected] = line.split('\t');
    if (procedure !== 'CellCases') stringCases.push({ source, procedure, expected });
  }
}
const transportSource = await readFile(path.join(fixtures, 'utf16_transport.bas'), 'utf8');
const transportCases = [
  { procedure: 'ReturnHigh', args: [], expected: '\ud83d' },
  ...['\ud83d', '\ude00', '\ud83d\ude00', '\ud83d\0\ude00', '\ud83d'.repeat(20000)]
    .map(expected => ({ procedure: 'Echo', args: [expected], expected })),
];
for (const item of stringCases) {
  const session = new wasm.SpreadsheetVbaSession();
  try {
    for (let call = 0; call < 3; call++) checkDefault(session.run(defaultBook(), item.source, item.procedure, [], 0, 'strings.xlsm', defaultOptions), item);
  } finally { session.free(); }
}
const checkTransport = (result, item) => {
  assert.equal(result.error, undefined, item.procedure);
  assert.deepEqual(result.result, { type: 'string', value: item.expected }, item.procedure);
};
for (const item of transportCases) checkTransport(wasm.run_spreadsheet_vba_with_outcome(defaultBook(), transportSource, item.procedure, item.args, 0, 'strings.xlsm', defaultOptions), item);
const checkParts = result => {
  assert.equal(result.error, undefined);
  assert.deepEqual(result.result.value.values, ['\ud83d', '\ude00', '\ud83d\ude00'].map(value => ({ type: 'string', value })));
};
checkParts(wasm.run_spreadsheet_vba_with_outcome(defaultBook(), transportSource, 'Parts', [], 0, 'strings.xlsm', defaultOptions));
for (const item of optionalDefaults) {
  const session = new wasm.SpreadsheetVbaSession();
  try {
    let workbook = defaultBook();
    for (let call = 0; call < 3; call++) {
      const result = session.run(workbook, item.source, item.procedure, [], 0, 'defaults.xlsm', defaultOptions);
      checkDefault(result, item); workbook = result.workbook;
    }
    checkDefault(wasm.run_spreadsheet_vba_with_outcome(defaultBook(), item.source, item.procedure, [], 0, 'defaults.xlsm', defaultOptions), item);
  } finally { session.free(); }
}
for (const item of cases) {
  const source = await readFile(path.join(fixtures, item.source), 'utf8');
  const session = new wasm.SpreadsheetVbaSession();
  try {
    let workbook = book(item);
    for (let call = 0; call < 3; call++) {
      const settings = options(item); // Native records use Japanese UI; language is a per-call option.
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
      const settings = options(item); // Native records use Japanese UI; language is a per-call option.
      const result = await runner.runVbaInWorker(workbook, source, item.procedure, [], 0, item.host_context.book_name, settings);
      check(result, item); workbook = result.workbook;
    }
    const saved = await runner.prepareVbaSaveInWorker(workbook, source, 0, item.host_context.book_name,
      { documentModules: item.document_modules });
    assert.equal(saved.error, undefined, item.name);
  }
  for (const item of optionalDefaults) {
    runner.closeVbaProject();
    assert.deepEqual(await runner.compileVbaInWorker(item.source, defaultOptions, defaultBook()), { compiled: true });
    let workbook = defaultBook();
    for (let call = 0; call < 3; call++) {
      const result = await runner.runVbaInWorker(workbook, item.source, item.procedure, [], 0, 'defaults.xlsm', defaultOptions);
      checkDefault(result, item); workbook = result.workbook;
    }
  }
  for (const item of stringCases) {
    runner.closeVbaProject();
    assert.deepEqual(await runner.compileVbaInWorker(item.source, defaultOptions, defaultBook()), { compiled: true });
    for (let call = 0; call < 3; call++) checkDefault(await runner.runVbaInWorker(defaultBook(), item.source, item.procedure, [], 0, 'strings.xlsm', defaultOptions), item);
  }
  runner.closeVbaProject();
  assert.deepEqual(await runner.compileVbaInWorker(transportSource, defaultOptions, defaultBook()), { compiled: true });
  for (const item of transportCases) checkTransport(await runner.runVbaInWorker(defaultBook(), transportSource, item.procedure, item.args, 0, 'strings.xlsm', defaultOptions), item);
  checkParts(await runner.runVbaInWorker(defaultBook(), transportSource, 'Parts', [], 0, 'strings.xlsm', defaultOptions));
} finally { runner.closeVbaProject(); globalThis.Worker = previous; }
console.log(JSON.stringify({ macro_argument_projects: cases.length, macro_argument_cases: cases.reduce((n, item) => n + item.measured_cases, 0), direct_outputs: cases.length * 3,
  worker_outputs: cases.length * 3, free_outputs: cases.length, fresh_outputs: cases.length,
  worker_compile_cases: cases.length, optional_default_cases: optionalDefaults.length, native_string_cases: stringCases.length,
  utf16_transport_cases: transportCases.length + 1, passed: true }));
