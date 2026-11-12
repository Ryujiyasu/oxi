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
const cases = JSON.parse(await readFile(path.join(root, 'tests/fixtures/vba/project_compilation.json'), 'utf8'));
const contexts = JSON.parse(await readFile(path.join(root, 'tests/fixtures/vba/project_compilation_context.json'), 'utf8'));
const signatures = JSON.parse(await readFile(path.join(root, 'tests/fixtures/vba/reference_compilation.json'), 'utf8'));
const bindings = JSON.parse(await readFile(path.join(root, 'tests/fixtures/vba/argument_binding.json'), 'utf8'));
const resolution = JSON.parse(await readFile(path.join(root, 'tests/fixtures/vba/name_resolution.json'), 'utf8'));
const expressions = JSON.parse(await readFile(path.join(root, 'tests/fixtures/vba/expression_compilation.json'), 'utf8'));
const lines = JSON.parse(await readFile(path.join(root, 'tests/fixtures/vba/logical_line_compilation.json'), 'utf8'));
const documents = JSON.parse(await readFile(path.join(root, 'tests/fixtures/vba/document_module_compilation.json'), 'utf8'));
const documentRuntime = JSON.parse(await readFile(path.join(root, 'tests/fixtures/vba/document_runtime.json'), 'utf8'));
const charts = JSON.parse(await readFile(path.join(root, 'tests/fixtures/vba/chart_documents.json'), 'utf8'));
const chartPersistence = JSON.parse(await readFile(path.join(root, 'tests/fixtures/vba/chart_persistence.json'), 'utf8'));
const unicodeModules = JSON.parse(await readFile(path.join(root, 'tests/fixtures/vba/unicode_modules.json'), 'utf8'));
const automationNames = JSON.parse(await readFile(path.join(root, 'tests/fixtures/vba/automation_names.json'), 'utf8'));
const callbyname = JSON.parse(await readFile(path.join(root, 'tests/fixtures/vba/callbyname_binding.json'), 'utf8'));
const referenceCases = JSON.parse(await readFile(path.join(root, 'tests/fixtures/vba/callbyname_references.json'), 'utf8'));
const coercionCases = JSON.parse(await readFile(path.join(root, 'tests/fixtures/vba/callbyname_coercion.json'), 'utf8'));
const macroEntryCases = JSON.parse(await readFile(path.join(root, 'tests/fixtures/vba/macro_entry_points.json'), 'utf8'));
const macroArgumentCases = JSON.parse(await readFile(path.join(root, 'tests/fixtures/vba/macro_arguments.json'), 'utf8'));
const macroNamedCases = JSON.parse(await readFile(path.join(root, 'tests/fixtures/vba/macro_named_arguments.json'), 'utf8'));
const macroCopybackCases = JSON.parse(await readFile(path.join(root, 'tests/fixtures/vba/macro_copyback.json'), 'utf8'));
const allCases = [...cases, ...contexts, ...signatures, ...bindings, ...resolution, ...expressions, ...lines, ...documents,
  ...documentRuntime.map(item => ({ ...item, compilation_only: true })),
  ...charts.map(item => ({ ...item, compilation_only: true })),
  ...chartPersistence.map(item => ({ ...item, compilation_only: true })),
  ...unicodeModules.map(item => ({ ...item, compilation_only: true })),
  ...automationNames.map(item => ({ ...item, compilation_only: true })),
  ...callbyname.map(item => ({ ...item, compilation_only: true })),
  ...referenceCases.map(item => ({ ...item, compilation_only: true })),
  ...coercionCases.map(item => ({ ...item, compilation_only: true })),
  ...macroEntryCases.map(item => ({ ...item, compilation_only: true })),
  ...macroArgumentCases.map(item => ({ ...item, compilation_only: true })),
  ...macroNamedCases.map(item => ({ ...item, compilation_only: true })),
  ...macroCopybackCases.map(item => ({ ...item, compilation_only: true }))];
const fixtures = path.join(root, 'tests/fixtures/vba');
const emptyBook = () => ({ sheets: [{ name: 'Sheet1', rows: [], col_count: 0, col_widths: [], default_col_width: 0, default_row_height: 15, merge_cells: [] }] });
for (const item of allCases) {
  const source = await readFile(path.join(fixtures, item.source), 'utf8');
  const options = { errorLanguage: 'ja-JP', ...item.compilation_options };
  const session = new wasm.SpreadsheetVbaSession();
  try {
    const result = session.compile(source, options);
    assert.deepEqual(result, wasm.compile_spreadsheet_vba(source, options));
    if (item.outcome === 'compile') {
      assert.equal(result.compiled, false, item.name);
      assert.equal(result.error.kind, 'compile', item.name);
      assert.equal(result.error.message, 'コンパイル エラー: ' + item.message, item.name);
    } else {
      assert.deepEqual(result, { compiled: true }, item.name);
      if (item.compilation_only) continue;
      let book = emptyBook();
      for (const step of item.steps) {
        const run = session.run(book, source, item.procedure, [], 0, 'Compile.xlsm', { errorLanguage: 'ja-JP' });
        assert.equal(run.error, undefined, item.name);
        assert.deepEqual(run.result, { type: 'string', value: step.value }, item.name);
        book = run.workbook;
      }
    }
  } finally { session.free(); }
}
const shim = `import init, { SpreadsheetVbaSession, spreadsheet_vba_project_name, spreadsheet_vba_source } from ${JSON.stringify(bindingUrl)};
import { readFile } from 'node:fs/promises';
await init({ module_or_path: await readFile(${JSON.stringify(path.join(pkg, 'oxidocs_wasm_bg.wasm'))}) });
export { SpreadsheetVbaSession, spreadsheet_vba_project_name, spreadsheet_vba_source };
export default function ready() { return Promise.resolve(); }`;
const workerSource = (await readFile(path.join(root, 'web/vba-worker.js'), 'utf8')).replace("'./vba-editor-settings.js'", JSON.stringify(new URL('./vba-editor-settings.js', import.meta.url).href)).replace("'./oxidocs_wasm.js'", JSON.stringify('data:text/javascript;base64,' + Buffer.from(shim).toString('base64')));
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
  for (const item of allCases) {
    runner.closeVbaProject();
    const source = await readFile(path.join(fixtures, item.source), 'utf8');
    const invoke = () => runner.compileVbaInWorker(source, { errorLanguage: 'ja-JP', ...item.compilation_options });
    if (item.outcome === 'compile') {
      await assert.rejects(invoke, error => error.name === 'VbaCompileError'
        && error.compilation.error.kind === 'compile'
        && error.message === 'コンパイル エラー: ' + item.message, item.name);
    } else { assert.deepEqual(await invoke(), { compiled: true }, item.name); }
  }
  // Exercise the production package-to-worker metadata path with real Excel
  // packages and no manually supplied document module options.
  for (const item of documents.filter(item => item.package_fixture)) {
    runner.closeVbaProject();
    const packageBytes = new Uint8Array(await readFile(path.join(fixtures, item.package_fixture)));
    const source = await readFile(path.join(fixtures, item.source), 'utf8');
    const metadata = wasm.spreadsheet_vba_document_modules(packageBytes);
    assert.ok(metadata.some(document => document.name === item.document_name));
    assert.deepEqual(await runner.compileVbaInWorker(source, { errorLanguage: 'ja-JP' }, wasm.parse_spreadsheet(packageBytes), packageBytes), { compiled: true });
    const invalid = source.replace(`${item.document_name}.Helper(7)`, `${item.document_name}.Helper()`);
    await assert.rejects(runner.compileVbaInWorker(invalid, { errorLanguage: 'ja-JP' }, undefined, packageBytes), error => error.name === 'VbaCompileError');
  }
  const stateCases = JSON.parse(await readFile(path.join(fixtures, 'project_compilation_state.json'), 'utf8'));
  for (const item of stateCases) {
    runner.closeVbaProject();
    const source = await readFile(path.join(fixtures, item.source), 'utf8');
    let book = emptyBook();
    const run = async expected => {
      const result = await runner.runVbaInWorker(book, source, item.procedure, [], 0, 'Compile.xlsm', { errorLanguage: 'ja-JP' });
      assert.deepEqual(result.result, { type: 'string', value: expected }, item.name);
      book = result.workbook;
    };
    await run(item.warmup_steps[0].value);
    if (item.message) await assert.rejects(runner.compileVbaInWorker(source, { errorLanguage: 'ja-JP' }), error => error.name === 'VbaCompileError');
    else assert.deepEqual(await runner.compileVbaInWorker(source), { compiled: true });
    for (const step of item.steps) await run(step.value);
    assert.equal(runner.getVbaSourceForSave(), source);
    // Cancelling a stateless compiler worker must retain the loaded macro worker.
    const pending = runner.compileVbaInWorker(source);
    assert.equal(runner.cancelActiveVbaExecution(), true);
    await assert.rejects(pending, error => error.name === 'AbortError');
    await run('4');
  }
  // Exact self-authored Native packages distinguish project readiness from
  // whole-source validation, including uncalled invalid procedures and dead branches.
  const nativeCompileRows = (await readFile(path.join(fixtures, 'project_compile_native_saved.tsv'), 'utf8'))
    .split(/\r?\n/).filter(line => line && !line.startsWith('#') && !line.startsWith('id\t'))
    .map(line => line.split('\t'));
  assert.equal(nativeCompileRows.length, 12);
  let nativeCompileObservations = 0;
  for (const row of nativeCompileRows) {
    const [id, compiled, performed, saved, , procedure, packageHex] = row;
    const bytes = new Uint8Array(Buffer.from(packageHex, 'hex'));
    const source = wasm.spreadsheet_vba_source(bytes, 932);
    const modes = id.startsWith('warm_') ? ['warm'] : id.startsWith('cold_') ? ['cold', 'warm'] : ['cold'];
    for (const mode of modes) {
      const readyAfterWarm = mode === 'warm';
      const expected = {compiled: readyAfterWarm || compiled === '1',
        performed: !readyAfterWarm && performed === '1', saved: readyAfterWarm || saved === '1'};
      const options = {errorLanguage: 'ja-JP', hostAnsiCodePage: 932, worksheetDbcsLanguage: true,
        projectName: 'OxiNamedProject', applicationState: {enableEvents: false, displayAlerts: false, calculation: -4105},
        fileContext: {packageBytes: bytes, sourceCodePage: 932, readOnly: true, editRevision: 0}};
      for (const pathKind of ['direct', 'worker']) {
        let book = wasm.parse_spreadsheet(bytes);
        const direct = pathKind === 'direct' ? new wasm.SpreadsheetVbaSession() : null;
        if (!direct) runner.closeVbaProject();
        try {
          if (direct) assert.equal(typeof direct.compile_project, 'function', 'A newly built implementation is required');
          if (mode === 'warm') {
            const run = direct ? direct.run(book, source, procedure, [], 0, 'CompileState.xlsm', options)
              : await runner.runVbaInWorker(book, source, procedure, [], 0, 'CompileState.xlsm', options, bytes);
            assert.equal(run.error, undefined, id);
            assert.deepEqual(run.result, {type: 'string', value: id.startsWith('warm_') ? 'CONTROL' : 'BS'}, id);
            book = run.workbook;
          }
          let outcome;
          if (direct) outcome = direct.compile_project(book, source, 0, 'CompileState.xlsm', options);
          else {
            try { outcome = await runner.compileVbaProjectInWorker(book, source, 0, 'CompileState.xlsm', options, bytes); }
            catch (error) {
              assert.equal(expected.compiled, false, id);
              assert.equal(error.name, 'VbaCompileError', id);
              assert.equal(error.compilation.compiled, false, id);
              assert.ok(error.execution, 'Compile diagnostics retain the loaded project');
              outcome = error.execution;
            }
          }
          assert.equal(outcome.compilation.compiled, expected.compiled, `${id} ${mode} ${pathKind}`);
          assert.equal(outcome.compilation.performed, expected.performed, `${id} ${mode} ${pathKind}`);
          assert.equal(outcome.workbook_selection.selected.saved, expected.saved, `${id} ${mode} ${pathKind}`);
          assert.equal(outcome.workbook_changed, false, id);
          assert.equal(outcome.error, undefined, id);
          if (!direct) assert.equal(runner.getVbaSourceForSave(), source, 'Project compilation retains editor source');
          nativeCompileObservations++;
        } finally { direct?.free(); if (!direct) runner.closeVbaProject(); }
      }
    }
  }
  assert.equal(nativeCompileObservations, 28);
} finally { runner.closeVbaProject(); globalThis.Worker = priorWorker; }
console.log(JSON.stringify({ excel_project_cases: allCases.length, document_compile_only_cases: documents.length, package_context_cases: 2, state_sequences: 2, direct_and_worker: true, passed: true }));
