// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

// Real suite WASM and browser worker transport across separate invocations.
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
const source = await readFile(path.join(root, 'tests/fixtures/vba/application_state.bas'), 'utf8');
const fixture = JSON.parse(await readFile(path.join(root, 'tests/fixtures/vba/application_state.json'), 'utf8'));
const emptyBook = () => ({ sheets: [{ name: 'Sheet1', rows: [], col_count: 0, col_widths: [], default_col_width: 0, default_row_height: 15, merge_cells: [] }] });
function run(previous, procedure, args = []) {
  return wasm.run_spreadsheet_vba_with_outcome(previous?.workbook ?? emptyBook(), source, procedure, args, 0, undefined, { applicationState: previous?.application_state });
}
for (const testCase of fixture.cases) {
  let previous = run(undefined, 'ResetState');
  previous = run(previous, 'Configure', [testCase.mode]);
  assert.equal(Boolean(previous.error), testCase.failed);
  assert.deepEqual(previous.debug_output, [testCase.inside]);
  for (const [procedure, column] of [['ReadState', 'after'], ['WriteMore', 'after_write'], ['Automatic', 'automatic'], ['ResetText', 'reset_text']]) {
    previous = run(previous, procedure);
    assert.equal(previous.error, undefined);
    assert.deepEqual(previous.result, { type: 'string', value: testCase[column] }, `${testCase.mode} ${procedure}`);
  }
}
for (const transition of fixture.transitions) {
  let previous = run(undefined, 'ResetState');
  previous = run(previous, 'SwitchMode', [transition.initial]);
  for (const [procedure, argument, column] of [['ChangeInput', 7, 'before'], ['SwitchMode', transition.target, 'after'], ['ChangeInput', 9, 'following']]) {
    previous = run(previous, procedure, [argument]);
    assert.deepEqual(previous.result, { type: 'string', value: transition[column] }, `${transition.initial}->${transition.target} ${column}`);
  }
}
for (const applicationState of [[], 42, 'manual', { calcluation: -4135 }, { calculation: 999 }, { calculation: 1.5 },
  { cursor: 4 }, { statusBar: 'x'.repeat(256) }, { enableEvents: 'false' }, { screenUpdating: 0 }, { displayAlerts: null }]) {
  assert.throws(() => wasm.run_spreadsheet_vba_with_outcome(emptyBook(), source, 'ReadState', [], 0, undefined, { applicationState }), /invalid VBA/, JSON.stringify(applicationState));
}
const normalized = wasm.run_spreadsheet_vba_with_outcome(emptyBook(), source, 'ReadState', [], 0, undefined, { applicationState: { cursor: 0, statusBar: '' } });
assert.equal(normalized.application_state.cursor, -4143);
assert.equal(normalized.application_state.statusBar, undefined);
for (const directoryState of [42, [], {currentDrive:'CC',directories:['C:\\']}, {currentDrive:'C',directories:[]},
  {currentDrive:'Z',directories:['C:\\']}, {currentDrive:'C',directories:['relative']},
  {currentDrive:'C',directories:['C:\\','c:\\other']}, {currentDrive:'C',directories:['C:\\'],unknown:1}]) {
  assert.throws(() => wasm.run_spreadsheet_vba_with_outcome(emptyBook(), source, 'ReadState', [], 0, undefined,
    {applicationState:{directoryState}}), /invalid VBA/, JSON.stringify(directoryState));
}

const directoryFixture=JSON.parse(await readFile(path.join(root,'tests/fixtures/vba/directory_lifetime.json'),'utf8'));
const directorySource=await readFile(path.join(root,'tests/fixtures/vba',directoryFixture.source),'utf8');
const directoryPackage=await readFile(path.join(root,'tests/fixtures/vba',directoryFixture.package));
const directoryModules=wasm.spreadsheet_vba_document_modules(directoryPackage);
const directoryBook=()=>wasm.parse_spreadsheet(directoryPackage);
const directoryOptions=(item,state)=>({errorLanguage:'ja-JP',applicationState:state,documentModules:directoryModules,
  fileContext:{packageBytes:directoryPackage,fullName:item.full_name,readOnly:true,folders:directoryFixture.folders,files:[]}});
const checkDirectory=(execution,item)=>{
  if(item.expected_termination){assert.equal(execution.error?.kind,'terminated');assert.deepEqual(execution.result,{type:'empty'});}
  else {assert.equal(execution.error,undefined);assert.deepEqual(execution.result,{type:'string',value:item.expected});}
  assert.deepEqual(execution.application_state.directoryState,item.directory_state_after);
};
{
  const first=directoryFixture.cases[0],state={displayAlerts:false,directoryState:directoryFixture.initial_directory_state};
  const direct=wasm.run_spreadsheet_vba_with_outcome(directoryBook(),directorySource,first.procedure,first.arguments,0,'input-a.xlsm',directoryOptions(first,state));
  checkDirectory(direct,first);
  let session,project,book,previousState=state;
  try {
    for(const item of directoryFixture.cases){
      if(project!==item.project){session?.free();session=new wasm.SpreadsheetVbaSession();project=item.project;book=directoryBook();}
      const execution=session.run(book,directorySource,item.procedure,item.arguments,0,'input-'+item.project+'.xlsm',directoryOptions(item,previousState));
      checkDirectory(execution,item);book=JSON.parse(JSON.stringify(execution.workbook));previousState=JSON.parse(JSON.stringify(execution.application_state));
    }
  } finally {session?.free();}
}

const shim = `import init, { SpreadsheetVbaSession, run_spreadsheet_vba_with_outcome, prepare_spreadsheet_vba_save, spreadsheet_vba_project_name, spreadsheet_vba_source } from ${JSON.stringify(bindingUrl)};
import { readFile } from 'node:fs/promises';
await init({ module_or_path: await readFile(${JSON.stringify(path.join(pkg, 'oxidocs_wasm_bg.wasm'))}) });
export { SpreadsheetVbaSession, run_spreadsheet_vba_with_outcome, prepare_spreadsheet_vba_save, spreadsheet_vba_project_name, spreadsheet_vba_source };
export default function ready() { return Promise.resolve(); }`;
const workerSource = (await readFile(path.join(root, 'web/vba-worker.js'), 'utf8')).replace("'./oxidocs_wasm.js'", JSON.stringify('data:text/javascript;base64,' + Buffer.from(shim).toString('base64')));
const bridge = `import { parentPort, workerData } from 'node:worker_threads';
const queued = []; let listener;
globalThis.self = { postMessage: value => parentPort.postMessage(value),
addEventListener: (type, handler) => { listener = handler; for (const data of queued) listener({ data }); } };
parentPort.on('message', data => listener ? listener({ data }) : queued.push(data));
await import('data:text/javascript;base64,' + Buffer.from(workerData).toString('base64'));`;
class BrowserWorker {
  constructor() { this.thread = new ThreadWorker(new URL('data:text/javascript;base64,' + Buffer.from(bridge).toString('base64')), { workerData: workerSource }); }
  addEventListener(type, handler) { this.thread.on(type, value => handler(type === 'message' ? { data: value } : value)); }
  postMessage(value) { this.thread.postMessage(value); }
  terminate() { void this.thread.terminate(); }
}
const originalWorker = globalThis.Worker;
try {
  globalThis.Worker = BrowserWorker;
  const { runVbaInWorker } = await import(pathToFileURL(path.join(root, 'web/vba-runner.js')));
  let workbook;
  async function invoke(procedure, args = [], explicitSource = source) {
    try {
      const result = await runVbaInWorker(workbook ?? emptyBook(), explicitSource, procedure, args, 0);
      workbook = result.workbook;
      return result;
    } catch (error) {
      if (!error.execution) throw error;
      workbook = error.execution.workbook;
      return error.execution;
    }
  }
  for (const testCase of fixture.cases) {
    workbook = emptyBook();
    await invoke('ResetState');
    const configured = await invoke('Configure', [testCase.mode]);
    assert.equal(Boolean(configured.error), testCase.failed);
    // Mutating a returned object must not mutate the runner's saved settings.
    configured.application_state.calculation = -4105;
    configured.application_state.statusBar = 'caller mutation';
    for (const [procedure, column] of [['ReadState', 'after'], ['WriteMore', 'after_write'], ['Automatic', 'automatic'], ['ResetText', 'reset_text']]) {
      const execution = await invoke(procedure);
      assert.deepEqual(execution.result, { type: 'string', value: testCase[column] }, `worker ${testCase.mode} ${procedure}`);
    }
  }
  for (const transition of fixture.transitions) {
    await invoke('ResetState');
    await invoke('SwitchMode', [transition.initial]);
    for (const [procedure, argument, column] of [['ChangeInput', 7, 'before'], ['SwitchMode', transition.target, 'after'], ['ChangeInput', 9, 'following']]) {
      assert.deepEqual((await invoke(procedure, [argument])).result, { type: 'string', value: transition[column] });
    }
  }
  await invoke('ResetState');
  await invoke('Configure', [0]);
  await assert.rejects(runVbaInWorker(workbook, source, 'ReadState', [], 0, undefined, { applicationState: { calculation: 999 } }), /invalid VBA application state/);
  assert.deepEqual((await invoke('ReadState')).result, { type: 'string', value: fixture.cases[0].after }, 'input failure preserves previous Application state');
  const running = invoke('ReadState');
  await assert.rejects(invoke('ReadState'), /already running/);
  await running;
  // Synchronous postMessage failure must not leave the runner locked.
  globalThis.Worker = class { addEventListener() {} terminate() {} postMessage() { throw new Error('clone failed'); } };
  await assert.rejects(invoke('ReadState'), /clone failed/);
  globalThis.Worker = BrowserWorker;
  assert.deepEqual((await invoke('ReadState')).result, { type: 'string', value: fixture.cases[0].after });
  {
    const runner=await import(pathToFileURL(path.join(root,'web/vba-runner.js')));
    runner.closeVbaProject();
    let project,book;
    for(const[index,item]of directoryFixture.cases.entries()){
      if(project!==item.project){runner.closeVbaProject();project=item.project;book=directoryBook();}
      const state=index===0?{displayAlerts:false,directoryState:directoryFixture.initial_directory_state}:{displayAlerts:false};
      let execution;
      try {execution=await runner.runVbaInWorker(book,directorySource,item.procedure,item.arguments,0,'input-'+item.project+'.xlsm',directoryOptions(item,state),directoryPackage);}
      catch(error){if(!error.execution)throw error;execution=error.execution;}
      checkDirectory(execution,item);book=JSON.parse(JSON.stringify(execution.workbook));
      if(index===1){
        for(const directoryState of [0,false,'',{currentDrive:'C',directories:['C:\\'],unknown:1}]){
          await assert.rejects(runner.runVbaInWorker(book,directorySource,'DirectoryRead',[],0,'input-a.xlsm',
            directoryOptions(item,{displayAlerts:false,directoryState}),directoryPackage), /invalid VBA/);
        }
      }
      // The returned snapshot is caller-owned and must not alter the cached process environment.
      execution.application_state.directoryState.currentDrive='Z';
      execution.application_state.directoryState.directories.length=0;
    }
  }
} finally { (await import(pathToFileURL(path.join(root, 'web/vba-runner.js')))).closeVbaProject(); globalThis.Worker = originalWorker; }
console.log(JSON.stringify({ application_sequences: fixture.cases.length, calculation_transitions: fixture.transitions.length, wasm_and_worker: true, passed: true }));
