// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.
import assert from 'node:assert/strict';
import {readFile} from 'node:fs/promises';
import path from 'node:path';
import {fileURLToPath, pathToFileURL} from 'node:url';
import {createHash} from 'node:crypto';
import {Worker as ThreadWorker} from 'node:worker_threads';
import {explainRow} from './row-model.js';
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const pkg = path.resolve(process.argv[2] || path.join(root, 'crates/oxidocs-wasm/pkg'));
const bindingUrl = 'data:text/javascript;base64,' + Buffer.from(await readFile(path.join(pkg, 'oxidocs_wasm.js'))).toString('base64');
const wasm = await import(bindingUrl);
await wasm.default({module_or_path: await readFile(path.join(pkg, 'oxidocs_wasm_bg.wasm'))});
const fixture = path.join(root, 'tests/fixtures/vba');
const source = await readFile(path.join(fixture, 'worksheet_row_height_loaded.bas'), 'utf8');
const manifest = JSON.parse(await readFile(path.join(fixture, 'row_height_loaded.json'), 'utf8'));
assert.equal(manifest.cases.length, 157);
const inputs = await Promise.all(manifest.cases.map(async item => {
    const bytes = await readFile(path.join(fixture, item.file));
    assert.equal(createHash('sha256').update(bytes).digest('hex'), item.sha256, item.file);
    return {...item, bytes};
}));
const options = {worksheetDbcsLanguage: false, hostAnsiCodePage: 932, errorLanguage: 'ja-JP'};
const seed = item => {
    const book = wasm.parse_spreadsheet(item.bytes);
    book.calculation.compatibility_version = 1;
    return book;
};
const expected = item => {
    const value = item.expected;
    return [value.row_height, value.height, value.hidden ? 'True' : 'False', value.use_standard ? 'True' : 'False'].join('|');
};
const verify = (outcome, item, label) => {
    assert.equal(outcome.error, undefined, label + ': ' + item.file + ': ' + JSON.stringify(outcome.error));
    assert.deepEqual(outcome.result, {type: 'string', value: expected(item)}, label + ': ' + item.file);
    const sheet = outcome.workbook.sheets[0];
    const row = sheet.rows.find(row => row.index === 4);
    if (row?.custom_height && !row.hidden) {
        const geometry = explainRow({sheet, index: 4, row, columns: [], heightOf: () => undefined, advance: () => 0});
        assert.equal(geometry.px, item.expected.height / 0.75, label + ': pinned browser row: ' + item.file);
    }
    return JSON.parse(JSON.stringify(outcome.workbook));
};
for (const item of inputs) verify(wasm.run_spreadsheet_vba_with_outcome(seed(item), source, 'LoadedHeightSnapshot', [], 0, 'loaded.xlsm', options), item, 'direct');
const session = new wasm.SpreadsheetVbaSession();
try {
    assert.deepEqual(session.compile(source, {}), {compiled: true});
    for (const item of inputs) verify(session.run(seed(item), source, 'LoadedHeightSnapshot', [], 0, 'loaded.xlsm', options), item, 'session');
} finally {session.free();}
const shim = `import init,{SpreadsheetVbaSession,spreadsheet_vba_project_name,spreadsheet_vba_source}from ${JSON.stringify(bindingUrl)};
import{readFile}from'node:fs/promises';await init({module_or_path:await readFile(${JSON.stringify(path.join(pkg, 'oxidocs_wasm_bg.wasm'))})});
export{SpreadsheetVbaSession,spreadsheet_vba_project_name,spreadsheet_vba_source};export default function ready(){return Promise.resolve();}`;
const workerSource = (await readFile(path.join(root, 'web/vba-worker.js'), 'utf8')).replace("'./oxidocs_wasm.js'", JSON.stringify('data:text/javascript;base64,' + Buffer.from(shim).toString('base64')));
const bridge = `import{parentPort,workerData}from'node:worker_threads';const queued=[];let listener;globalThis.self={postMessage:value=>parentPort.postMessage(value),addEventListener:(type,handler)=>{listener=handler;for(const data of queued)listener({data});}};parentPort.on('message',data=>listener?listener({data}):queued.push(data));await import('data:text/javascript;base64,'+Buffer.from(workerData).toString('base64'));`;
class BrowserWorker {
    constructor() {this.thread = new ThreadWorker(new URL('data:text/javascript;base64,' + Buffer.from(bridge).toString('base64')), {workerData: workerSource});}
    addEventListener(type, handler) {this.thread.on(type, value => handler(type === 'message' ? {data: value} : value));}
    postMessage(value) {this.thread.postMessage(value);}
    terminate() {return this.thread.terminate();}
}
const previous = globalThis.Worker;
globalThis.Worker = BrowserWorker;
const runner = await import(pathToFileURL(path.join(root, 'web/vba-runner.js')));
try {
    runner.closeVbaProject();
    assert.deepEqual(await runner.compileVbaInWorker(source, options, seed(inputs[0])), {compiled: true});
    for (const item of inputs) verify(await runner.runVbaInWorker(seed(item), source, 'LoadedHeightSnapshot', [], 0, 'loaded.xlsm', options), item, 'Worker');
} finally {runner.closeVbaProject(); globalThis.Worker = previous;}
console.log(JSON.stringify({passed: true, native_loaded_geometry_cases: 157, row_properties: 4, direct: true, session: true, worker: true, browser_pinned_rows: true}));
