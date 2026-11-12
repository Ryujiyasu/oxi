// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.
import assert from 'node:assert/strict';
import {readFile} from 'node:fs/promises';
import path from 'node:path';
import {fileURLToPath, pathToFileURL} from 'node:url';
import {Worker as ThreadWorker} from 'node:worker_threads';
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const pkg = path.resolve(process.argv[2] || path.join(root, 'crates/oxidocs-wasm/pkg'));
const bindingUrl = 'data:text/javascript;base64,' + Buffer.from(await readFile(path.join(pkg, 'oxidocs_wasm.js'))).toString('base64');
const wasm = await import(bindingUrl);
await wasm.default({module_or_path: await readFile(path.join(pkg, 'oxidocs_wasm_bg.wasm'))});
const fixture = path.join(root, 'tests/fixtures/vba');
const source = await readFile(path.join(fixture, 'worksheet_cell_width_controls.bas'), 'utf8');
const original = await readFile(path.join(root, 'tests/fixtures/vba/cell_width_controls.xlsm'));
const cases = [];
for (const line of (await readFile(path.join(fixture, 'worksheet_cell_width_controls.tsv'), 'utf8')).split(/\r?\n/)) {
    if (!line || line.startsWith('#')) continue;
    const [version, method, argumentsJson, expected] = line.split('\t');
    cases.push({version: Number(version), method, args: JSON.parse(argumentsJson), expected});
}
assert.equal(cases.length, 38);
const options = {worksheetDbcsLanguage: false, hostAnsiCodePage: 932, errorLanguage: 'ja-JP'};
const seed = version => {
    const book = wasm.parse_spreadsheet(original);
    book.calculation.compatibility_version = version;
    return book;
};
const verify = (outcome, item, label) => {
    assert.equal(outcome.error, undefined, label + ': ' + JSON.stringify(outcome.error));
    assert.deepEqual(outcome.result, {type: 'string', value: item.expected}, label + ': ' + item.method + '(' + item.args.join(',') + ')');
    return outcome.workbook;
};
// Both existing JS Map callers and JSON's numeric string keys must retain
// column provenance. All returned workbook metadata must survive JSON copying.
for (const custom of [new Map([[3, true]]), {'3': true}]) {
    const book = seed(1);
    book.sheets[0].col_widths[3] = 9.140625;
    book.sheets[0].col_custom_widths = custom;
    book.defined_names = [...(book.defined_names || []), ['__proto__', '1']];
    book.defined_name_attributes = new Map([['__proto__', new Map([['comment', 'preserve']])]]);
    const item = {method: 'WidthProbeRoot', args: ['CELL("width",D1)', 'application'],
        expected: '0|Variant():False|A1:1:2|5:8|11:False'};
    const returned = verify(wasm.run_spreadsheet_vba_with_outcome(book, source, item.method, item.args, 0, 'arrays.xlsx', options), item, 'workbook metadata input');
    assert.deepEqual(JSON.parse(JSON.stringify(returned)), returned, 'JSON copying preserves the complete workbook');
    assert.equal(Object.hasOwn(returned.defined_name_attributes, '__proto__'), true);
    assert.equal(returned.defined_name_attributes.__proto__.comment, 'preserve', 'workbook names remain ordinary metadata keys');
    const saved = wasm.edit_xlsx_from_workbook(original, returned);
    const reopened = wasm.parse_spreadsheet(saved);
    assert.deepEqual(JSON.parse(JSON.stringify(reopened)), reopened, 'parsed workbooks preserve JSON metadata');
    assert.equal(Object.hasOwn(reopened.defined_name_attributes, '__proto__'), true);
    verify(wasm.run_spreadsheet_vba_with_outcome(reopened, source, item.method, item.args, 0, 'arrays.xlsx', options), item, 'workbook metadata save/reopen');
}
// A partially updated client can hold Maps and JSON objects in one workbook.
// Both orders must retain each sheet's explicit column-width provenance.
for (const forms of [['Map', 'Map'], ['object', 'object'], ['Map', 'object'], ['object', 'Map']]) {
    const book = seed(1);
    book.sheets = [book.sheets[0], structuredClone(book.sheets[0])];
    book.sheets[1].name = 'MixedInput';
    for (let index = 0; index < 2; index++) {
        book.sheets[index].col_widths[3] = 9.140625;
        book.sheets[index].col_custom_widths = forms[index] === 'Map' ? new Map([[3, true]]) : {'3': true};
    }
    const item = {method: 'WidthProbeRoot', args: ['CELL("width",D1)', 'application'],
        expected: '0|Variant():False|A1:1:2|5:8|11:False'};
    const outcome = wasm.run_spreadsheet_vba_with_outcome(book, source, item.method, item.args, 0, 'arrays.xlsx', options);
    const returned = verify(outcome, item, 'mixed workbook metadata ' + forms.join('/'));
    assert.equal(outcome.status_bar, undefined, 'absent execution state keeps its protocol representation');
    assert.deepEqual(JSON.parse(JSON.stringify(returned)), returned);
    assert.deepEqual(returned.sheets.map(sheet => sheet.col_custom_widths['3']), [true, true]);
    const reopened = wasm.parse_spreadsheet(wasm.edit_xlsx_from_workbook(original, returned));
    assert.deepEqual(reopened.sheets.map(sheet => sheet.col_custom_widths['3']), [true, true], 'both sheets retain column intent after save/reopen');
    verify(wasm.run_spreadsheet_vba_with_outcome(reopened, source, item.method, item.args, 0, 'arrays.xlsx', options), item, 'mixed metadata save/reopen');
}
for (const version of [1]) {
    let book = seed(version);
    for (const item of cases.filter(c => c.version === version)) {
        book = verify(wasm.run_spreadsheet_vba_with_outcome(book, source, item.method, item.args, 0, 'arrays.xlsx', options), item, 'direct v' + version);
        book = JSON.parse(JSON.stringify(book));
    }
    const session = new wasm.SpreadsheetVbaSession();
    try {
        assert.deepEqual(session.compile(source, {}), {compiled: true});
        book = seed(version);
        for (const item of cases.filter(c => c.version === version)) {
            book = verify(session.run(book, source, item.method, item.args, 0, 'arrays.xlsx', options), item, 'session v' + version);
            book = JSON.parse(JSON.stringify(book));
        }
    } finally {session.free();}
}
const shim = `import init,{SpreadsheetVbaSession,spreadsheet_vba_project_name,spreadsheet_vba_source}from ${JSON.stringify(bindingUrl)};
import{readFile}from'node:fs/promises';await init({module_or_path:await readFile(${JSON.stringify(path.join(pkg, 'oxidocs_wasm_bg.wasm'))})});
export{SpreadsheetVbaSession,spreadsheet_vba_project_name,spreadsheet_vba_source};export default function ready(){return Promise.resolve();}`;
const workerSource = (await readFile(path.join(root, 'web/vba-worker.js'), 'utf8')).replace("'./vba-editor-settings.js'", JSON.stringify(new URL('./vba-editor-settings.js', import.meta.url).href)).replace("'./oxidocs_wasm.js'", JSON.stringify('data:text/javascript;base64,' + Buffer.from(shim).toString('base64')));
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
    for (const version of [1]) {
        runner.closeVbaProject();
        assert.deepEqual(await runner.compileVbaInWorker(source, options, seed(version)), {compiled: true});
        let book = seed(version);
        for (const item of cases.filter(c => c.version === version)) {
            book = verify(await runner.runVbaInWorker(book, source, item.method, item.args, 0, 'arrays.xlsx', options), item, 'Worker v' + version);
            book = JSON.parse(JSON.stringify(book));
        }
    }
} finally {runner.closeVbaProject(); globalThis.Worker = previous;}
console.log(JSON.stringify({passed: true, native_cell_width_control_cases: 38, compatibility_versions: 1, direct: true, session: true, worker: true}));
