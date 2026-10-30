// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.
import assert from 'node:assert/strict';
import {readFile, writeFile, mkdir} from 'node:fs/promises';
import path from 'node:path';
import {fileURLToPath, pathToFileURL} from 'node:url';
import {Worker as ThreadWorker} from 'node:worker_threads';
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const pkg = path.resolve(process.argv[2] || path.join(root, 'crates/oxidocs-wasm/pkg'));
const bindingUrl = 'data:text/javascript;base64,' + Buffer.from(await readFile(path.join(pkg, 'oxidocs_wasm.js'))).toString('base64');
const wasm = await import(bindingUrl);
await wasm.default({module_or_path: await readFile(path.join(pkg, 'oxidocs_wasm_bg.wasm'))});
const fixture = path.join(root, 'tests/fixtures/vba');
const source = await readFile(path.join(fixture, 'worksheet_cell_width.bas'), 'utf8');
const original = await readFile(path.join(root, 'tests/fixtures/vba/cell_width.xlsm'));
const cases = [];
for (const line of (await readFile(path.join(fixture, 'worksheet_cell_width.tsv'), 'utf8')).split(/\r?\n/)) {
    if (!line || line.startsWith('#')) continue;
    const [version, method, argumentsJson, expected] = line.split('\t');
    cases.push({version: Number(version), method, args: JSON.parse(argumentsJson), expected});
}
assert.equal(cases.length, 49);
// Excel 16.0 build 20527: entire-column shifts carry width and visibility;
// partial cell shifts leave the column properties in place.
const shiftSource = `Public Function ColumnShiftWidth(ByVal mode As Long) As String
Columns(1).ColumnWidth = 12: Columns(2).ColumnWidth = 20: Columns(3).ColumnWidth = 30
Columns(4).ColumnWidth = 8.43: Columns(5).ColumnWidth = 8.43
Range("B1").Value2 = 11: Range("C1").Value2 = 22
If mode = 0 Then Columns(2).Insert
If mode = 1 Then Columns(2).Delete
If mode = 2 Then Range("B1").Insert xlShiftToRight
If mode = 3 Then Range("B1").Delete xlShiftToLeft
If mode = 4 Then Columns(1).Insert
If mode = 5 Then Columns(1).Hidden = True: Columns(2).Insert
If mode = 6 Then Columns(3).Hidden = True: Columns(2).Delete
If mode = 7 Then Range("B:C").EntireColumn.Insert
If mode = 8 Then Columns(2).Insert xlShiftToRight, 1
If mode = 9 Then Columns(2).Hidden = True: Columns(2).Insert xlShiftToRight, 1
ColumnShiftWidth = CStr(Columns(1).ColumnWidth) & "|" & CStr(Columns(2).ColumnWidth) & "|" & CStr(Columns(3).ColumnWidth) & "|" & CStr(Columns(4).ColumnWidth)
End Function
Public Function ReadShiftWidths() As String
ReadShiftWidths = CStr(Columns(1).ColumnWidth) & "|" & CStr(Columns(2).ColumnWidth) & "|" & CStr(Columns(3).ColumnWidth) & "|" & CStr(Columns(4).ColumnWidth)
End Function`;
const shiftCases = ['12|12|20|30', '12|30|8.43|8.43', '12|20|30|8.43', '12|20|30|8.43', '8.43|12|20|30', '0|12|20|30', '12|0|8.43|8.43', '12|12|12|20', '12|20|20|30', '12|20|0|30']
    .map((expected, mode) => ({method: 'ColumnShiftWidth', args: [mode], expected}));
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
    const shifts = new wasm.SpreadsheetVbaSession();
    try {
        assert.deepEqual(shifts.compile(shiftSource, {}), {compiled: true});
        for (const item of shiftCases) {
            verify(wasm.run_spreadsheet_vba_with_outcome(seed(version), shiftSource, item.method, item.args, 0, 'columns.xlsx', options), item, 'shift direct');
            verify(shifts.run(seed(version), shiftSource, item.method, item.args, 0, 'columns.xlsx', options), item, 'shift session');
        }
    } finally {shifts.free();}
}
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
    for (const version of [1]) {
        runner.closeVbaProject();
        assert.deepEqual(await runner.compileVbaInWorker(source, options, seed(version)), {compiled: true});
        let book = seed(version);
        for (const item of cases.filter(c => c.version === version)) {
            book = verify(await runner.runVbaInWorker(book, source, item.method, item.args, 0, 'arrays.xlsx', options), item, 'Worker v' + version);
            book = JSON.parse(JSON.stringify(book));
        }
        runner.closeVbaProject();
        assert.deepEqual(await runner.compileVbaInWorker(shiftSource, options, seed(version)), {compiled: true});
        for (const item of shiftCases) {
            verify(await runner.runVbaInWorker(seed(version), shiftSource, item.method, item.args, 0, 'columns.xlsx', options), item, 'shift Worker');
        }
    }
} finally {runner.closeVbaProject(); globalThis.Worker = previous;}
for (const item of shiftCases.filter(({args}) => [0, 1, 5, 6, 8, 9].includes(args[0]))) {
    const outcome = wasm.run_spreadsheet_vba_with_outcome(seed(1), shiftSource, item.method, item.args, 0, 'columns.xlsx', options);
    verify(outcome, item, 'before shift save');
    const reopened = wasm.parse_spreadsheet(wasm.edit_xlsx_from_workbook(original, outcome.workbook));
    verify(wasm.run_spreadsheet_vba_with_outcome(reopened, shiftSource, 'ReadShiftWidths', [], 0, 'columns.xlsx', options),
        {method: 'ReadShiftWidths', args: [], expected: item.expected}, 'after shift save');
}
const persisted = [];
for (const [name, changes, expected] of [
    ['custom', [[1, 12.5]], '0|Variant():False|A1:1:2|5:13|11:False'],
    ['explicit-standard-size', [[3, 0]], '0|Variant():False|A1:1:2|5:8|11:False'],
    ['new-standard', [[2, 10.5], [0, 0]], '0|Variant():False|A1:1:2|5:11|11:True'],
    ['reset-standard', [[1, 12.5], [0, 0]], '0|Variant():False|A1:1:2|5:8|11:True'],
]) {
    let book = seed(1);
    for (const args of changes) {
        const out = wasm.run_spreadsheet_vba_with_outcome(book, source, 'WidthProbeSet', args, 0, 'arrays.xlsx', options);
        assert.equal(out.error, undefined, name);
        book = out.workbook;
    }
    const bytes = wasm.edit_xlsx_from_workbook(original, book);
    book = wasm.parse_spreadsheet(bytes);
    const args = ['CELL("width",D1)', 'application'];
    verify(wasm.run_spreadsheet_vba_with_outcome(book, source, 'WidthProbeRoot', args, 0, 'arrays.xlsx', options), {method: 'WidthProbeRoot', args, expected}, 'saved ' + name);
    persisted.push({name, expected});
    if (process.argv[3]) {
        await mkdir(process.argv[3], {recursive: true});
        await writeFile(path.join(process.argv[3], name + '.xlsm'), bytes);
    }
}
if (process.argv[3]) await writeFile(path.join(process.argv[3], 'cases.json'), JSON.stringify(persisted, null, 2) + '\n');
console.log(JSON.stringify({passed: true, native_cell_width_cases: 49, native_column_shift_cases: 10, compatibility_versions: 1, direct: true, session: true, worker: true, saved_round_trips: persisted.length, shift_saved_round_trips: 6}));
