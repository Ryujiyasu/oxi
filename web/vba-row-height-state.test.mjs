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
const source = await readFile(path.join(fixture, 'worksheet_row_height_state.bas'), 'utf8');
const original = await readFile(path.join(root, 'tests/fixtures/vba/row_height_state.xlsm'));
const cases = [];
for (const line of (await readFile(path.join(fixture, 'worksheet_row_height_state.tsv'), 'utf8')).split(/\r?\n/)) {
    if (!line || line.startsWith('#')) continue;
    const [version, method, argumentsJson, expected] = line.split('\t');
    cases.push({version: Number(version), method, args: JSON.parse(argumentsJson), expected});
}
assert.equal(cases.length, 30);
// Excel 16.0 build 20527: whole-row operations carry row properties, while
// partial shifts leave them in place even when the affected rows are empty.
const shiftSource = `Public Function RowShiftHeight(ByVal mode As Long) As String
Rows(1).RowHeight = 25: Rows(2).RowHeight = 33: Rows(3).RowHeight = 45
Rows(4).RowHeight = 15: Rows(5).RowHeight = 15
Range("A2").Value2 = 11: Range("A3").Value2 = 22
If mode = 0 Then Rows(2).Insert
If mode = 1 Then Rows(2).Delete
If mode = 2 Then Range("A2").Insert xlShiftDown
If mode = 3 Then Range("A2").Delete xlShiftUp
If mode = 4 Then Rows(1).Insert
If mode = 5 Then Rows(1).Hidden = True: Rows(2).Insert
If mode = 6 Then Range("A2:A3").ClearContents: Range("A2").Insert xlShiftDown
If mode = 7 Then Rows(2).Insert xlShiftDown, 1
If mode = 8 Then Rows(2).Hidden = True: Rows(2).Insert xlShiftDown, 1
RowShiftHeight = CStr(Rows(1).RowHeight) & "|" & CStr(Rows(2).RowHeight) & "|" & CStr(Rows(3).RowHeight) & "|" & CStr(Rows(4).RowHeight)
End Function`;
const shiftCases = ['25|25|33|45', '25|45|15|15', '25|33|45|15', '25|33|45|15', '15|25|33|45', '0|25|33|45', '25|33|45|15', '25|33|33|45', '25|33|0|45']
    .map((expected, mode) => ({method: 'RowShiftHeight', args: [mode], expected}));
const formatSource = `Public Function InsertFormatProbe(ByVal kind As Long, ByVal origin As Long) As String
Range("A1:C3").Clear
Range("A1").Interior.Color = 255: Range("B1").Interior.Color = 65280
Range("A2").Interior.Color = 16711680: Range("B2").Interior.Color = 65535
If kind = 0 Then Rows(2).Insert xlShiftDown, origin
If kind = 1 Then Columns(2).Insert xlShiftToRight, origin
If kind = 2 Then Range("A2").Insert xlShiftDown, origin
If kind = 3 Then Range("B1").Insert xlShiftToRight, origin
InsertFormatProbe = ReadInsertColors()
End Function
Public Function ReadInsertColors() As String
ReadInsertColors = CStr(Range("A1").Interior.Color) & "|" & CStr(Range("B1").Interior.Color) & "|" & CStr(Range("A2").Interior.Color) & "|" & CStr(Range("B2").Interior.Color) & "|" & CStr(Range("A3").Interior.Color) & "|" & CStr(Range("C1").Interior.Color)
End Function
Public Function InsertTemplateProbe(ByVal kind As Long) As String
If kind = 0 Or kind = 2 Then Rows(1).Interior.Color = 255 Else Columns(1).Interior.Color = 255
If kind = 0 Then Rows(2).Insert
If kind = 1 Then Columns(2).Insert
If kind = 2 Then Range("A2:C2").Insert xlShiftDown
If kind = 3 Then Range("B1:B3").Insert xlShiftToRight
InsertTemplateProbe = CStr(Range("A2").Interior.Color) & "|" & CStr(Range("B2").Interior.Color) & "|" & CStr(Range("C2").Interior.Color) & "|" & CStr(Range("Z2").Interior.Color) & "|" & CStr(Range("B100").Interior.Color)
End Function
Public Function InsertFormatDetails() As String
Range("A1:A3").Clear
Range("A1").Value2 = 123: Range("A1").Interior.Color = 255
Range("A1").Font.Bold = True: Range("A1").NumberFormat = "0.00"
Range("A1").AddComment "up"
Range("A1").Validation.Add 1, 1, 1, 1, 9
Range("A2").Value2 = 456: Range("A2").AddComment "down"
Range("A2").Insert xlShiftDown, 0
InsertFormatDetails = ReadInsertDetails()
End Function
Public Function ReadInsertDetails() As String
ReadInsertDetails = CStr(IsEmpty(Range("A2").Value2)) & "|" & CStr(Range("A2").Interior.Color) & "|" & CStr(Range("A2").Font.Bold) & "|" & Range("A2").NumberFormat & "|" & CStr(Range("A2").Comment Is Nothing) & "|" & CStr(Range("A2").Validation.Type) & "|" & CStr(Range("A3").Value2) & "|" & Range("A3").Comment.Text
End Function`;
const formatCases = [
    [0, 0, '255|65280|255|65280|16711680|16777215'], [0, 1, '255|65280|16711680|65535|16711680|16777215'],
    [1, 0, '255|255|16711680|16711680|16777215|65280'], [1, 1, '255|65280|16711680|65535|16777215|65280'],
    [2, 0, '255|65280|255|65535|16711680|16777215'], [2, 1, '255|65280|16711680|65535|16711680|16777215'],
    [3, 0, '255|255|16711680|65535|16777215|65280'], [3, 1, '255|65280|16711680|65535|16777215|65280'],
].map(([kind, origin, expected]) => ({method: 'InsertFormatProbe', args: [kind, origin], expected}));
const templateCases = ['255|255|255|255|16777215', '255|255|16777215|16777215|255', '255|255|255|16777215|16777215', '255|255|16777215|16777215|16777215']
    .map((expected, kind) => ({method: 'InsertTemplateProbe', args: [kind], expected}));
const formatDetails = {method: 'InsertFormatDetails', args: [], expected: 'True|255|True|0.00|True|1|456|down'};
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
            verify(wasm.run_spreadsheet_vba_with_outcome(seed(version), shiftSource, item.method, item.args, 0, 'rows.xlsx', options), item, 'shift direct');
            verify(shifts.run(seed(version), shiftSource, item.method, item.args, 0, 'rows.xlsx', options), item, 'shift session');
        }
    } finally {shifts.free();}
    for (const item of [...formatCases, ...templateCases, formatDetails]) {
        verify(wasm.run_spreadsheet_vba_with_outcome(seed(version), formatSource, item.method, item.args, 0, 'rows.xlsx', options), item, 'format direct');
        const formats = new wasm.SpreadsheetVbaSession();
        try {
            assert.deepEqual(formats.compile(formatSource, {}), {compiled: true});
            verify(formats.run(seed(version), formatSource, item.method, item.args, 0, 'rows.xlsx', options), item, 'format session');
        } finally {formats.free();}
    }
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
            verify(await runner.runVbaInWorker(seed(version), shiftSource, item.method, item.args, 0, 'rows.xlsx', options), item, 'shift Worker');
        }
        for (const item of [...formatCases, ...templateCases, formatDetails]) {
            runner.closeVbaProject();
            assert.deepEqual(await runner.compileVbaInWorker(formatSource, options, seed(version)), {compiled: true});
            verify(await runner.runVbaInWorker(seed(version), formatSource, item.method, item.args, 0, 'rows.xlsx', options), item, 'format Worker');
        }
    }
} finally {runner.closeVbaProject(); globalThis.Worker = previous;}
for (const item of formatCases) {
    const outcome = wasm.run_spreadsheet_vba_with_outcome(seed(1), formatSource, item.method, item.args, 0, 'rows.xlsx', options);
    verify(outcome, item, 'before format save');
    const reopened = wasm.parse_spreadsheet(wasm.edit_xlsx_from_workbook(original, outcome.workbook));
    verify(wasm.run_spreadsheet_vba_with_outcome(reopened, formatSource, 'ReadInsertColors', [], 0, 'rows.xlsx', options),
        {method: 'ReadInsertColors', args: [], expected: item.expected}, 'after format save');
}
console.log(JSON.stringify({passed: true, native_row_height_state_cases: 30, native_row_shift_cases: 9, native_insert_format_cases: 13, format_saved_round_trips: 8, compatibility_versions: 1, direct: true, session: true, worker: true}));
