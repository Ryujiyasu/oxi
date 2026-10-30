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
const source = await readFile(path.join(fixture, 'worksheet_dirty_calculation.bas'), 'utf8');
const original = await readFile(path.join(root, 'crates/oxicells-core/tests/fixtures/hidden_rows_cols.xlsx'));
const cases = [];
for (const line of (await readFile(path.join(fixture, 'worksheet_dirty_calculation.tsv'), 'utf8')).split(/\r?\n/)) {
    if (!line || line.startsWith('#')) continue;
    const [version, method, argumentsJson, expected] = line.split('\t');
    cases.push({version: Number(version), method, args: JSON.parse(argumentsJson), expected});
}
assert.equal(cases.length, 216);
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
const setup = outcome => {
    assert.equal(outcome.error, undefined);
    assert.equal(outcome.result.value, 21);
    return outcome.workbook;
};
// Excel 20527, three fresh processes for each flag, axis and operation.
// The single-line If also checks Else after an argument-free member call.
const shiftSource = `Option Explicit
Private hitsB As Long
Private trace As String, traceStage As String
Public Function VolB() As Long
    Application.Volatile False
    hitsB = hitsB + 1: trace = trace & traceStage & ":B|": VolB = 20
End Function
Public Function VolTrue() As Long
    Application.Volatile True
    hitsB = hitsB + 1: trace = trace & traceStage & ":T|": VolTrue = 10
End Function
Public Function CallTrace() As String
    CallTrace = trace
End Function
Public Function VolShiftProbe(ByVal kind As Long, ByVal axis As Long) As String
    Dim previous As Long, report As String
    previous = Application.Calculation: Application.Calculation = xlCalculationManual
    Range("Z3:AA4").ClearContents: hitsB = 0: trace = "": traceStage = "seed"
    If kind = 0 Then Range("Z3").Formula2 = "=RAND()+VolB()" Else Range("Z3").Formula2 = "=VolTrue()"
    Application.CalculateFull
    report = "seed:" & CStr(hitsB)
    traceStage = "move"
    If axis = 0 Then Rows(1).Insert Else Columns(1).Insert
    report = report & "|moved:" & CStr(hitsB)
    traceStage = "calc": Application.Calculate
    report = report & "|calc:" & CStr(hitsB)
    traceStage = "again": Application.Calculate
    report = report & "|again:" & CStr(hitsB)
    VolShiftProbe = report
    Range("Z3:AA4").ClearContents
    If axis = 0 Then Rows(1).Delete Else Columns(1).Delete
    Application.Calculation = previous
End Function
`;
const shiftPacket = await readFile(path.join(fixture, 'udf_returned_references.xlsm'));
const shiftCases = [];
for (const automatic of [false, true]) for (const deleting of [false, true]) for (const kind of [0, 1]) for (const axis of [0, 1]) {
    let source = deleting ? shiftSource
        .replace('If axis = 0 Then Rows(1).Insert Else Columns(1).Insert', '__DELETE_MOVE__')
        .replace('If axis = 0 Then Rows(1).Delete Else Columns(1).Delete', 'If axis = 0 Then Rows(1).Insert Else Columns(1).Insert')
        .replace('__DELETE_MOVE__', 'If axis = 0 Then Rows(1).Delete Else Columns(1).Delete')
        .replaceAll('Range("Z3:AA4").ClearContents', 'Range("Y2:AA4").ClearContents') : shiftSource;
    if (automatic) source = source.replace('Application.Calculation = xlCalculationManual', 'Application.Calculation = xlCalculationAutomatic');
    shiftCases.push({automatic, deleting, kind, axis, source,
        expected: automatic
            ? (kind ? 'seed:2|moved:3|calc:4|again:5' : 'seed:2|moved:3|calc:3|again:3')
            : (kind ? 'seed:2|moved:2|calc:3|again:4' : 'seed:2|moved:2|calc:3|again:3'),
        trace: automatic
            ? (kind ? 'seed:T|seed:T|move:T|calc:T|again:T|' : 'seed:B|seed:B|move:B|')
            : (kind ? 'seed:T|seed:T|calc:T|again:T|' : 'seed:B|seed:B|calc:B|')});
}
// Confirm that resuming an automatic calculation did not repeat the edit.
// Excel checks an unaffected sentinel at its once-shifted destination.
for (const item of shiftCases.filter(item => item.automatic && !item.deleting)) {
    const setupLine = '    Range("Z3:AA4").ClearContents: hitsB = 0: trace = "": traceStage = "seed"';
    const movedLine = '    report = report & "|moved:" & CStr(hitsB)';
    assert(item.source.includes(setupLine) && item.source.includes(movedLine));
    const source = item.source.replace(setupLine, '    Range("Z9").Value2 = 99\n' + setupLine)
        .replace(movedLine, `    If axis = 0 Then
        report = report & "|moved:" & CStr(hitsB) & "," & CStr(Range("Z10").Value2 = 99)
    Else
        report = report & "|moved:" & CStr(hitsB) & "," & CStr(Range("AA9").Value2 = 99)
    End If`)
        .replace('    If axis = 0 Then Rows(1).Delete Else Columns(1).Delete\n    Application.Calculation = previous',
            '    If axis = 0 Then Rows(1).Delete Else Columns(1).Delete\n    Range("Z9").ClearContents\n    Application.Calculation = previous');
    shiftCases.push({...item, source, expected:item.expected.replace('|moved:3', '|moved:3,True')});
}
for (const item of shiftCases) {
    const book = wasm.parse_spreadsheet(shiftPacket);
    const args = [book, item.source, 'VolShiftProbe', [item.kind, item.axis], 0, 'shift.xlsm', options];
    verify(wasm.run_spreadsheet_vba_with_outcome(...args), {method:'VolShiftProbe', args:args[3], expected:item.expected}, 'direct structural shift');
    const session = new wasm.SpreadsheetVbaSession();
    try {
        assert.deepEqual(session.compile(item.source, {}), {compiled:true});
        const result = session.run(...args);
        const next = JSON.parse(JSON.stringify(verify(result, {method:'VolShiftProbe', args:args[3], expected:item.expected}, 'session structural shift')));
        verify(session.run(next, item.source, 'CallTrace', [], 0, 'shift.xlsm', options), {method:'CallTrace', args:[], expected:item.trace}, 'session structural shift trace');
    } finally { session.free(); }
}
// Excel 20527, three fresh processes for each axis and operation.
const savedShiftSource = `Option Explicit
Public Function VolSavedProbe(ByVal axis As Long) As String
Dim previous As Long
previous = Application.Calculation: Application.Calculation = xlCalculationManual
ThisWorkbook.Saved = True
If axis = 0 Then
Rows(1).Insert
Else
Columns(1).Insert
End If
VolSavedProbe = CStr(ThisWorkbook.Saved)
If axis = 0 Then
Rows(1).Delete
Else
Columns(1).Delete
End If
Application.Calculation = previous
End Function
`;
const savedShiftCases = [];
for (const deleting of [false, true]) for (const axis of [0, 1]) {
    const source = deleting ? savedShiftSource
        .replace('Rows(1).Insert', '__ROW_DEL__').replace('Rows(1).Delete', 'Rows(1).Insert').replace('__ROW_DEL__', 'Rows(1).Delete')
        .replace('Columns(1).Insert', '__COL_DEL__').replace('Columns(1).Delete', 'Columns(1).Insert').replace('__COL_DEL__', 'Columns(1).Delete')
        : savedShiftSource;
    savedShiftCases.push({deleting, axis, source, expected:'False'});
}
for (const item of savedShiftCases) {
    const args = [wasm.parse_spreadsheet(shiftPacket), item.source, 'VolSavedProbe', [item.axis], 0, 'saved-shift.xlsm', options];
    verify(wasm.run_spreadsheet_vba_with_outcome(...args), {method:'VolSavedProbe', args:args[3], expected:item.expected}, 'direct saved structural shift');
    const session = new wasm.SpreadsheetVbaSession();
    try {
        assert.deepEqual(session.compile(item.source, {}), {compiled:true});
        verify(session.run(...args), {method:'VolSavedProbe', args:args[3], expected:item.expected}, 'session saved structural shift');
    } finally { session.free(); }
}
for (const version of [1, 2]) {
    let book = setup(wasm.run_spreadsheet_vba_with_outcome(seed(version), source, 'Driver', [], 0, 'arrays.xlsx', options));
    for (const item of cases.filter(c => c.version === version)) {
        book = verify(wasm.run_spreadsheet_vba_with_outcome(book, source, item.method, item.args, 0, 'arrays.xlsx', options), item, 'direct v' + version);
        if (item.method === 'BeginDirtyState') book = JSON.parse(JSON.stringify(book));
    }
    const session = new wasm.SpreadsheetVbaSession();
    try {
        assert.deepEqual(session.compile(source, {}), {compiled: true});
        book = setup(session.run(seed(version), source, 'Driver', [], 0, 'arrays.xlsx', options));
        for (const item of cases.filter(c => c.version === version)) {
            book = verify(session.run(book, source, item.method, item.args, 0, 'arrays.xlsx', options), item, 'session v' + version);
        }
    } finally {session.free();}
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
    for (const item of shiftCases) {
        runner.closeVbaProject();
        const book = wasm.parse_spreadsheet(shiftPacket);
        assert.deepEqual(await runner.compileVbaInWorker(item.source, options, book), {compiled:true});
        const result = await runner.runVbaInWorker(book, item.source, 'VolShiftProbe', [item.kind, item.axis], 0, 'shift.xlsm', options);
        const next = JSON.parse(JSON.stringify(verify(result, {method:'VolShiftProbe', args:[item.kind, item.axis], expected:item.expected}, 'Worker structural shift')));
        verify(await runner.runVbaInWorker(next, item.source, 'CallTrace', [], 0, 'shift.xlsm', options), {method:'CallTrace', args:[], expected:item.trace}, 'Worker structural shift trace');
    }
    for (const item of savedShiftCases) {
        runner.closeVbaProject();
        const book = wasm.parse_spreadsheet(shiftPacket);
        assert.deepEqual(await runner.compileVbaInWorker(item.source, options, book), {compiled:true});
        verify(await runner.runVbaInWorker(book, item.source, 'VolSavedProbe', [item.axis], 0, 'saved-shift.xlsm', options),
            {method:'VolSavedProbe', args:[item.axis], expected:item.expected}, 'Worker saved structural shift');
    }
    for (const version of [1, 2]) {
        runner.closeVbaProject();
        assert.deepEqual(await runner.compileVbaInWorker(source, options, seed(version)), {compiled: true});
        let book = setup(await runner.runVbaInWorker(seed(version), source, 'Driver', [], 0, 'arrays.xlsx', options));
        for (const item of cases.filter(c => c.version === version)) {
            book = verify(await runner.runVbaInWorker(book, source, item.method, item.args, 0, 'arrays.xlsx', options), item, 'Worker v' + version);
        }
    }
} finally {runner.closeVbaProject(); globalThis.Worker = previous;}
console.log(JSON.stringify({passed: true, native_dirty_calculation_cases: 216, native_structural_shift_cases: 20, native_saved_shift_cases: 4, compatibility_versions: 2, direct: true, session: true, worker: true}));
