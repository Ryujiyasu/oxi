// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.
import assert from 'node:assert/strict';
import {readFile, writeFile, mkdir} from 'node:fs/promises';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {createHash} from 'node:crypto';
import {Worker as ThreadWorker} from 'node:worker_threads';
import {bindVbaFileDownloads} from './vba-file-downloads.js';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const pkg = path.resolve(process.argv[2] || path.join(root, 'crates/oxidocs-wasm/pkg'));
const bindingUrl = 'data:text/javascript;base64,' + Buffer.from(await readFile(path.join(pkg, 'oxidocs_wasm.js'))).toString('base64');
const wasm = await import(bindingUrl);
await wasm.default({module_or_path: await readFile(path.join(pkg, 'oxidocs_wasm_bg.wasm'))});
const fixture = path.join(root, 'tests/fixtures/vba');
const corpus = JSON.parse(await readFile(path.join(fixture, 'save_copy_native_cases.json'), 'utf8'));
const sha = bytes => createHash('sha256').update(bytes).digest('hex');
const variants = JSON.parse(await readFile(path.join(fixture, 'save_copy_filename_native.json'), 'utf8'));
const scalar = JSON.parse(await readFile(path.join(fixture, 'save_copy_scalar_native.json'), 'utf8'));
const directoryOutputs = JSON.parse(await readFile(path.join(fixture, 'directory_outputs_native.json'), 'utf8'));
const legacyCount = corpus.groups.length;
const variantCount = variants.groups.length;
const scalarCount = scalar.groups.length;
const groups = await Promise.all([...corpus.groups, ...variants.groups, ...scalar.groups, ...directoryOutputs.groups].map(async group => {
    const sourceBytes = await readFile(path.join(fixture, group.source));
    const bytes = await readFile(path.join(fixture, group.package));
    assert.equal(sha(sourceBytes), group.source_sha256);
    assert.equal(sha(bytes), group.package_sha256);
    return {...group, sourceText: sourceBytes.toString('utf8'), bytes};
}));
const seed = group => {
    const book = wasm.parse_spreadsheet(group.bytes);
    book.calculation.compatibility_version = 1;
    return book;
};
const fileName = group => group.full_name.split(/[\\/]/).at(-1);
const options = group => ({errorLanguage: 'ja-JP', hostAnsiCodePage: 932,
    ...(group.vba_help_file ? {vbaHelpFile: group.vba_help_file} : {}),
    applicationState: {displayAlerts: false},
    documentModules: wasm.spreadsheet_vba_document_modules(group.bytes),
    fileContext: {packageBytes: group.bytes, fullName: group.full_name, readOnly: group.read_only,
        folders: group.folders, files: group.initial_files}});
const exactCases = group => group.cases.filter(item => !item.result_includes_file_length);
const verify = (outcome, item, group, exact = true) => {
    assert.equal(outcome.error, undefined, `${group.name}: ${JSON.stringify(item.arguments)}: ${JSON.stringify(outcome.error)}`);
    if (exact) assert.deepEqual(outcome.result, {type: 'string', value: item.expected}, `${group.name}: ${JSON.stringify(item.arguments)}`);
    if (item.deleted_files) assert.deepEqual(outcome.deleted_files || [], item.deleted_files);
    if (item.files) {
        const actual = outcome.files || [];
        assert.deepEqual(actual.map(file => file.path).sort(), item.files.map(file => file.path).sort());
        for (const file of actual) {
            const expected = item.files.find(value => value.path === file.path);
            const bytes = new Uint8Array(file.bytes);
            if (expected.bytes) { assert.deepEqual([...bytes], expected.bytes, file.path); continue; }
            assert.equal(wasm.spreadsheet_vba_source(bytes).replaceAll('\r\n', '\n').trimEnd(), group.sourceText.replaceAll('\r\n', '\n').trimEnd());
            const saved = wasm.parse_spreadsheet(bytes);
            for (const [address, row, column] of [['A1', 1, 0], ['A2', 2, 0], ['A3', 3, 0], ['B1', 1, 1]]) {
                const cell = saved.sheets[0].rows.find(value => value.index === row)?.cells.find(value => value.col === column);
                assert.deepEqual(!cell || cell.value === 'Empty' ? null : cell.value, expected.cells[address], `${file.path} ${address}`);
            }
        }
        return;
    }
    if (!item.copied) { assert.equal(outcome.files?.length || 0, 0); return; }
    assert.equal(outcome.files.length, 1);
    const file = outcome.files[0];
    assert.equal(file.path, item.target);
    assert.equal(wasm.spreadsheet_vba_source(new Uint8Array(file.bytes)).replaceAll('\r\n', '\n').trimEnd(), group.sourceText.replaceAll('\r\n', '\n').trimEnd());
    const book = wasm.parse_spreadsheet(new Uint8Array(file.bytes));
    const row = book.sheets[0].rows.find(row => row.index === 4);
    if (item.saved.row4_attributes) {
        const attr = item.saved.row4_attributes;
        assert.equal(row?.height ?? null, attr.ht == null ? null : Math.fround(Number(attr.ht)));
        assert.equal(Boolean(row?.hidden), attr.hidden === '1');
        assert.equal(Boolean(row?.custom_height), attr.customHeight === '1');
    }
    if (item.saved.b2_formula) {
        const cell = book.sheets[0].rows.find(row => row.index === 2).cells.find(cell => cell.col === 1);
        assert.equal(cell.formula, item.saved.b2_formula);
        assert.equal(cell.value.Number, item.saved.b2_value);
    }
};

// Retain complete Native byte-count results for an explicit comparison run.
if (['--all', '--variants', '--scalar', '--directory'].includes(process.argv[3])) {
    const rows = [], exports = [];
    const directory = process.argv[5] && path.resolve(process.argv[5]);
    if (directory) await mkdir(directory, {recursive: true});
    const selected = process.argv[3] === '--directory' ? groups.slice(legacyCount + variantCount + scalarCount)
        : process.argv[3] === '--scalar' ? groups.slice(legacyCount + variantCount, legacyCount + variantCount + scalarCount)
        : process.argv[3] === '--variants' ? groups.slice(legacyCount, legacyCount + variantCount) : groups.slice(0, legacyCount);
    for (const [gi, group] of selected.entries()) {
        const session = new wasm.SpreadsheetVbaSession();
        let book = seed(group);
        try {
            for (const [ci, item] of group.cases.entries()) {
                const outcome = session.run(book, group.sourceText, item.procedure, item.arguments, 0, fileName(group), options(group));
                let validationError;
                try { verify(outcome, item, group, false); } catch (error) { validationError = error.message; }
                const matched = !outcome.error && !validationError && outcome.result?.type === 'string' && outcome.result.value === item.expected;
                rows.push({group: group.name, index: ci, arguments: item.arguments, expected: item.expected, actual: outcome.result, error: outcome.error, validationError, matched});
                if (directory) for (const [fi, file] of (outcome.files || []).entries()) {
                    const saved = path.join(directory, `copy-g${gi}-c${ci}${fi ? '-f'+fi : ''}.xlsm`);
                    await writeFile(saved, new Uint8Array(file.bytes));
                    exports.push({group: group.name, index: ci, path: saved, sha256: sha(new Uint8Array(file.bytes)), virtualPath: file.path, source: group.source, expected: item.files?.find(value => value.path === file.path) || item.saved});
                }
                book = outcome.workbook;
            }
        } finally {session.free();}
    }
    const report = {cases: rows.length, matched: rows.filter(row => row.matched).length, mismatches: rows.filter(row => !row.matched), rows, exports};
    await writeFile(path.resolve(process.argv[4]), JSON.stringify(report, null, 2) + '\n');
    console.log(JSON.stringify({cases: report.cases, matched: report.matched, mismatches: report.mismatches.length, exportedFiles: exports.length}));
    process.exit(report.mismatches.length ? 1 : 0);
}

let checked = 0;
for (const group of groups) {
    const session = new wasm.SpreadsheetVbaSession();
    let book = seed(group);
    let directBook = seed(group);
    try {
        for (const item of exactCases(group)) {
            const direct = wasm.run_spreadsheet_vba_with_outcome(directBook, group.sourceText, item.procedure, item.arguments, 0, fileName(group), options(group));
            verify(direct, item, group);
            directBook = JSON.parse(JSON.stringify(direct.workbook));
            const outcome = session.run(book, group.sourceText, item.procedure, item.arguments, 0, fileName(group), options(group));
            verify(outcome, item, group);
            book = JSON.parse(JSON.stringify(outcome.workbook));
            checked++;
        }
    } finally {session.free();}
}
assert.equal(checked, 275);

const shim = `import init,{SpreadsheetVbaSession,spreadsheet_vba_project_name,spreadsheet_vba_source}from ${JSON.stringify(bindingUrl)};
import{readFile}from'node:fs/promises';await init({module_or_path:await readFile(${JSON.stringify(path.join(pkg, 'oxidocs_wasm_bg.wasm'))})});
export{SpreadsheetVbaSession,spreadsheet_vba_project_name,spreadsheet_vba_source};export default function ready(){return Promise.resolve();}`;
const workerSource = (await readFile(path.join(root, 'web/vba-worker.js'), 'utf8')).replace("'./oxidocs_wasm.js'", JSON.stringify('data:text/javascript;base64,' + Buffer.from(shim).toString('base64')));
const workerCode = `import{parentPort}from'node:worker_threads';const queued=[];let listener;parentPort.on('message',data=>listener?listener({data}):queued.push(data));globalThis.self={addEventListener:(_,cb)=>{listener=cb;for(const data of queued)listener({data});queued.length=0;},postMessage:data=>parentPort.postMessage(data)};await import(${JSON.stringify('data:text/javascript;base64,' + Buffer.from(workerSource).toString('base64'))});`;
const workers = [], previousWorker = globalThis.Worker;
globalThis.Worker = class {
    constructor() {this.thread = new ThreadWorker(new URL('data:text/javascript;base64,' + Buffer.from(workerCode).toString('base64')), {type: 'module'});workers.push(this.thread);}
    addEventListener(type, callback) {this.thread.on(type, value => callback(type === 'message' ? {data: value} : value));}
    postMessage(value) {this.thread.postMessage(value);}
    terminate() {return this.thread.terminate();}
};
const runner = await import('./vba-runner.js');
let sameNameDownloadCase;
try {
    for (const group of groups) {
        runner.closeVbaProject();
        let book = seed(group);
        for (const item of exactCases(group)) {
            const outcome = await runner.runVbaInWorker(book, group.sourceText, item.procedure, item.arguments, 0, fileName(group), options(group), group.bytes);
            verify(outcome, item, group);
            if (group === groups.at(-1) && item === group.cases[0]) sameNameDownloadCase = {group, item, outcome};
            book = JSON.parse(JSON.stringify(outcome.workbook));
        }
    }
    {
        assert.ok(sameNameDownloadCase);
        const {group, item, outcome} = sameNameDownloadCase;
        verify(outcome, item, group);
        const previousDocument = globalThis.document;
        const element = () => ({children: [], append(child) {this.children.push(child);}, replaceChildren(...children) {this.children = children;}});
        globalThis.document = {createElement: element};
        const container = element(), downloads = bindVbaFileDownloads(container);
        try {
            downloads.show(outcome);
            assert.equal(container.children.length, 2);
            for (const child of container.children) {
                const link = child.children[0];
                assert.equal(link.download, 'same.txt');
                const expected = item.files.find(file => file.path === link.title);
                assert.ok(expected, link.title);
                assert.deepEqual([...new Uint8Array(await (await fetch(link.href)).arrayBuffer())], expected.bytes);
            }
        } finally {downloads.clear();globalThis.document = previousDocument;}
    }
    const group = groups[0], item = exactCases(group)[0], target = item.target;
    runner.closeVbaProject();
    const copied = await runner.runVbaInWorker(seed(group), group.sourceText, item.procedure, item.arguments, 0, fileName(group), options(group), group.bytes);
    const file = copied.files[0], length = file.bytes.length;
    runner.closeVbaProject();
    const source = `${group.sourceText}\nPublic Function InspectCopiedFile(ByVal target As String) As Long\n InspectCopiedFile = FileLen(target)\nEnd Function\nPublic Sub EditCopiedFile(ByVal target As String)\n Dim marker As Byte\n marker = 65\n Open target For Binary As #1\n Put #1, 1, marker\n Close #1\nEnd Sub\nPublic Sub RenameCopiedFile(ByVal target As String, ByVal renamed As String)\n Name target As renamed\nEnd Sub\nPublic Sub DeleteCopiedFile(ByVal target As String)\n Kill target\nEnd Sub\n`;
    const run = (book, procedure, args) => runner.runVbaInWorker(book, source, procedure, args, 0, fileName(group), options(group), group.bytes);
    const inspected = await run(copied.workbook, 'InspectCopiedFile', [target]);
    assert.equal(inspected.result.value, length);
    const edited = await run(inspected.workbook, 'EditCopiedFile', [target]);
    assert.equal(edited.files[0].bytes[0], 65);
    assert.equal(edited.files[0].bytes.length, length);
    const previousDocument = globalThis.document;
    const element = () => ({children: [], append(child) {this.children.push(child);}, replaceChildren(...children) {this.children = children;}});
    globalThis.document = {createElement: element};
    const container = element(), downloads = bindVbaFileDownloads(container);
    try {
        downloads.show(copied);downloads.show(edited);
        const downloaded = new Uint8Array(await (await fetch(container.children[0].children[0].href)).arrayBuffer());
        assert.equal(downloaded[0], 65);assert.equal(downloaded.length, length);
        const renamed = target.replace(/\.xlsm$/, '-renamed.xlsm');
        const moved = await run(edited.workbook, 'RenameCopiedFile', [target, renamed]);
        assert.deepEqual(moved.deleted_files, [target]);assert.equal(moved.files[0].path, renamed);
        downloads.show(moved);assert.equal(container.children.length, 1);
        const removed = await run(moved.workbook, 'DeleteCopiedFile', [renamed]);
        assert.deepEqual(removed.deleted_files, [renamed]);downloads.show(removed);assert.equal(container.children.length, 0);
    } finally {downloads.clear();globalThis.document = previousDocument;}
} finally {runner.closeVbaProject();globalThis.Worker = previousWorker;await Promise.all(workers.map(worker => worker.terminate()));}
console.log(JSON.stringify({passed: true, native_cases: checked, direct: true, session: true, worker: true, restart_file_persistence: true, actual_binary_downloads: true, binary_edits: true, rename_and_delete: true}));
