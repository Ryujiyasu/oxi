let activeExecution = null;
let applicationState;
let fileSystemState;
let projectWorker;
let projectWorkerType;
let compilationWorker;
let compilationWorkerType;
let sourceHistory;
let selectedWorkbookIdentity = '0';
const workbookSourceHistory = new Map();
let objectReferences = [];
// Project resets discard VBA globals and open handles, while produced files
// remain part of the containing application's filesystem.
const producedFiles = new Map();
const fileFolders = new Set();
const workersWithFiles = new WeakSet();
const fileKey = path => path.trim().replaceAll('/', '\\').toLowerCase();

function copyApplicationState(state) {
    if (!state || typeof state !== 'object' || Array.isArray(state)) return state;
    const directory = state.directoryState;
    return {...state, ...(directory && typeof directory === 'object' && !Array.isArray(directory) ? {directoryState: {
        ...directory, directories: Array.isArray(directory.directories) ? [...directory.directories] : directory.directories,
    }} : {})};
}

function copyFileSystemState(state) {
    if (!state || typeof state !== 'object' || Array.isArray(state)) return state;
    const paths = value => Array.isArray(value) ? [...value] : value;
    return {...state, folders: paths(state.folders), mountedFolders: paths(state.mountedFolders), mountedFiles: paths(state.mountedFiles),
        ...(Object.hasOwn(state, 'volumes') ? {volumes: Array.isArray(state.volumes)
            ? state.volumes.map(entry => entry && typeof entry === 'object' && !Array.isArray(entry) ? {...entry} : entry) : state.volumes} : {}),
        ...(Object.hasOwn(state, 'attributes') ? {attributes: Array.isArray(state.attributes)
            ? state.attributes.map(entry => entry && typeof entry === 'object' && !Array.isArray(entry) ? {...entry} : entry) : state.attributes} : {}),
        files: Array.isArray(state.files) ? state.files.map(file => file && typeof file === 'object' && !Array.isArray(file)
            ? {...file, bytes: Array.isArray(file.bytes) ? [...file.bytes] : ArrayBuffer.isView(file.bytes) && typeof file.bytes.slice === 'function' ? file.bytes.slice() : file.bytes}
            : file) : state.files};
}

function rememberProducedFiles(execution) {
    for (const path of execution.deleted_files || []) producedFiles.delete(fileKey(path));
    for (const file of execution.files || []) {
        producedFiles.set(fileKey(file.path), { path: file.path, bytes: new Uint8Array(file.bytes) });
        const path = file.path.replaceAll('/', '\\');
        const slash = path.lastIndexOf('\\');
        if (slash > 0) fileFolders.add(path.slice(0, slash));
    }
}

function retainedFileContext(context, mountFiles) {
    if (!mountFiles) return context;
    if (!context && !producedFiles.size && !fileFolders.size) return undefined;
    const files = new Map((context?.files || []).map(file => [fileKey(file.path), file]));
    for (const [key, file] of producedFiles) files.set(key, file);
    const folders = new Set([...fileFolders, ...(context?.folders || [])]);
    return { ...context, files: Array.from(files.values()), folders: Array.from(folders) };
}

export function getVbaSourceForSave(fallback = '') {
    return sourceHistory?.source ?? fallback;
}

export function getVbaObjectReferences() { return objectReferences.slice(); }

// Preserve editing operations that cannot be reconstructed from final text,
// including replacing a declaration and then undoing it before running again.
export function recordVbaSourceChange(previous, next, selection) {
    if (activeExecution || !sourceHistory || sourceHistory.source !== previous) {
        sourceHistory = undefined;
        return false;
    }
    const lines = text => {
        const result = text.replaceAll('\r\n', '\n').split('\n');
        if (result.at(-1) === '') result.pop();
        return result;
    };
    const before = lines(previous), after = lines(next);
    let first = 0;
    while (first < before.length && first < after.length && before[first] === after[first]) first++;
    let oldEnd = before.length, newEnd = after.length;
    while (oldEnd > first && newEnd > first && before[oldEnd - 1] === after[newEnd - 1]) { oldEnd--; newEnd--; }
    if (selection && Number.isInteger(selection.start) && Number.isInteger(selection.end)
        && selection.start >= 0 && selection.end > selection.start && selection.end <= previous.length) {
        const selectedFirst = previous.slice(0, selection.start).split('\n').length - 1;
        const selectedLast = previous.slice(0, selection.end - 1).split('\n').length - 1;
        const unchanged = first === before.length && first === after.length;
        first = unchanged ? selectedFirst : Math.min(first, selectedFirst);
        oldEnd = unchanged ? selectedLast + 1 : Math.max(oldEnd, selectedLast + 1);
        newEnd = Math.min(after.length, oldEnd + after.length - before.length);
    } else if (previous === next) { return false; }
    const inserted = after.slice(first, newEnd);
    sourceHistory.edits.push({ line: first + 1, deleteLines: oldEnd - first,
        text: inserted.length ? inserted.join('\n') + '\n' : '' });
    sourceHistory.source = next;
    return true;
}

function getProjectWorker(action = 'run') {
    const compiling = action === 'compile';
    const cached = compiling ? compilationWorker : projectWorker;
    const cachedType = compiling ? compilationWorkerType : projectWorkerType;
    if (cached && cachedType !== globalThis.Worker) closeVbaProject();
    else if (cached) return cached;
    const worker = new Worker(new URL('./vba-worker.js', import.meta.url), { type: 'module' });
    if (compiling) { compilationWorker = worker; compilationWorkerType = globalThis.Worker; }
    else { projectWorker = worker; projectWorkerType = globalThis.Worker; }
    worker.addEventListener('message', event => {
        const execution = activeExecution;
        if (!execution || execution.worker !== worker) return;
        activeExecution = null;
        if (event.data?.execution?.application_state) {
            applicationState = copyApplicationState(event.data.execution.application_state);
        }
        if (event.data?.execution) {
            if (event.data.execution.file_system_state) fileSystemState = copyFileSystemState(event.data.execution.file_system_state);
            rememberProducedFiles(event.data.execution);
            const live = new Set((event.data.execution.other_workbooks || []).map(book => book.identity));
            if (!event.data.execution.workbook_closed) live.add(selectedWorkbookIdentity);
            for (const identity of workbookSourceHistory.keys()) {
                if (!live.has(identity)) workbookSourceHistory.delete(identity);
            }
            workersWithFiles.add(worker);
            if (Array.isArray(event.data.execution.reference_state?.tokens)) {
                objectReferences = event.data.execution.reference_state.tokens.slice();
            }
            if (event.data.execution.workbook_closed) {
                workbookSourceHistory.delete(selectedWorkbookIdentity);
                sourceHistory = undefined;
            } else if (execution.action !== 'release' && execution.action !== 'release-all') {
                sourceHistory = execution.source?.trim() ? { source: execution.source, edits: [] } : undefined;
            }
        }
        const selected = event.data?.selection ?? event.data?.execution?.workbook_selection;
        if (selected) {
            const selection = selected;
            if (sourceHistory) workbookSourceHistory.set(selectedWorkbookIdentity, sourceHistory);
            else workbookSourceHistory.delete(selectedWorkbookIdentity);
            selectedWorkbookIdentity = selection.selected.identity;
            sourceHistory = workbookSourceHistory.get(selectedWorkbookIdentity)
                ?? (selection.selected.source.trim() ? {source: selection.selected.source, edits: []} : undefined);
            if (sourceHistory) selection.selected.source = sourceHistory.source;
            applicationState = copyApplicationState(selection.application_state);
            fileSystemState = copyFileSystemState(selection.file_system_state);
            workersWithFiles.add(worker);
        }
        if (event.data?.ok) {
            execution.resolve(execution.action === 'compile-project' ? event.data.execution
                : event.data.selection ?? event.data.compilation ?? event.data.execution);
        } else {
            const error = new Error(event.data?.error || 'VBA worker failed.');
            if (event.data?.compilation) {
                error.name = 'VbaCompileError';
                error.compilation = event.data.compilation;
            }
            if (event.data?.execution) {
                if (!event.data.compilation) error.name = 'VbaExecutionError';
                error.execution = event.data.execution;
            }
            execution.reject(error);
        }
    });
    worker.addEventListener('error', event => {
        if (projectWorker === worker) projectWorker = null;
        if (compilationWorker === worker) compilationWorker = null;
        if (!compiling) { sourceHistory = undefined; selectedWorkbookIdentity = '0'; workbookSourceHistory.clear(); objectReferences = []; }
        worker.terminate();
        const execution = activeExecution;
        if (!execution || execution.worker !== worker) return;
        activeExecution = null;
        execution.reject(new Error(event.message || 'VBA worker failed.'));
    });
    return worker;
}

export function parseVbaArguments(text) {
    const source = text.trim();
    if (!source) return [];
    let args;
    try {
        args = JSON.parse(source);
    } catch (error) {
        throw new Error(`VBA arguments must be valid JSON: ${error.message}`);
    }
    if (!Array.isArray(args)) {
        throw new Error('VBA arguments must be a JSON array.');
    }
        const typedScalar = value => {
            if (!value || typeof value !== 'object' || Array.isArray(value)
                || typeof value.type !== 'string'
                || Object.keys(value).some(key => key !== 'type' && key !== 'value')) return false;
            const units = ['empty', 'missing', 'null', 'nothing', 'null_string'];
            if (units.includes(value.type)) return !Object.hasOwn(value, 'value');
            if (!Object.hasOwn(value, 'value')) return false;
            const payload = value.value;
            switch (value.type) {
                case 'boolean': return typeof payload === 'boolean';
                case 'byte':
                case 'integer':
                case 'long':
                case 'error': return Number.isSafeInteger(payload);
                case 'single':
                case 'double':
                case 'date': return typeof payload === 'number' && Number.isFinite(payload);
                case 'long_long': return typeof payload === 'string' && /^-?\d+$/.test(payload);
                case 'currency':
                case 'decimal': return typeof payload === 'string' && /^-?\d+(?:\.\d+)?$/.test(payload);
                case 'string': return typeof payload === 'string';
                case 'utf16': return Array.isArray(payload)
                    && payload.every(unit => Number.isInteger(unit) && unit >= 0 && unit <= 65535);
                case 'object':
                    return payload && typeof payload === 'object' && !Array.isArray(payload)
                        && typeof payload.reference === 'string'
                        && Object.keys(payload).every(key => ['reference', 'kind', 'handle'].includes(key))
                        && (!Object.hasOwn(payload, 'kind') || typeof payload.kind === 'string')
                        && (!Object.hasOwn(payload, 'handle') || (Number.isSafeInteger(payload.handle) && payload.handle >= 0)
                            || (typeof payload.handle === 'string' && /^\d+$/.test(payload.handle)));
                case 'array': {
                    if (!payload || typeof payload !== 'object' || Array.isArray(payload)
                        || Object.keys(payload).some(key => !['dimensions', 'values', 'element_type', 'resizable'].includes(key))
                        || !Array.isArray(payload.dimensions) || !Array.isArray(payload.values)
                        || payload.dimensions.length > 65535
                        || !['variant', 'boolean', 'byte', 'integer', 'long', 'long_long', 'single', 'double', 'currency', 'date', 'error', 'string', 'object'].includes(payload.element_type)
                        || (Object.hasOwn(payload, 'resizable') && typeof payload.resizable !== 'boolean')) return false;
                    for (const dimension of payload.dimensions) {
                        if (!dimension || typeof dimension !== 'object' || Array.isArray(dimension)
                            || Object.keys(dimension).length !== 2
                            || !Object.hasOwn(dimension, 'lower_bound') || !Object.hasOwn(dimension, 'length')
                            || !Number.isInteger(dimension.lower_bound) || dimension.lower_bound < -2147483648 || dimension.lower_bound > 2147483647
                            || !Number.isInteger(dimension.length) || dimension.length < 0 || dimension.length > 4294967295) return false;
                    }
                    const count = !payload.dimensions.length || payload.dimensions.some(dimension => dimension.length === 0)
                        ? 0 : payload.dimensions.reduce((total, dimension) => total * dimension.length, 1);
                    return Number.isSafeInteger(count) && count === payload.values.length && payload.values.every(supportedValue);
                }
                default: return false;
            }
        };
        const supportedValue = value => value === null
            || typeof value === 'boolean'
            || typeof value === 'string'
            || (typeof value === 'number' && Number.isFinite(value))
            || typedScalar(value);
    args.forEach((value, index) => {
        const supported = supportedValue(value);
        if (!supported) {
            throw new Error(`VBA argument ${index + 1} must be null, a boolean, a number, a string, or a typed VBA scalar, object reference, or array.`);
        }
    });
    return args;
}

export function runVbaInWorker(workbook, source, procedure, args, activeSheet, fileName, options = {}, packageBytes) {
    return startVbaWorker('run', workbook, source, procedure, args, activeSheet, fileName, options, packageBytes);
}

export function invokeVbaMemberInWorker(workbook, source, reference, member, callKind, args,
    activeSheet, fileName, options = {}, packageBytes) {
    return startVbaWorker('invoke', workbook, source, undefined, args, activeSheet, fileName,
        options, packageBytes, { reference, member, callKind });
}

export function releaseAllVbaObjectsInWorker(workbook, activeSheet, fileName, options = {}, packageBytes) {
    return startVbaWorker('release-all', workbook, undefined, undefined, [], activeSheet, fileName, options, packageBytes);
}

export function releaseVbaObjectsInWorker(workbook, source, references,
    activeSheet, fileName, options = {}, packageBytes) {
    return startVbaWorker('release', workbook, source, undefined, [], activeSheet, fileName,
        options, packageBytes, { references });
}

// An execution result is a caller-owned value. Explicitly release its leases
// after the consumer is done; nested array/record elements are released backwards.
export function releaseVbaResultInWorker(workbook, source, result,
    activeSheet, fileName, options = {}, packageBytes) {
    const references = [];
    const collect = value => {
        if (value?.type === 'object' && typeof value.value?.reference === 'string') references.push(value.value.reference);
        if (value?.type === 'array') for (const item of value.value.values) collect(item);
        if (value?.type === 'record') for (const [, item] of value.value.fields) collect(item);
    };
    collect(result);
    return releaseVbaObjectsInWorker(workbook, source, [...new Set(references)].reverse(),
        activeSheet, fileName, options, packageBytes);
}

// The same worker owns every open workbook. A selection moves the editor's
// view while preserving unsent source edits and all caller-owned references.
export function selectVbaWorkbookInWorker(workbookIdentity, workbook, activeSheet, editRevision) {
    if (typeof workbookIdentity !== 'string') return Promise.reject(new Error('Invalid workbook identity.'));
    return startVbaWorker('select-workbook', workbook, undefined, undefined, [], activeSheet, undefined, {}, undefined,
        {workbookIdentity, editRevision});
}

// Untitled template copies keep their VBA project or legacy macro sheets.
export function vbaWorkbookExtension(name, workbook, hasProject) {
    const template = /\.(xltm|xltx)$/i.exec(name || '');
    if (template) return template[1].toLowerCase();
    return /\.xlsm$/i.test(name || '') || hasProject
        || workbook?.sheets?.some(sheet => sheet.kind === 'macro_sheet' || sheet.kind === 'international_macro_sheet')
        ? 'xlsm' : 'xlsx';
}

export function vbaWorkbookMimeType(extension) {
    return {
        xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        xlsm: 'application/vnd.ms-excel.sheet.macroEnabled.12',
        xltx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.template',
        xltm: 'application/vnd.ms-excel.template.macroEnabled.12',
    }[extension] || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
}

// Both editors use the same application worker and keep only their view
// bookkeeping here. The worker remains the authority for each workbook IR.
export function bindVbaWorkbookPicker(control, {capture, apply, busy, report}) {
    let identity = selectedWorkbookIdentity;
    let selecting = false;
    const names = new Map();
    const views = new Map();
    const refresh = () => { control.disabled = selecting || busy() || names.size < 2; };
    const fill = () => {
        control.replaceChildren(...Array.from(names, ([value, label]) => {
            const option = control.ownerDocument.createElement('option');
            option.value = value;
            option.textContent = label;
            return option;
        }));
        control.value = identity;
        refresh();
    };
    const update = (execution, name) => {
        if (identity !== selectedWorkbookIdentity && !execution.workbook_selection) { names.clear(); views.clear(); }
        const previousIdentity = identity;
        identity = selectedWorkbookIdentity;
        names.clear();
        const selection = execution.workbook_selection;
        if (selection) {
            views.delete(previousIdentity);
            identity = selection.selected.identity;
            const previous = identity === previousIdentity ? capture().view : views.get(identity);
            // A rename keeps the editor's position and invocation, while the
            // authoritative saved package supplies its new source baseline.
            const view = identity === previousIdentity && previous
                ? {procedure: previous.procedure, args: previous.args, cell: previous.cell, editRevision: previous.editRevision, savedEditRevision: previous.savedEditRevision} : previous;
            apply(selection, view);
            names.set(identity, selection.selected.name);
            for (const book of selection.other_workbooks) names.set(book.identity, book.name);
        } else if (!execution.workbook_closed) {
            names.set(identity, name || 'Book1');
            for (const book of execution.other_workbooks || []) names.set(book.identity, book.name);
        }
        for (const key of views.keys()) if (!names.has(key)) views.delete(key);
        fill();
    };
    control.addEventListener('change', async () => {
        const wanted = control.value;
        if (selecting || busy() || wanted === identity) { control.value = identity; return; }
        selecting = true;
        refresh();
        const previousIdentity = identity;
        try {
            const current = capture();
            views.set(identity, current.view);
            const selection = await selectVbaWorkbookInWorker(wanted, current.workbook, current.activeSheet, current.editRevision);
            try { apply(selection, views.get(wanted)); }
            catch (error) {
                // Restore the actual invoking owner if rendering/preparation
                // rejects the new view before changing the editor.
                await selectVbaWorkbookInWorker(previousIdentity);
                throw error;
            }
            identity = wanted;
            names.clear();
            names.set(identity, selection.selected.name);
            for (const book of selection.other_workbooks) names.set(book.identity, book.name);
        } catch (error) { report(error?.message || String(error), true); }
        finally { selecting = false; fill(); }
    });
    return {
        update, refresh,
        reset(name) {
            identity = '0'; names.clear(); views.clear();
            if (name) names.set(identity, name);
            fill();
        },
    };
}

export function prepareVbaSaveInWorker(workbook, source, activeSheet, fileName, options = {}, packageBytes) {
    return startVbaWorker('save', workbook, source, undefined, [], activeSheet, fileName, options, packageBytes);
}

// Save and SaveCopyAs use the same runtime callbacks as VBA commands. An
// AfterSave error can coexist with a file that was already written.
export async function saveVbaWorkbookInWorker(workbook, source, activeSheet, fileName, options = {}, packageBytes, copyName) {
    try {
        return await startVbaWorker('save-workbook', workbook, source, undefined, [], activeSheet, fileName,
            options, packageBytes, { copyName });
    } catch (error) {
        if (error?.execution) return error.execution;
        throw error;
    }
}

export function compileVbaInWorker(source, options = {}, workbook, packageBytes) {
    const context = workbook && { sheets: workbook.sheets.map(sheet => ({
        codeName: sheet.code_name ?? sheet.codeName,
    })).filter(sheet => typeof sheet.codeName === 'string' && sheet.codeName.length) };
    return startVbaWorker('compile', context, source, undefined, [], undefined, undefined, options, packageBytes);
}

// Explicit project compilation shares the workbook worker and its live storage.
// `compileVbaInWorker` remains the independent, read-only source validator.
export function compileVbaProjectInWorker(workbook, source, activeSheet, fileName, options = {}, packageBytes) {
    return startVbaWorker('compile-project', workbook, source, undefined, [], activeSheet, fileName, options, packageBytes);
}

function startVbaWorker(action, workbook, source, procedure, args, activeSheet, fileName, options, packageBytes, actionData = {}) {
    if (activeExecution) {
        return Promise.reject(new Error('A VBA macro is already running.'));
    }
    return new Promise((resolve, reject) => {
        const worker = getProjectWorker(action);
        const execution = { worker, resolve, reject, source, action };
        activeExecution = execution;
        const errorLanguage = globalThis.navigator?.language?.toLowerCase().startsWith('ja') ? 'ja-JP' : 'en-US';
        try {
            const namespace = action === 'compile' ? undefined : copyFileSystemState(options.fileSystemState ?? (!workersWithFiles.has(worker) ? fileSystemState : undefined));
            const fileContext = action === 'compile' ? undefined : namespace || fileSystemState ? options.fileContext : retainedFileContext(options.fileContext, !workersWithFiles.has(worker));
            const state = copyApplicationState(options.applicationState ?? applicationState);
            if (state && typeof state === 'object' && !Array.isArray(state) && state.directoryState == null && applicationState?.directoryState) {
                state.directoryState = copyApplicationState(applicationState).directoryState;
            }
            worker.postMessage({
                action, workbook, source, procedure, args, activeSheet, fileName,
                options: action === 'compile' ? { errorLanguage, ...options } : { errorLanguage,
                    ...(action !== 'release' && action !== 'release-all' && sourceHistory?.source === source && sourceHistory.edits.length ? { sourceEdits: sourceHistory.edits } : {}), ...options, ...(state ? {applicationState: state} : {}),
                    ...(namespace != null ? {fileSystemState: namespace} : {}),
                    ...(fileContext ? { fileContext } : {}) },
                packageBytes, ...actionData,
            });
        } catch (error) {
            activeExecution = null;
            // A failed structured clone has not reached the worker. Its
            // project is still valid and can accept the next invocation.
            reject(error);
        }
    });
}

export function cancelActiveVbaExecution() {
    if (!activeExecution) return false;
    const { worker, reject, action } = activeExecution;
    activeExecution = null;
    worker.terminate();
    if (projectWorker === worker) projectWorker = null;
    if (compilationWorker === worker) compilationWorker = null;
    if (action !== 'compile') { sourceHistory = undefined; selectedWorkbookIdentity = '0'; workbookSourceHistory.clear(); objectReferences = []; }
    const error = new Error(action === 'compile' ? 'VBA compilation cancelled.' : 'VBA execution cancelled.');
    error.name = 'AbortError';
    reject(error);
    return true;
}

export function closeVbaProject() {
    sourceHistory = undefined;
    selectedWorkbookIdentity = '0';
    workbookSourceHistory.clear();
    objectReferences = [];
    let closed = cancelActiveVbaExecution();
    if (projectWorker) { projectWorker.terminate(); projectWorker = null; closed = true; }
    if (compilationWorker) { compilationWorker.terminate(); compilationWorker = null; closed = true; }
    return closed;
}
