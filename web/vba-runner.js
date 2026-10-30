let activeExecution = null;
let applicationState;
let fileSystemState;
let projectWorker;
let projectWorkerType;
let compilationWorker;
let compilationWorkerType;
let sourceHistory;
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
            workersWithFiles.add(worker);
            sourceHistory = execution.source?.trim() ? { source: execution.source, edits: [] } : undefined;
        }
        if (event.data?.ok) {
            execution.resolve(event.data.compilation ?? event.data.execution);
        } else {
            const error = new Error(event.data?.error || 'VBA worker failed.');
            if (event.data?.compilation) {
                error.name = 'VbaCompileError';
                error.compilation = event.data.compilation;
            }
            if (event.data?.execution) {
                error.name = 'VbaExecutionError';
                error.execution = event.data.execution;
            }
            execution.reject(error);
        }
    });
    worker.addEventListener('error', event => {
        if (projectWorker === worker) projectWorker = null;
        if (compilationWorker === worker) compilationWorker = null;
        if (!compiling) sourceHistory = undefined;
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
    args.forEach((value, index) => {
        const supported = value === null
            || typeof value === 'boolean'
            || typeof value === 'string'
            || (typeof value === 'number' && Number.isFinite(value));
        if (!supported) {
            throw new Error(`VBA argument ${index + 1} must be null, a boolean, a number, or a string.`);
        }
    });
    return args;
}

export function runVbaInWorker(workbook, source, procedure, args, activeSheet, fileName, options = {}, packageBytes) {
    return startVbaWorker('run', workbook, source, procedure, args, activeSheet, fileName, options, packageBytes);
}

export function prepareVbaSaveInWorker(workbook, source, activeSheet, fileName, options = {}, packageBytes) {
    return startVbaWorker('save', workbook, source, undefined, [], activeSheet, fileName, options, packageBytes);
}

export function compileVbaInWorker(source, options = {}, workbook, packageBytes) {
    const context = workbook && { sheets: workbook.sheets.map(sheet => ({
        codeName: sheet.code_name ?? sheet.codeName,
    })).filter(sheet => typeof sheet.codeName === 'string' && sheet.codeName.length) };
    return startVbaWorker('compile', context, source, undefined, [], undefined, undefined, options, packageBytes);
}

function startVbaWorker(action, workbook, source, procedure, args, activeSheet, fileName, options, packageBytes) {
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
                    ...(sourceHistory?.source === source && sourceHistory.edits.length ? { sourceEdits: sourceHistory.edits } : {}), ...options, ...(state ? {applicationState: state} : {}),
                    ...(namespace != null ? {fileSystemState: namespace} : {}),
                    ...(fileContext ? { fileContext } : {}) },
                packageBytes,
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
    if (action !== 'compile') sourceHistory = undefined;
    const error = new Error(action === 'compile' ? 'VBA compilation cancelled.' : 'VBA execution cancelled.');
    error.name = 'AbortError';
    reject(error);
    return true;
}

export function closeVbaProject() {
    sourceHistory = undefined;
    let closed = cancelActiveVbaExecution();
    if (projectWorker) { projectWorker.terminate(); projectWorker = null; closed = true; }
    if (compilationWorker) { compilationWorker.terminate(); compilationWorker = null; closed = true; }
    return closed;
}
