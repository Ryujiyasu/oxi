let activeExecution = null;

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

export function runVbaInWorker(workbook, source, procedure, args, activeSheet) {
    if (activeExecution) {
        return Promise.reject(new Error('A VBA macro is already running.'));
    }
    return new Promise((resolve, reject) => {
        const worker = new Worker(new URL('./vba-worker.js', import.meta.url), { type: 'module' });
        const execution = { worker, reject };
        activeExecution = execution;
        worker.addEventListener('message', event => {
            if (activeExecution !== execution) return;
            activeExecution = null;
            worker.terminate();
            if (event.data?.ok) {
                resolve(event.data.execution);
            } else {
                reject(new Error(event.data?.error || 'VBA worker failed.'));
            }
        });
        worker.addEventListener('error', event => {
            if (activeExecution !== execution) return;
            activeExecution = null;
            worker.terminate();
            reject(new Error(event.message || 'VBA worker failed.'));
        });
        worker.postMessage({ workbook, source, procedure, args, activeSheet });
    });
}

export function cancelActiveVbaExecution() {
    if (!activeExecution) return false;
    const { worker, reject } = activeExecution;
    activeExecution = null;
    worker.terminate();
    const error = new Error('VBA execution cancelled.');
    error.name = 'AbortError';
    reject(error);
    return true;
}
