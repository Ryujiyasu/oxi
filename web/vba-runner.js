let activeExecution = null;

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
