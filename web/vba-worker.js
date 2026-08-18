import init, { run_spreadsheet_vba } from './oxidocs_wasm.js';

const ready = init();

self.addEventListener('message', async event => {
    const { workbook, source, procedure, args, activeSheet } = event.data;
    try {
        await ready;
        const execution = run_spreadsheet_vba(
            workbook,
            source,
            procedure,
            args,
            activeSheet,
        );
        self.postMessage({ ok: true, execution });
    } catch (error) {
        self.postMessage({
            ok: false,
            error: error?.message || String(error),
        });
    }
});
