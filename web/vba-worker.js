import init, { run_spreadsheet_vba } from './oxidocs_wasm.js';

const ready = init();

self.addEventListener('message', async event => {
    const { workbook, source, procedure, args, activeSheet, fileName } = event.data;
    try {
        await ready;
        const execution = run_spreadsheet_vba(
            workbook,
            source,
            procedure,
            args,
            activeSheet,
            // What `ActiveWorkbook.Name` answers. Nothing here leaves the
            // workbook calling itself Book1, which is what Excel calls one
            // that was never saved.
            fileName,
        );
        self.postMessage({ ok: true, execution });
    } catch (error) {
        self.postMessage({
            ok: false,
            error: error?.message || String(error),
        });
    }
});
