import init, { SpreadsheetVbaSession, spreadsheet_vba_project_name, spreadsheet_vba_source } from './oxidocs_wasm.js';

const ready = init();
const project = ready.then(() => new SpreadsheetVbaSession());

self.addEventListener('message', async event => {
    const { workbook, source, procedure, args, activeSheet, fileName, options, packageBytes, action = 'run' } = event.data;
    try {
        await ready;
        const session = await project;
        const { projectNameCodePage, projectSourceCodePage, ...runtimeOptions } = options || {};
        if (projectSourceCodePage != null && (!Number.isInteger(projectSourceCodePage) || projectSourceCodePage < 1 || projectSourceCodePage > 65535)) {
            throw new Error('Invalid VBA source code page.');
        }
        if (runtimeOptions.hostAnsiCodePage != null && ![932, 1252].includes(runtimeOptions.hostAnsiCodePage)) {
            throw new Error('Unsupported VBA host ANSI code page.');
        }
        if (runtimeOptions.projectName == null && packageBytes) {
            if (projectNameCodePage != null && (!Number.isInteger(projectNameCodePage) || projectNameCodePage < 1 || projectNameCodePage > 65535)) {
                throw new Error('Invalid VBA project name code page.');
            }
            runtimeOptions.projectName = spreadsheet_vba_project_name(packageBytes, projectNameCodePage);
        }
        if (action === 'compile') {
            const compilationOptions = { ...runtimeOptions };
            delete compilationOptions.hostAnsiCodePage;
            delete compilationOptions.worksheetDbcsLanguage;
            const compilation = session.compile(source, { ...compilationOptions,
                documentModules: [...(packageBytes ? session.document_modules(packageBytes) : []),
                    ...(runtimeOptions.documentModules || [])],
                referenceSymbols: [...(workbook?.sheets || []).map(sheet => ({ name: sheet.codeName })),
                    ...(runtimeOptions.referenceSymbols || [])] });
            self.postMessage({ ok: compilation.compiled, compilation, error: compilation.error?.message });
            return;
        }
        if (action !== 'run' && action !== 'save') throw new Error('Invalid VBA worker action.');
        if (runtimeOptions.documentModules == null && packageBytes) {
            runtimeOptions.documentModules = session.document_modules(packageBytes);
        }
        const calculationSource = action === 'save' && !source.trim() && packageBytes
            ? spreadsheet_vba_source(packageBytes, projectSourceCodePage) || '' : source;
        if (packageBytes) {
            runtimeOptions.fileContext = { ...runtimeOptions.fileContext, packageBytes };
        }
        const execution = action === 'save'
            ? session.prepare_save(workbook, calculationSource, activeSheet, fileName, runtimeOptions)
            : session.run(
            workbook,
            source,
            procedure,
            args,
            activeSheet,
            // What `ActiveWorkbook.Name` answers. Nothing here leaves the
            // workbook calling itself Book1, which is what Excel calls one
            // that was never saved.
            fileName,
            runtimeOptions,
        );
        self.postMessage({ ok: !execution.error, execution, error: execution.error?.message });
    } catch (error) {
        self.postMessage({
            ok: false,
            error: error?.message || String(error),
        });
    }
});
