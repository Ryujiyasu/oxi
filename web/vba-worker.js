import init, { SpreadsheetVbaSession, spreadsheet_vba_project_name, spreadsheet_vba_source } from './oxidocs_wasm.js';

const ready = init();
const project = ready.then(() => new SpreadsheetVbaSession());

self.addEventListener('message', async event => {
    const { workbook, source, procedure, args, activeSheet, fileName, options, packageBytes, reference, member, callKind, references, workbookIdentity, action = 'run' } = event.data;
    try {
        await ready;
        const session = await project;
        if (action === 'select-workbook') {
            if (typeof workbookIdentity !== 'string') throw new Error('Invalid workbook identity.');
            const selection = workbook === undefined ? session.select_workbook(workbookIdentity)
                : session.select_workbook_with_edits(workbookIdentity, workbook, activeSheet);
            self.postMessage({ ok: true, selection });
            return;
        }
        const { projectNameCodePage, projectSourceCodePage, ...runtimeOptions } = options || {};
        const inputPackage = packageBytes ?? runtimeOptions.fileContext?.packageBytes;
        if (projectSourceCodePage != null && (!Number.isInteger(projectSourceCodePage) || projectSourceCodePage < 1 || projectSourceCodePage > 65535)) {
            throw new Error('Invalid VBA source code page.');
        }
        if (runtimeOptions.hostAnsiCodePage != null && ![932, 1252].includes(runtimeOptions.hostAnsiCodePage)) {
            throw new Error('Unsupported VBA host ANSI code page.');
        }
        if (runtimeOptions.projectName == null && inputPackage) {
            if (projectNameCodePage != null && (!Number.isInteger(projectNameCodePage) || projectNameCodePage < 1 || projectNameCodePage > 65535)) {
                throw new Error('Invalid VBA project name code page.');
            }
            runtimeOptions.projectName = spreadsheet_vba_project_name(inputPackage, projectNameCodePage);
        }
        if (action === 'compile') {
            const compilationOptions = { ...runtimeOptions };
            delete compilationOptions.hostAnsiCodePage;
            delete compilationOptions.worksheetDbcsLanguage;
            const compilation = session.compile(source, { ...compilationOptions,
                documentModules: [...(inputPackage ? session.document_modules(inputPackage) : []),
                    ...(runtimeOptions.documentModules || [])],
                referenceSymbols: [...(workbook?.sheets || []).map(sheet => ({ name: sheet.codeName })),
                    ...(runtimeOptions.referenceSymbols || [])] });
            self.postMessage({ ok: compilation.compiled, compilation, error: compilation.error?.message });
            return;
        }
        if (!['run', 'save', 'invoke', 'release', 'release-all'].includes(action)) throw new Error('Invalid VBA worker action.');
        if (action === 'invoke' && (typeof reference !== 'string' || typeof member !== 'string'
            || !Number.isSafeInteger(callKind))) throw new Error('Invalid VBA member invocation.');
        if (action === 'release' && (!Array.isArray(references)
            || !references.every(reference => typeof reference === 'string'))) throw new Error('Invalid VBA object references.');
        // The runtime imports package bindings for a new or persisted project.
        // After macro-free SaveAs it retains the still-loaded project's bindings.
        const calculationSource = action === 'save' && !source.trim() && inputPackage
            ? spreadsheet_vba_source(inputPackage, projectSourceCodePage) || '' : source;
        if (inputPackage) {
            runtimeOptions.fileContext = { ...runtimeOptions.fileContext, packageBytes: inputPackage };
        }
        const execution = action === 'invoke'
            ? session.invoke_member(workbook, source, { reference, member, kind: callKind, args }, activeSheet, fileName, runtimeOptions)
            : action === 'release'
            ? session.release_objects(workbook, references, activeSheet, fileName, runtimeOptions)
            : action === 'release-all'
            ? session.release_all_objects(workbook, activeSheet, fileName, runtimeOptions)
            : action === 'save'
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
        // The completed macro still belongs to its original source owner.
        // Move the editor only after that owner's frames have unwound.
        if (execution.workbook_closed && execution.active_workbook !== undefined) {
            execution.workbook_selection = session.select_workbook(execution.active_workbook);
        }
        self.postMessage({ ok: !execution.error, execution, error: execution.error?.message });
    } catch (error) {
        self.postMessage({
            ok: false,
            error: error?.message || String(error),
        });
    }
});
