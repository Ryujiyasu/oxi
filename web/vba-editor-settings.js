// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

const hostCodePages = new Set(['', '932', '1252']);
const errorLanguages = new Set(['', 'en-US', 'ja-JP']);
const storageKey = 'oxi.vba.environment.v1';

export function preferredVbaProcedure(procedures) {
    return procedures.find(procedure => procedure.visibility === 'public'
        && procedure.required_parameter_count === 0 && procedure.module_kind !== 'document')
        || procedures.find(procedure => procedure.visibility === 'public' && procedure.required_parameter_count === 0)
        || procedures.find(procedure => procedure.required_parameter_count === 0)
        || procedures[0];
}

export function vbaEditorOptions(settings) {
    if (!hostCodePages.has(settings.hostCodePage)) throw new Error('Unsupported Windows character set.');
    if (!errorLanguages.has(settings.errorLanguage)) throw new Error('Unsupported VBA error language.');
    if (typeof settings.helpFile !== 'string') throw new Error('Invalid VBA help file.');
    return {
        ...(settings.hostCodePage ? { hostAnsiCodePage: Number(settings.hostCodePage) } : {}),
        errorLanguage: settings.errorLanguage || 'en-US',
        ...(settings.helpFile ? { vbaHelpFile: settings.helpFile } : {}),
    };
}

// Both macro editors use explicit target settings. Neither browser language
// nor the workbook's source encoding identifies the Windows runtime locale.
export function bindVbaEditorSettings({ hostCodePage, errorLanguage, helpFile,
    sourceInput, loadSourceButton, readProjectSource, onSourceLoaded,
    onProfileChange, report, messages }) {
    let busyCount = 0;
    const settings = () => ({ hostCodePage: hostCodePage.value,
        errorLanguage: errorLanguage.value, helpFile: helpFile.value });
    const options = () => vbaEditorOptions(settings());
    try {
        const saved = JSON.parse(globalThis.localStorage?.getItem(storageKey) || 'null');
        if (saved && typeof saved.helpFile === 'string') {
            vbaEditorOptions(saved);
            hostCodePage.value = saved.hostCodePage;
            errorLanguage.value = saved.errorLanguage;
            helpFile.value = saved.helpFile;
        }
    } catch { /* Keep the explicit defaults if storage is unavailable or invalid. */ }

    const changed = () => {
        if (busyCount) return;
        try {
            options();
            // Changing the target starts a fresh execution environment. The
            // editor text stays in place, including edits that have not run.
            onProfileChange();
            try { globalThis.localStorage?.setItem(storageKey, JSON.stringify(settings())); } catch { /* Session settings still apply. */ }
        } catch (error) { report(error.message || String(error), true); }
    };
    for (const control of [hostCodePage, errorLanguage, helpFile]) control.addEventListener('change', changed);

    loadSourceButton.addEventListener('click', () => {
        if (busyCount) return;
        try {
            const source = readProjectSource(options().hostAnsiCodePage);
            if (typeof source !== 'string' || !source.trim()) {
                report(messages.noSource(), false);
                return;
            }
            const previous = sourceInput.value;
            sourceInput.value = source;
            onSourceLoaded(previous, sourceInput.value);
            report(messages.loaded(), false);
        } catch (error) { report(error.message || String(error), true); }
    });

    return {
        options,
        beginOperation() {
            const update = () => {
                for (const control of [hostCodePage, errorLanguage, helpFile, loadSourceButton]) control.disabled = busyCount > 0;
            };
            busyCount++;
            update();
            let ended = false;
            return () => {
                if (ended) return;
                ended = true;
                busyCount--;
                update();
            };
        },
    };
}
