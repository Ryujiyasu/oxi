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

// Standard-font preferences are shared by the editor and its execution worker.
// Keep this module synchronous so the worker can register its first-message listener.
const VBA_FONT_PREFERENCE_FILE = 'oxi-vba-standard-font.bin';
const VBA_FONT_PREFERENCE_BYTES = 4096;

function vbaFontPreferenceChecksum(bytes) {
    let crc = 0xffffffff;
    for (let at = 0; at < bytes.length; at++) {
        crc ^= at >= 8 && at < 12 ? 0 : bytes[at];
        for (let bit = 0; bit < 8; bit++) crc = (crc >>> 1) ^ (0xedb88320 & -(crc & 1));
    }
    return (crc ^ 0xffffffff) >>> 0;
}

function encodeVbaFontPreference(profile) {
    if (profile !== null && (!Array.isArray(profile)
        || profile.length > (VBA_FONT_PREFERENCE_BYTES - 12) / 2
        || !profile.every(unit => Number.isInteger(unit) && unit >= 0 && unit <= 65535))) {
        throw new TypeError('Invalid standard-font preference.');
    }
    const bytes = new Uint8Array(VBA_FONT_PREFERENCE_BYTES), view = new DataView(bytes.buffer);
    view.setUint32(0, 0x5046584f, true);
    view.setUint32(4, profile === null ? 0xffffffff : profile.length, true);
    profile?.forEach((unit, index) => view.setUint16(12 + index * 2, unit, true));
    view.setUint32(8, vbaFontPreferenceChecksum(bytes), true);
    return bytes;
}

function decodeVbaFontPreference(bytes) {
    if (bytes.byteLength !== VBA_FONT_PREFERENCE_BYTES) return undefined;
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    if (view.getUint32(0, true) !== 0x5046584f || view.getUint32(8, true) !== vbaFontPreferenceChecksum(bytes)) return undefined;
    const length = view.getUint32(4, true);
    if (length === 0xffffffff) return null;
    if (length > (VBA_FONT_PREFERENCE_BYTES - 12) / 2) return undefined;
    return Array.from({length}, (_, index) => view.getUint16(12 + index * 2, true));
}

export async function readVbaStandardFontPreference() {
    if (!globalThis.navigator?.storage?.getDirectory) return {supported: false, profileUtf16: null};
    try {
        const root = await navigator.storage.getDirectory();
        try {
            const file = await root.getFileHandle(VBA_FONT_PREFERENCE_FILE);
            const bytes = new Uint8Array(await (await file.getFile()).arrayBuffer());
            return {supported: true, profileUtf16: decodeVbaFontPreference(bytes) ?? null};
        } catch (error) {
            if (error.name === 'NotFoundError') return {supported: true, profileUtf16: null};
            throw error;
        }
    } catch {
        return {supported: false, profileUtf16: null};
    }
}

export async function openVbaStandardFontPreferenceStore() {
    const root = await navigator.storage.getDirectory();
    const file = await root.getFileHandle(VBA_FONT_PREFERENCE_FILE, {create: true});
    let access;
    try { access = await file.createSyncAccessHandle({mode: 'readwrite-unsafe'}); }
    catch (error) {
        if (error.name !== 'TypeError') throw error;
        access = await file.createSyncAccessHandle();
    }
    return {
        read() {
            if (access.getSize() === 0) return null;
            for (let attempt = 0; attempt < 3; attempt++) {
                const bytes = new Uint8Array(VBA_FONT_PREFERENCE_BYTES);
                const read = access.read(bytes, {at: 0});
                const value = read === bytes.length ? decodeVbaFontPreference(bytes) : undefined;
                if (value !== undefined) return value;
            }
            return null;
        },
        write(profile) {
            const bytes = encodeVbaFontPreference(profile);
            if (access.write(bytes, {at: 0}) !== bytes.length) throw new Error('The standard-font preference could not be saved.');
            access.flush();
        },
        close() { access.close(); },
    };
}
