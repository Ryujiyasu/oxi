// Copies the browser build into media/ and teaches the pages to take their
// document from VS Code instead of a file picker, and to save through it.
//
// The editor IS the browser demo — same HTML, same WebAssembly module — so the
// extension never carries a second renderer that could drift from the engine.

const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..', '..', '..');
const web = path.join(root, 'web');
const media = path.resolve(__dirname, '..', 'media');

// Files the page needs besides the ones it names itself: the engine binary
// (loaded by oxidocs_wasm.js through import.meta.url) and the row-height table.
const EXTRA_ASSETS = ['oxidocs_wasm_bg.wasm', 'row-heights.json'];

// Every './name.js' the page imports, followed through the modules it reaches,
// so a new module in web/ cannot be left out of the extension unnoticed. The
// Word page also imports the site's modules as '../docs/name.js'; those are
// copied to media/docs/ and the page is pointed there, since the webview may
// read nothing outside media/.
const LOCAL_REF = /['"](\.\/|\.\.\/docs\/)([A-Za-z0-9_.-]+\.(?:js|mjs|json))['"]/g;
const siteDocs = path.join(root, 'docs');

/** Files the page needs, as paths relative to web/ ('docs/x.js' for the site's). */
function localAssets(html) {
  const found = new Set();
  const queue = [{ text: html, dir: '' }];
  while (queue.length) {
    const { text, dir } = queue.pop();
    for (const m of text.matchAll(LOCAL_REF)) {
      const name = (m[1] === './' ? dir : 'docs/') + m[2];
      if (found.has(name)) continue;
      found.add(name);
      const file = sourceOf(name);
      // oxidocs_wasm.js is wasm-bindgen output: its './oxidocs_wasm_bg.js' is
      // the WebAssembly import namespace, not a file.
      if (name.endsWith('.js') && name !== 'oxidocs_wasm.js' && fs.existsSync(file)) {
        queue.push({ text: fs.readFileSync(file, 'utf8'), dir: name.startsWith('docs/') ? 'docs/' : '' });
      }
    }
  }
  return [...found, ...EXTRA_ASSETS];
}

function sourceOf(name) {
  return name.startsWith('docs/') ? path.join(siteDocs, name.slice(5)) : path.join(web, name);
}

const BRIDGE = `

// === VS Code bridge (inserted by editors/vscode/scripts/sync-web.js) ===
const __oxiVscode = typeof acquireVsCodeApi === 'function' ? acquireVsCodeApi() : null;
if (__oxiVscode) {
    document.body.classList.add('oxi-vscode-host');
    const __oxiToBase64 = (bytes) => {
        let s = '';
        for (let i = 0; i < bytes.length; i += 0x8000) {
            s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
        }
        return btoa(s);
    };

    // Saving: VS Code asks for the file, and the page builds it exactly as its
    // Download button does (downloadEdited applies the edits to the bytes it
    // was opened with). The blob that download would hand to the browser is
    // caught here and sent to VS Code, which writes the file.
    let __oxiCapture = null;
    const __oxiBlobs = new Map();
    const __oxiCreateUrl = URL.createObjectURL;
    URL.createObjectURL = function (obj) {
        const url = __oxiCreateUrl.call(URL, obj);
        if (__oxiCapture && obj instanceof Blob) __oxiBlobs.set(url, obj);
        return url;
    };
    const __oxiAnchorClick = HTMLAnchorElement.prototype.click;
    HTMLAnchorElement.prototype.click = function () {
        const blob = __oxiCapture && this.download ? __oxiBlobs.get(this.href) : null;
        if (blob) {
            __oxiBlobs.delete(this.href);
            __oxiCapture(blob);
            return;
        }
        return __oxiAnchorClick.call(this);
    };
    const __oxiSave = async (id) => {
        let blob = null;
        __oxiCapture = (b) => { blob = b; };
        try {
            await downloadEdited();
        } catch (err) {
            __oxiVscode.postMessage({ type: 'oxi:saved', id, error: String((err && err.message) || err) });
            return;
        } finally {
            __oxiCapture = null;
            __oxiBlobs.clear();
        }
        if (blob) {
            const bytes = new Uint8Array(await blob.arrayBuffer());
            __oxiVscode.postMessage({ type: 'oxi:saved', id, data: __oxiToBase64(bytes) });
            return;
        }
        // No file was built: either there were no edits (VS Code then writes
        // the bytes it gave the page) or the page reported an error.
        const status = document.getElementById('status');
        const failed = status && status.className === 'error';
        __oxiVscode.postMessage({ type: 'oxi:saved', id, error: failed ? status.textContent : null });
    };

    // Any change typed into the document marks the VS Code tab as modified.
    document.addEventListener('input', (event) => {
        const target = event.target;
        if (target && target.closest && target.closest('#pages-container')) {
            __oxiVscode.postMessage({ type: 'oxi:edit' });
        }
    }, true);

    // The page's own save button saves through VS Code.
    const __oxiQatSave = document.getElementById('qatSave');
    if (__oxiQatSave) {
        __oxiQatSave.addEventListener('click', (event) => {
            event.preventDefault();
            event.stopPropagation();
            __oxiVscode.postMessage({ type: 'oxi:requestSave' });
        }, true);
    }

    window.addEventListener('message', async (event) => {
        const msg = event.data;
        if (!msg) return;
        if (msg.type === 'oxi:save') {
            await __oxiSave(msg.id);
            return;
        }
        if (msg.type !== 'oxi:open') return;
        try {
            const binary = atob(msg.data);
            const bytes = new Uint8Array(binary.length);
            for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
            await handleFile(new File([bytes], msg.name));
        } catch (err) {
            __oxiVscode.postMessage({ type: 'oxi:error', message: String((err && err.message) || err) });
        }
    });
    // The engine is loaded at module evaluation; wait for it before asking for
    // bytes, so the first document does not race the WebAssembly instantiation.
    const __oxiAnnounce = () => __oxiVscode.postMessage({ type: 'oxi:ready' });
    if (typeof wasmReady !== 'undefined' && wasmReady) {
        __oxiAnnounce();
    } else {
        const started = Date.now();
        const poll = setInterval(() => {
            if (typeof wasmReady !== 'undefined' && wasmReady) {
                clearInterval(poll);
                __oxiAnnounce();
            } else if (Date.now() - started > 60000) {
                clearInterval(poll);
                __oxiVscode.postMessage({ type: 'oxi:error', message: 'the engine did not finish loading' });
            }
        }, 100);
    }
}
`;

// The Word page (web/docs.html) edits through the document model, so it has
// a bridge of its own: the open file's bytes are always `currentDocx`, every
// edit is announced as 'oxi:document-edited', and the machine's fonts come
// from the extension, which can read them where the webview cannot.
const DOCS_BRIDGE = `

// === VS Code bridge (inserted by editors/vscode/scripts/sync-web.js) ===
const __oxiVscode = typeof acquireVsCodeApi === 'function' ? acquireVsCodeApi() : null;
if (__oxiVscode) {
    document.body.classList.add('oxi-vscode-host');
    const __oxiToBase64 = (bytes) => {
        let s = '';
        for (let i = 0; i < bytes.length; i += 0x8000) {
            s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
        }
        return btoa(s);
    };
    const __oxiFromBase64 = (text) => {
        const binary = atob(text);
        const bytes = new Uint8Array(binary.length);
        for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
        return bytes;
    };

    // Fonts: the extension reads the machine's own font programs; a face the
    // machine lacks is drawn in the face Oxi ships for it.
    const __oxiFontWaiters = new Map();
    let __oxiNextFont = 1;
    fontLoader.standIns = true;
    fontLoader.provider = (family, bold, italic) => new Promise((resolve) => {
        const id = __oxiNextFont++;
        __oxiFontWaiters.set(id, resolve);
        __oxiVscode.postMessage({ type: 'oxi:font', id, family, bold, italic });
    });

    // Undo: each step the page begins becomes an edit on VS Code's own stack,
    // whose undo and redo (Ctrl+Z, Ctrl+Y, the Edit menu, the page's buttons)
    // come back here as 'oxi:undo' and 'oxi:redo'. VS Code then also knows
    // when undoing returns the file to what was saved.
    hostOwnsUndo = true;
    document.addEventListener('oxi:document-edited', (event) => {
        __oxiVscode.postMessage({ type: 'oxi:edit', step: !!(event.detail && event.detail.step) });
    });
    document.addEventListener('oxi:request-undo', () => __oxiVscode.postMessage({ type: 'oxi:requestUndo' }));
    document.addEventListener('oxi:request-redo', () => __oxiVscode.postMessage({ type: 'oxi:requestRedo' }));

    // The page's own save (its button, the File menu, Ctrl+S) saves through
    // VS Code, which owns the file.
    const __oxiSaveButton = document.getElementById('btnSave');
    if (__oxiSaveButton) {
        __oxiSaveButton.addEventListener('click', (event) => {
            event.preventDefault();
            event.stopImmediatePropagation();
            __oxiVscode.postMessage({ type: 'oxi:requestSave' });
        }, true);
    }
    // A file dropped on the page would replace what is shown, not the file
    // VS Code has open.
    window.addEventListener('drop', (event) => { event.preventDefault(); event.stopPropagation(); }, true);
    window.addEventListener('dragover', (event) => { event.preventDefault(); event.stopPropagation(); }, true);

    window.addEventListener('message', async (event) => {
        const msg = event.data;
        if (!msg) return;
        if (msg.type === 'oxi:font') {
            const resolve = __oxiFontWaiters.get(msg.id);
            __oxiFontWaiters.delete(msg.id);
            if (resolve) resolve(msg.data ? { bytes: __oxiFromBase64(msg.data), index: msg.index || 0 } : null);
            return;
        }
        if (msg.type === 'oxi:undo') { undoEdit(); return; }
        if (msg.type === 'oxi:redo') { redoEdit(); return; }
        if (msg.type === 'oxi:save') {
            // Typing after a save is a new step, so undoing it comes back to
            // exactly what was saved.
            breakUndoGroup();
            if (!currentDocx) {
                __oxiVscode.postMessage({ type: 'oxi:saved', id: msg.id, error: 'no document is open' });
            } else {
                __oxiVscode.postMessage({ type: 'oxi:saved', id: msg.id, data: __oxiToBase64(currentDocx) });
            }
            return;
        }
        if (msg.type !== 'oxi:open') return;
        try {
            await openFile(new File([__oxiFromBase64(msg.data)], msg.name));
        } catch (err) {
            __oxiVscode.postMessage({ type: 'oxi:error', message: String((err && err.message) || err) });
        }
    });
    // The module awaited the engine before it got here.
    __oxiVscode.postMessage({ type: 'oxi:ready' });
}
`;

function patchHtml(html, bridge, page) {
  // Google Fonts is a cross-origin request the webview's policy refuses; the
  // page falls back to the system stack, which is what VS Code uses anyway.
  html = html.replace(
    /\s*<link[^>]*(fonts\.googleapis\.com|fonts\.gstatic\.com)[^>]*>/g,
    ''
  );

  html = html.replace(/(['"])\.\.\/docs\//g, '$1./docs/');

  const moduleStart = html.indexOf('<script type="module">');
  if (moduleStart === -1) throw new Error(`web/${page} has no module script to extend`);
  const moduleEnd = html.indexOf('</script>', moduleStart);
  if (moduleEnd === -1) throw new Error(`web/${page} module script is unterminated`);

  return html.slice(0, moduleEnd) + bridge + html.slice(moduleEnd);
}

// index.html opens .xlsx and .pptx; docs.html is the Word editor.
const PAGES = [
  { page: 'index.html', bridge: BRIDGE },
  { page: 'docs.html', bridge: DOCS_BRIDGE },
];

function main() {
  fs.mkdirSync(media, { recursive: true });

  const assets = new Set();
  for (const { page, bridge } of PAGES) {
    const html = fs.readFileSync(path.join(web, page), 'utf8');
    fs.writeFileSync(path.join(media, page), patchHtml(html, bridge, page));
    console.log(`${page}  <- web/${page} (+ VS Code bridge)`);
    for (const asset of localAssets(html)) assets.add(asset);
  }

  const missing = [];
  for (const asset of assets) {
    const from = sourceOf(asset);
    if (!fs.existsSync(from)) {
      missing.push(asset);
      continue;
    }
    const to = path.join(media, asset);
    fs.mkdirSync(path.dirname(to), { recursive: true });
    fs.copyFileSync(from, to);
    const kb = (fs.statSync(from).size / 1024).toFixed(0);
    console.log(`${asset}  <- ${path.relative(root, from)} (${kb} kB)`);
  }
  if (missing.length) {
    throw new Error(`web/ is missing files the pages load: ${missing.join(', ')}` +
      ' (build the WebAssembly package and copy it into web/ first)');
  }
}

main();
