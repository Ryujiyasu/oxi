// Copies the browser build into media/ and teaches the page to take a
// document from VS Code instead of a file picker, and to hand the edited
// file back when VS Code saves.
//
// The viewer IS the browser demo — same HTML, same WebAssembly module — so the
// extension never carries a second renderer that could drift from the engine.

const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..', '..', '..');
const web = path.join(root, 'web');
const media = path.resolve(__dirname, '..', 'media');

// Files the page needs besides the modules it imports: the engine's binary
// and the data a module fetches by name.
const EXTRA = ['oxidocs_wasm_bg.wasm', 'row-model.js', 'row-heights.json'];

/**
 * Every local module the page loads, followed through their own imports and
 * workers, so a module added to web/ comes along without this list being kept
 * by hand (a forgotten one leaves the editor blank: the page's module fails
 * to load and never says it is ready).
 */
function modulesOf(source, seen = new Set()) {
  const refs = /(?:from\s*|import\s*\(\s*|new URL\(\s*)'\.\/([^']+\.js)'/g;
  for (const [, name] of source.matchAll(refs)) {
    if (seen.has(name)) continue;
    seen.add(name);
    const file = path.join(web, name);
    if (fs.existsSync(file)) modulesOf(fs.readFileSync(file, 'utf8'), seen);
  }
  return seen;
}

const BRIDGE = `

// === VS Code bridge (inserted by editors/vscode/scripts/sync-web.js) ===
const __oxiVscode = typeof acquireVsCodeApi === 'function' ? acquireVsCodeApi() : null;
if (__oxiVscode) {
    document.body.classList.add('oxi-vscode-host');
    window.addEventListener('message', async (event) => {
        const msg = event.data;
        if (!msg || msg.type !== 'oxi:open') return;
        try {
            const binary = atob(msg.data);
            const bytes = new Uint8Array(binary.length);
            for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
            const sampleButtons = document.getElementById('sampleButtons');
            if (sampleButtons) sampleButtons.style.display = 'none';
            await handleFile(new File([bytes], msg.name));
        } catch (err) {
            __oxiVscode.postMessage({ type: 'oxi:error', message: String((err && err.message) || err) });
        }
    });
    // Saving: VS Code asks for the edited file and writes it itself. The page
    // says when there is something to save, so the tab shows it as changed.
    window.addEventListener('message', async (event) => {
        const msg = event.data;
        if (!msg || msg.type !== 'oxi:save') return;
        try {
            const edited = await buildEditedFile();
            __oxiVscode.postMessage(edited
                ? { type: 'oxi:saved', id: msg.id, data: __oxiBase64(edited.bytes), warning: edited.message || null }
                : { type: 'oxi:saved', id: msg.id, unchanged: true });
        } catch (err) {
            __oxiVscode.postMessage({ type: 'oxi:saved', id: msg.id, error: String((err && err.message) || err) });
        }
    });
    function __oxiBase64(bytes) {
        let binary = '';
        for (let i = 0; i < bytes.length; i += 0x8000) binary += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
        return btoa(binary);
    }
    let __oxiReported = false;
    const __oxiCheck = () => {
        const dirty = hasUnsavedEdits();
        if (dirty && !__oxiReported) __oxiVscode.postMessage({ type: 'oxi:edit' });
        __oxiReported = dirty;
    };
    // Typing reports at once; edits made some other way (VBA, sheet tools)
    // are caught by the look every second.
    document.addEventListener('input', () => setTimeout(__oxiCheck, 0), true);
    setInterval(__oxiCheck, 1000);
    window.addEventListener('message', (event) => {
        if (event.data && event.data.type === 'oxi:open') __oxiReported = false;
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

function patchHtml(html) {
  // Google Fonts is a cross-origin request the webview's policy refuses; the
  // page falls back to the system stack, which is what VS Code uses anyway.
  html = html.replace(
    /\s*<link[^>]*(fonts\.googleapis\.com|fonts\.gstatic\.com)[^>]*>/g,
    ''
  );

  const moduleStart = html.indexOf('<script type="module">');
  if (moduleStart === -1) throw new Error('web/index.html has no module script to extend');
  const moduleEnd = html.indexOf('</script>', moduleStart);
  if (moduleEnd === -1) throw new Error('web/index.html module script is unterminated');

  return html.slice(0, moduleEnd) + BRIDGE + html.slice(moduleEnd);
}

function main() {
  fs.mkdirSync(media, { recursive: true });

  const html = fs.readFileSync(path.join(web, 'index.html'), 'utf8');
  fs.writeFileSync(path.join(media, 'index.html'), patchHtml(html));
  console.log('index.html  <- web/index.html (+ VS Code bridge)');

  const assets = new Set([...modulesOf(html), ...EXTRA]);
  for (const asset of assets) {
    const from = path.join(web, asset);
    if (!fs.existsSync(from)) {
      throw new Error(`web/${asset} is missing: build the engine into web/ first (see README)`);
    }
    fs.copyFileSync(from, path.join(media, asset));
    const kb = (fs.statSync(from).size / 1024).toFixed(0);
    console.log(`${asset}  <- web/${asset} (${kb} kB)`);
  }
}

main();
