// Copies the browser build into media/ and teaches the page to take its
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
// so a new module in web/ cannot be left out of the extension unnoticed.
const LOCAL_REF = /['"]\.\/([A-Za-z0-9_.-]+\.(?:js|mjs|json))['"]/g;

function localAssets(html) {
  const found = new Set();
  const queue = [html];
  while (queue.length) {
    const text = queue.pop();
    for (const m of text.matchAll(LOCAL_REF)) {
      const name = m[1];
      if (found.has(name)) continue;
      found.add(name);
      const file = path.join(web, name);
      // oxidocs_wasm.js is wasm-bindgen output: its './oxidocs_wasm_bg.js' is
      // the WebAssembly import namespace, not a file.
      if (name.endsWith('.js') && name !== 'oxidocs_wasm.js' && fs.existsSync(file)) {
        queue.push(fs.readFileSync(file, 'utf8'));
      }
    }
  }
  return [...found, ...EXTRA_ASSETS];
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

  const missing = [];
  for (const asset of localAssets(html)) {
    const from = path.join(web, asset);
    if (!fs.existsSync(from)) {
      missing.push(asset);
      continue;
    }
    fs.copyFileSync(from, path.join(media, asset));
    const kb = (fs.statSync(from).size / 1024).toFixed(0);
    console.log(`${asset}  <- web/${asset} (${kb} kB)`);
  }
  if (missing.length) {
    throw new Error(`web/ is missing files the page loads: ${missing.join(', ')}` +
      ' (build the WebAssembly package and copy it into web/ first)');
  }
}

main();
