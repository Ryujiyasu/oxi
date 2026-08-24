// Copies the browser build into media/ and teaches the page to accept a
// document from VS Code instead of a file picker.
//
// The viewer IS the browser demo — same HTML, same WebAssembly module — so the
// extension never carries a second renderer that could drift from the engine.

const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..', '..', '..');
const web = path.join(root, 'web');
const media = path.resolve(__dirname, '..', 'media');

const ASSETS = [
  'oxidocs_wasm.js',
  'oxidocs_wasm_bg.wasm',
  'vba-runner.js',
  'vba-worker.js',
  'row-model.js',
  'row-heights.json',
];

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

  for (const asset of ASSETS) {
    const from = path.join(web, asset);
    if (!fs.existsSync(from)) {
      console.log(`${asset}  -- not present in web/, skipped`);
      continue;
    }
    fs.copyFileSync(from, path.join(media, asset));
    const kb = (fs.statSync(from).size / 1024).toFixed(0);
    console.log(`${asset}  <- web/${asset} (${kb} kB)`);
  }
}

main();
