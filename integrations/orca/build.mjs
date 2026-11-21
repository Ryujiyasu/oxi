// Builds the Orca plugin "Oxi Slides": a read-only .pptx viewer made from the
// Oxi web editor page (web/index.html) and the slides-only engine.
//
//   cd crates/oxidocs-wasm && wasm-pack build --target web --release \
//       --out-dir pkg-slides -- --features slides,external-font-catalog
//   node integrations/orca/build.mjs [--pkg <wasm-pack out dir>] [--out <dir>]
//
// The output folder is the plugin: orca-plugin.json at its root. Check it with
// `orca plugin check <dir>` and try it with `orca plugin dev load <dir>`.
import fs from 'node:fs';
import path from 'node:path';
import zlib from 'node:zlib';
import { fileURLToPath } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(here, '..', '..');
const args = process.argv.slice(2);
const opt = (name, fallback) => {
  const i = args.indexOf(name);
  return path.resolve(i === -1 ? fallback : args[i + 1]);
};
const pkg = opt('--pkg', path.join(root, 'crates', 'oxidocs-wasm', 'pkg-slides'));
const out = opt('--out', path.join(here, 'dist'));
const web = path.join(root, 'web');
const editor = path.join(out, 'editor');

// Orca's isolated frames read WebAssembly, scripts and JSON whole, at most
// 16 MiB a file. The engine is larger, so it ships gzipped in parts.
const PART_BYTES = 15 * 1024 * 1024;
// A plugin folder holds at most 50 MiB.
const PLUGIN_MAX_BYTES = 50 * 1024 * 1024;

// Rewrites `html` with `from` → `to`, and fails when the page has changed so
// that `from` is no longer there (the build must not ship a half-patched page).
function patch(html, from, to, what) {
  const next = typeof from === 'string' ? html.replace(from, to) : html.replace(from, to);
  if (next === html) throw new Error(`web/index.html: ${what} not found`);
  return next;
}

// --- the page ---------------------------------------------------------------

const STYLE = `
<style>
/* Oxi Slides (Orca plugin): a viewer. No title bar or ribbon; the slide area
   takes the room, and the status bar carries the slide show button. */
.oxi-orca-host #titleBar, .oxi-orca-host #ribbon { display: none !important; }
.oxi-orca-host .oxi-orca-show {
    margin-left: 12px; padding: 1px 10px; border: 1px solid rgba(255,255,255,.55);
    border-radius: 3px; background: transparent; color: inherit; font: inherit; cursor: pointer;
}
.oxi-orca-host .oxi-orca-show:hover { background: rgba(255,255,255,.15); }
</style>
`;

const LOADER = (parts) => `
// === Oxi engine loader (inserted by integrations/orca/build.mjs) ===
async function __oxiEngineBytes() {
    const parts = await Promise.all(${JSON.stringify(parts)}.map(async (name) => {
        const r = await fetch('./' + name);
        if (!r.ok) throw new Error(name + ': HTTP ' + r.status);
        return new Uint8Array(await r.arrayBuffer());
    }));
    const gz = new Blob(parts);
    return await new Response(gz.stream().pipeThrough(new DecompressionStream('gzip'))).arrayBuffer();
}
`;

const BRIDGE = `
// === Orca bridge (inserted by integrations/orca/build.mjs) ===
if (window.orca && orca.document) {
    document.body.classList.add('oxi-vscode-host', 'oxi-orca-host');
    // A file dropped on the page would replace the deck the tab shows.
    window.addEventListener('drop', (e) => { e.preventDefault(); e.stopPropagation(); }, true);
    window.addEventListener('dragover', (e) => { e.preventDefault(); e.stopPropagation(); }, true);

    const until = async (test, ms = 60000) => {
        const t0 = Date.now();
        while (!test()) {
            if (Date.now() - t0 > ms) return false;
            await new Promise((r) => setTimeout(r, 50));
        }
        return true;
    };

    const show = document.createElement('button');
    show.className = 'oxi-orca-show';
    show.textContent = '\\u25B6 ' + ribbonText('Slide Show');
    show.addEventListener('click', () => startSlideShow(currentSlideIndex || 0));
    document.querySelector('#statusBar .status-left')?.appendChild(show);

    // The same deck again (an agent rewrote it): stay on the slide being read.
    let fileType = '.pptx';
    const open = async (bytes, reload) => {
        const place = reload ? (currentSlideIndex || 0) : 0;
        await handleFile(new File([new Uint8Array(bytes)], 'presentation' + fileType));
        if (place > 0 && await until(() => document.querySelectorAll('.slide-thumb').length, 5000)) {
            const thumbs = document.querySelectorAll('.slide-thumb');
            thumbs[Math.min(place, thumbs.length - 1)].click();
        }
    };

    // Ask Orca for the document at once (it expects the request within
    // seconds), and draw it when the engine has started. Read only: the
    // viewer never calls change(), so Orca never saves anything from it.
    const fail = (err) => {
        console.error('Oxi Slides:', err);
        showStatus(String((err && err.message) || err), 'error');
    };
    (async () => {
        const engine = until(() => typeof wasmReady !== 'undefined' && wasmReady);
        let doc;
        if ('session' in orca.document) {
            doc = await orca.document.session({
                decode: (content) => content.bytes,
                encode: (bytes) => ({ kind: 'bytes', bytes }),
            });
        } else {
            const first = await orca.document.get();
            if (first.fileType) fileType = first.fileType;
            const replaced = [];
            doc = { value: first.content.bytes, onReplace: (cb) => replaced.push(cb) };
            orca.document.onDidChange((d) => replaced.forEach((cb) => cb(d.content.bytes)));
        }
        doc.onReplace((bytes) => open(bytes, true).catch(fail));
        if (!await engine) throw new Error('The Oxi engine did not finish loading');
        await open(doc.value, false);
    })().catch(fail);
}
`;

function buildPage(parts) {
  let html = fs.readFileSync(path.join(web, 'index.html'), 'utf8');
  // No network in Orca's frame: Google Fonts would only be refused.
  html = html.replace(/\s*<link[^>]*(fonts\.googleapis\.com|fonts\.gstatic\.com)[^>]*>/g, '');
  // The slides-only engine lacks the Word, Excel and VBA exports; importing
  // the module as a namespace leaves those undefined instead of failing.
  html = patch(html, /import init, \{([^}]*)\} from '\.\/oxidocs_wasm\.js';/,
    (_, names) => `import init, * as __oxiEngine from './oxidocs_wasm.js';\nconst {${names}} = __oxiEngine;`,
    'the named import of ./oxidocs_wasm.js');
  html = patch(html, 'await init();', 'await init({ module_or_path: await __oxiEngineBytes() });',
    'the engine start (await init();)');
  // The slide being read is drawn like the thumbnails: not editable.
  html = patch(html, 'buildSlideContent(wrapper, slide, slideIdx, scale, true);',
    'buildSlideContent(wrapper, slide, slideIdx, scale, false);', 'the editable main slide');
  html = patch(html, '</head>', STYLE + '</head>', '</head>');
  const start = html.indexOf('<script type="module">');
  const end = html.indexOf('</script>', start);
  if (start === -1 || end === -1) throw new Error('web/index.html: module script not found');
  return html.slice(0, end) + LOADER(parts) + BRIDGE + html.slice(end);
}

// --- assemble ---------------------------------------------------------------

fs.rmSync(out, { recursive: true, force: true });
fs.mkdirSync(editor, { recursive: true });

const wasm = path.join(pkg, 'oxidocs_wasm_bg.wasm');
if (!fs.existsSync(wasm)) {
  throw new Error(`${wasm} is missing: build the slides engine first (see the top of this file)`);
}
const gz = zlib.gzipSync(fs.readFileSync(wasm), { level: 9 });
const parts = [];
for (let i = 0; i * PART_BYTES < gz.length; i++) {
  const name = `oxidocs_engine.gz.${i}.wasm`;
  fs.writeFileSync(path.join(editor, name), gz.subarray(i * PART_BYTES, (i + 1) * PART_BYTES));
  parts.push(name);
}
fs.copyFileSync(path.join(pkg, 'oxidocs_wasm.js'), path.join(editor, 'oxidocs_wasm.js'));

const html = buildPage(parts);
fs.writeFileSync(path.join(editor, 'index.html'), html);

// Local files the page and its scripts load (samples and demo pages are left out).
const LOCAL_REF = /['"]\.\/([A-Za-z0-9_.-]+\.(?:js|mjs|json))['"]/g;
const copied = new Set(['oxidocs_wasm.js']);
const queue = [html];
while (queue.length) {
  for (const m of queue.pop().matchAll(LOCAL_REF)) {
    const name = m[1];
    if (copied.has(name) || !fs.existsSync(path.join(web, name))) continue;
    copied.add(name);
    fs.copyFileSync(path.join(web, name), path.join(editor, name));
    if (/\.m?js$/.test(name)) queue.push(fs.readFileSync(path.join(web, name), 'utf8'));
  }
}
// Loaded by name at run time rather than as a './' literal.
for (const name of ['vba-worker.js', 'row-heights.json', 'favicon.ico']) {
  if (!copied.has(name) && fs.existsSync(path.join(web, name))) {
    fs.copyFileSync(path.join(web, name), path.join(editor, name));
  }
}
fs.copyFileSync(path.join(root, 'docs', 'icon-64.png'), path.join(editor, 'icon-64.png'));

for (const name of ['orca-plugin.json', 'README.md']) {
  fs.copyFileSync(path.join(here, 'plugin', name), path.join(out, name));
}
// Icon, cover and screenshots named in the manifest; the plugin directory
// shows them. screenshots/ holds the scripts that made them.
fs.cpSync(path.join(here, 'plugin', 'images'), path.join(out, 'images'), { recursive: true });
fs.copyFileSync(path.join(root, 'LICENSE'), path.join(out, 'LICENSE'));

let bytes = 0;
let files = 0;
const walk = (dir) => {
  for (const e of fs.readdirSync(dir, { withFileTypes: true })) {
    const p = path.join(dir, e.name);
    if (e.isDirectory()) walk(p);
    else { files += 1; bytes += fs.statSync(p).size; }
  }
};
walk(out);
if (bytes > PLUGIN_MAX_BYTES) throw new Error(`plugin is ${bytes} bytes, over Orca's 50 MiB`);
console.log(`${path.relative(process.cwd(), out) || '.'}: ${files} files, ${(bytes / 1048576).toFixed(1)} MiB`);
