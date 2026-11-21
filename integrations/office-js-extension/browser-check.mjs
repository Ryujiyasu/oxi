// SPDX-License-Identifier: MPL-2.0
// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

// The extension in a real browser: headless Edge loads it, opens a stand-in
// editor page on localhost whose add-in pane holds an add-in page that asks
// for Microsoft's office.js. The extension must answer that request with
// oxi-office.js (no request leaves the machine: the rule redirects it before
// it is sent), and the add-in's Excel.run must reach the editor through it.
//
//   node integrations/office-js-extension/build.mjs
//   node integrations/office-js-extension/browser-check.mjs [path to msedge.exe]
import { spawn, spawnSync } from 'node:child_process';
import { createServer } from 'node:http';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
// OXI_CHECK_NATIVE=<path to an .xll>: also call that XLL (and, with
// OXI_CHECK_COM=<ProgID>, a COM add-in) from the page through the extension
// and the native-messaging hosts (install.ps1 must have registered them for
// this extension's id).
const nativeXll = process.env.OXI_CHECK_NATIVE ?? '';
const nativeCom = process.env.OXI_CHECK_COM ?? '';
const edge = process.argv[2] ?? 'C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe';

// The stand-in editor: answers oxi-office.js's snapshot and sync requests
// with a one-sheet workbook it keeps in memory, and records what it was sent.
const editor = `<!doctype html><meta charset="utf-8"><title>editor</title>
<iframe id="addinFrame"></iframe>
<script>
// The add-in opens once the extension has said it is there (its redirect
// rule is in place by then), as the editor opens one on the user's request.
const opening = setInterval(() => {
  if (document.documentElement.dataset.oxiOfficeExtension === 'on') {
    clearInterval(opening);
    document.getElementById('addinFrame').src = '/addin.html';
  }
}, 50);
const cells = new Map([['0,0', { value: 'hello', text: 'hello' }]]);
window.received = [];
// OxiXll / OxiCom as the editor asks them: a request posted to the page,
// relayed by the content script.
window.native = 'skipped';
const asked = new Map();
let seq = 0;
window.addEventListener('message', event => {
  const data = event.data;
  if (event.source !== window || !data || data.oxiXll !== 'reply') return;
  asked.get(data.id)?.(data);
});
const ask = request => new Promise(resolve => { const id = ++seq; asked.set(id, resolve); window.postMessage({ oxiXll: 'request', id, ...request }, '*'); });
const xll = ${JSON.stringify(nativeXll)};
const com = ${JSON.stringify(nativeCom)};
if (xll) {
  const waitFlag = setInterval(async () => {
    if (document.documentElement.dataset.oxiOfficeExtension !== 'on') return;
    clearInterval(waitFlag);
    const out = {};
    out.list = (await ask({ xll, op: 'list' })).functions?.length ?? 'none';
    out.add = (await ask({ xll, op: 'call', function: 'OxiAdd', args: [2, 3] })).value;
    out.range = (await ask({ xll, op: 'call', function: 'OxiRangeSum', args: [[[1, 2], [3, 4]]] })).value;
    if (com) out.com = (await ask({ kind: 'com', progId: com, function: 'Add', args: [4, 5] })).value;
    window.native = JSON.stringify(out);
  }, 50);
}
const snapshot = () => ({ names: ['Sheet1'], active: 0, selection: { sheet: 0, row: 0, col: 0, rows: 1, cols: 1 },
  sheets: [{ cells: [...cells].map(([k, v]) => [...k.split(',').map(Number), v]), used: { row: 0, col: 0, rows: 2, cols: 1 } }],
  definedNames: [], tables: [] });
window.addEventListener('message', event => {
  const data = event.data;
  if (!data || data.oxiOffice !== 'request') return;
  const reply = extra => event.source.postMessage({ oxiOffice: 'reply', id: data.id, ...extra }, '*');
  if (data.op === 'snapshot') return reply({ snapshot: snapshot() });
  if (data.op !== 'sync') return;
  for (const item of data.batch || []) {
    window.received.push(item);
    if (item.op === 'write') for (const e of item.edits) cells.set(e.row + ',' + e.col, { value: e.value, text: String(e.value) });
  }
  reply({ snapshot: snapshot() });
});
</script>`;

// The add-in: what any Office add-in page does.
const addin = `<!doctype html><meta charset="utf-8"><title>addin</title>
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
<script>
window.result = 'waiting';
(async () => {
  try {
    await Office.onReady();
    await Excel.run(async ctx => {
      const sheet = ctx.workbook.worksheets.getActiveWorksheet();
      const a1 = sheet.getRange('A1').load('values');
      sheet.getRange('A2').values = [[42]];
      await ctx.sync();
      const a2 = sheet.getRange('A2').load('values');
      await ctx.sync();
      window.result = JSON.stringify({ a1: a1.values, a2: a2.values, oxi: typeof OxiOffice });
    });
  } catch (e) { window.result = 'error: ' + e.message; }
})();
</script>`;

const server = createServer((req, res) => {
  const body = req.url.startsWith('/addin') ? addin : editor;
  res.writeHead(200, { 'content-type': 'text/html; charset=utf-8' });
  res.end(body);
});
await new Promise(r => server.listen(0, '127.0.0.1', r));
const port = server.address().port;
const profile = mkdtempSync(join(tmpdir(), 'oxi-ext-'));
const debugPort = 9300 + Math.floor(Math.random() * 500);
const browser = spawn(edge, [
  '--headless=new', `--remote-debugging-port=${debugPort}`, `--user-data-dir=${profile}`,
  `--load-extension=${here}`, `--disable-extensions-except=${here}`, '--no-first-run',
  `http://localhost:${port}/editor.html`,
], { stdio: 'ignore' });

const sleep = ms => new Promise(r => setTimeout(r, ms));
async function targets() {
  for (let i = 0; i < 50; i++) {
    try { return await (await fetch(`http://127.0.0.1:${debugPort}/json`)).json(); } catch { await sleep(200); }
  }
  throw new Error('the browser did not open its debugging port');
}
function evaluate(wsUrl, expression) {
  return new Promise((resolve, reject) => {
    const ws = new WebSocket(wsUrl);
    ws.onopen = () => ws.send(JSON.stringify({ id: 1, method: 'Runtime.evaluate', params: { expression, returnByValue: true, awaitPromise: true } }));
    ws.onmessage = m => { const d = JSON.parse(m.data); if (d.id === 1) { ws.close(); resolve(d.result?.result?.value); } };
    ws.onerror = reject;
  });
}

let outcome = null;
let failed = 'timed out';
try {
  for (let i = 0; i < 60 && !outcome; i++) {
    await sleep(500);
    const list = await targets();
    const frame = list.find(t => t.url.includes('/addin.html'));
    const page = list.find(t => t.url.includes('/editor.html') && t.type === 'page');
    if (!frame && page) {
      // The add-in frame shares the page's process: read it through the page.
      const value = await evaluate(page.webSocketDebuggerUrl,
        `(() => { try { return document.getElementById('addinFrame').contentWindow.result; } catch (e) { return 'cross-origin'; } })()`);
      if (value && value !== 'waiting' && value !== 'cross-origin') outcome = { result: value, page };
    } else if (frame) {
      const value = await evaluate(frame.webSocketDebuggerUrl, 'window.result');
      if (value && value !== 'waiting') outcome = { result: value, page };
    }
  }
  if (!outcome && process.env.OXI_DEBUG) {
    const list = await targets();
    for (const t of list) console.log('  target', t.type, t.url);
    const page = list.find(t => t.url.includes('/editor.html'));
    if (page) console.log('  fetch addin', await evaluate(page.webSocketDebuggerUrl, `fetch('/addin.html').then(r => r.status + ' ' + r.headers.get('content-type'), e => 'failed ' + e.message)`));
    if (page) console.log('  page flag', await evaluate(page.webSocketDebuggerUrl, `document.documentElement.dataset.oxiOfficeExtension + ' src=' + document.getElementById('addinFrame').src`));
    if (page) console.log('  frame says', await evaluate(page.webSocketDebuggerUrl, `(() => { const w = document.getElementById('addinFrame').contentWindow; try { return w.location.href + ' ' + w.document.readyState + ' ' + String(w.result) + ' office=' + typeof w.Office + ' scripts=' + [...w.document.scripts].map(s => s.src).join(','); } catch (e) { return 'cannot read: ' + e.message; } })()`));
  }
  if (outcome) {
    const received = await evaluate(outcome.page.webSocketDebuggerUrl, 'JSON.stringify(window.received)');
    const flag = await evaluate(outcome.page.webSocketDebuggerUrl, 'document.documentElement.dataset.oxiOfficeExtension');
    let native = 'skipped';
    for (let i = 0; i < 40 && nativeXll; i++) {
      native = await evaluate(outcome.page.webSocketDebuggerUrl, 'window.native');
      if (native !== 'skipped') break;
      await sleep(250);
    }
    console.log(`  native messaging: ${native}`);
    const nativeGood = !nativeXll || (() => {
      const n = JSON.parse(native);
      return n.list > 0 && n.add === 5 && n.range === 10 && (!nativeCom || n.com === 9);
    })();
    const parsed = outcome.result.startsWith('{') ? JSON.parse(outcome.result) : null;
    const good = parsed && parsed.oxi === 'object' && parsed.a1[0][0] === 'hello' && parsed.a2[0][0] === 42 && flag === 'on'
      && JSON.parse(received).some(item => item.op === 'write') && nativeGood;
    console.log(`  add-in saw: ${outcome.result}`);
    console.log(`  editor received: ${received}`);
    console.log(`  page flag: ${flag}`);
    failed = good ? null : 'unexpected answers';
  }
} finally {
  // The browser's helpers outlive its first process: end the whole tree.
  if (process.platform === 'win32') spawnSync('taskkill', ['/PID', String(browser.pid), '/T', '/F']);
  else browser.kill();
  server.close();
  await sleep(500);
  try { rmSync(profile, { recursive: true, force: true }); } catch { /* the browser may still hold it */ }
}
console.log(failed ? `FAIL - ${failed}` : 'ok - office.js was answered by oxi-office.js in a real browser and reached the editor');
process.exit(failed ? 1 : 0);
