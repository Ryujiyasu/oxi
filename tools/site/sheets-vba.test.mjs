// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

// Usage: node tools/site/sheets-vba.test.mjs <wasm-package> <empty-output-dir>
// Exercises the actual page and module Worker in an isolated Edge instance.
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import http from 'node:http';
import {spawn} from 'node:child_process';
import {once} from 'node:events';
import {fileURLToPath} from 'node:url';
import {createHash} from 'node:crypto';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..');
const pkg = path.resolve(process.argv[2]);
const output = path.resolve(process.argv[3]);
assert(!fs.existsSync(output), 'Use a fresh output directory');
fs.mkdirSync(output, {recursive: true});
const profile = path.join(output, 'owned-edge-profile');
const downloads = path.join(output, 'downloads');
fs.mkdirSync(profile); fs.mkdirSync(downloads);
const result = {passed: false, checks: [], exceptions: [], requests: [], package: pkg, files: {}};
let browser, browserCall, browserSocket, pageSocket, exit, exited = false;
let server;
const hash = data => createHash('sha256').update(data).digest('hex');
const delay = ms => new Promise(resolve => setTimeout(resolve, ms));
function connect(socket) {
  let next = 0;
  const pending = new Map();
  socket.addEventListener('message', event => {
    const message = JSON.parse(event.data);
    if (message.method === 'Runtime.exceptionThrown') result.exceptions.push(message.params);
    const task = pending.get(message.id);
    if (!task) return;
    pending.delete(message.id);
    message.error ? task.reject(Error(JSON.stringify(message.error))) : task.resolve(message.result);
  });
  socket.addEventListener('close', () => {
    for (const task of pending.values()) task.reject(Error('CDP closed'));
    pending.clear();
  });
  return (method, params = {}) => new Promise((resolve, reject) => {
    const id = ++next; pending.set(id, {resolve, reject});
    socket.send(JSON.stringify({id, method, params}));
  });
}
try {
  const helpers = ['vba-runner.js', 'vba-worker.js', 'vba-editor-settings.js', 'vba-file-downloads.js'];
  const workflow = fs.readFileSync(path.join(root, '.github/workflows/site.yml'), 'utf8');
  for (const name of helpers) assert(workflow.includes(`web/${name}`), `Deployment must include ${name}`);
  server = http.createServer((request, response) => {
    try {
      const name = decodeURIComponent(new URL(request.url, 'http://localhost').pathname).slice(1) || 'sheets.html';
      assert(!name.split('/').includes('..'));
      const file = name.startsWith('oxidocs_wasm') ? path.join(pkg, name)
        : path.join(root, helpers.includes(name) ? 'web' : 'docs', name);
      if (!fs.existsSync(file) || !fs.statSync(file).isFile()) { response.writeHead(404); response.end(); return; }
      const bytes = fs.readFileSync(file);
      result.files[file] = hash(bytes);
      result.requests.push(name);
      response.setHeader('Content-Type', name.endsWith('.wasm') ? 'application/wasm'
        : /\.(m?js)$/.test(name) ? 'text/javascript;charset=utf-8'
        : name.endsWith('.html') ? 'text/html;charset=utf-8' : 'application/octet-stream');
      response.end(bytes);
    } catch (error) { response.writeHead(500); response.end(String(error)); }
  });
  server.listen(0, '127.0.0.1'); await once(server, 'listening');
  const executable = process.env.EDGE_PATH || path.join(process.env['ProgramFiles(x86)'], 'Microsoft/Edge/Application/msedge.exe');
  browser = spawn(executable, ['--headless=new', '--remote-debugging-port=0', `--user-data-dir=${profile}`,
    '--no-first-run', '--no-default-browser-check', '--disable-extensions', '--disable-background-networking',
    '--disable-sync', '--disable-component-update', 'about:blank'], {windowsHide: true, stdio: 'ignore'});
  result.browser_pid = browser.pid;
  result.browser_started_at = Date.now();
  exit = new Promise(resolve => browser.once('exit', code => { exited = true; result.browser_exit = code; resolve(); }));
  const waitUntil = async predicate => {
    const deadline = Date.now() + 120000;
    while (!await predicate()) { assert(!exited, 'Browser exited'); assert(Date.now() < deadline, 'Timed out'); await delay(100); }
  };
  const portFile = path.join(profile, 'DevToolsActivePort');
  await waitUntil(() => fs.existsSync(portFile));
  const endpoint = `http://127.0.0.1:${Number(fs.readFileSync(portFile, 'utf8').split('\n')[0])}`;
  const version = await (await fetch(`${endpoint}/json/version`)).json();
  result.browser_version = version.Browser;
  browserSocket = new WebSocket(version.webSocketDebuggerUrl); await once(browserSocket, 'open');
  browserCall = connect(browserSocket);
  const target = await browserCall('Target.createTarget', {url: 'about:blank'});
  const page = (await (await fetch(`${endpoint}/json/list`)).json()).find(item => item.id === target.targetId);
  pageSocket = new WebSocket(page.webSocketDebuggerUrl); await once(pageSocket, 'open');
  const call = connect(pageSocket);
  await call('Runtime.enable'); await call('Page.enable');
  await call('Emulation.setDeviceMetricsOverride', {width: 1280, height: 900, deviceScaleFactor: 1, mobile: false});
  await browserCall('Browser.setDownloadBehavior', {behavior: 'allow', downloadPath: downloads});
  const evaluate = async expression => {
    const answer = await call('Runtime.evaluate', {expression, returnByValue: true, awaitPromise: true});
    assert(!answer.exceptionDetails, JSON.stringify(answer.exceptionDetails));
    return answer.result.value;
  };
  await call('Page.navigate', {url: `http://127.0.0.1:${server.address().port}/sheets.html`});
  await waitUntil(() => evaluate(`document.getElementById('btnTryVba')?.disabled === false`));
  assert(!result.requests.includes('vba-worker.js'), 'Opening the page must not execute macros');
  await evaluate(`document.getElementById('btnTryVba').click()`);
  await waitUntil(() => evaluate(`document.getElementById('vbaDialog').open`));
  assert(!result.requests.includes('vba-worker.js'), 'Opening the editor must not execute macros');
  result.checks.push('sample and editor load without executing macros');
  const run = async (source, procedure = 'Demo', args = '[]') => {
    await evaluate(`document.getElementById('vbaSource').value=${JSON.stringify(source)};
      document.getElementById('vbaProcedure').value=${JSON.stringify(procedure)};
      document.getElementById('vbaArguments').value=${JSON.stringify(args)};
      document.getElementById('vbaRun').click();`);
    await waitUntil(() => evaluate(`!document.getElementById('vbaRun').disabled`));
    return evaluate(`document.getElementById('vbaOutput').textContent`);
  };
  const text = await run('Sub Demo()\nRange("A1").Value = "UI VBA saved"\nRange("B3").Value = 42\nDebug.Print Range("A1").Value\nEnd Sub');
  assert(text.includes('UI VBA saved'), text);
  assert(await evaluate(`!document.getElementById('btnSave').disabled`));
  assert(result.requests.includes('vba-worker.js'));
  result.checks.push('actual Worker executes VBA, prints output, enables saving');
  // Download through the real UI, then parse the actual saved ZIP in the browser.
  await evaluate(`document.getElementById('vbaDialog').close(); document.getElementById('btnSave').click()`);
  await waitUntil(() => fs.existsSync(path.join(downloads, 'sample.xlsx')));
  const saved = fs.readFileSync(path.join(downloads, 'sample.xlsx'));
  result.saved_file_sha256 = hash(saved);
  const cellValues = await evaluate(`(async()=>{const w=await import('./oxidocs_wasm.js');
    const book=w.parse_spreadsheet(new Uint8Array(${JSON.stringify([...saved])}));
    return book.sheets[0].rows.flatMap(r=>r.cells.map(c=>({row:r.index,col:c.col,value:c.value})));})()`);
  fs.writeFileSync(path.join(output, 'saved-cells.json'), JSON.stringify(cellValues, null, 2));
  assert.deepEqual(cellValues.find(c => c.row === 1 && c.col === 0)?.value, {String: 'UI VBA saved'});
  assert.deepEqual(cellValues.find(c => c.row === 3 && c.col === 1)?.value, {Number: 42});
  result.checks.push('downloaded workbook contains macro edits');
  await evaluate(`document.getElementById('cellRef').value='B3';
    document.getElementById('cellRef').dispatchEvent(new KeyboardEvent('keydown',{key:'Enter',bubbles:true}));
    document.getElementById('fxIn').value='77';
    document.getElementById('fxIn').dispatchEvent(new KeyboardEvent('keydown',{key:'Enter',bubbles:true}));`);
  await evaluate(`document.getElementById('btnVba').click()`);
  const manual = await run('Sub Demo()\nDebug.Print Range("B3").Value\nEnd Sub');
  assert(manual.includes('77'), manual);
  assert(await evaluate(`!document.getElementById('btnSave').disabled`));
  await evaluate(`document.getElementById('vbaRun').dispatchEvent(new KeyboardEvent('keydown',{key:'Delete',bubbles:true}));`);
  const protectedValue = await run('Sub Demo()\nDebug.Print Range("B3").Value\nEnd Sub');
  assert(protectedValue.includes('77'), protectedValue);
  result.checks.push('manual edits survive read-only macros; dialog keys do not edit sheet');
  const failed = await run('Sub Demo()\nRange("A1").Value = "before failure"\nErr.Raise 5\nEnd Sub');
  assert(!failed.includes('Could not apply changes'), failed);
  const read = await run('Sub Demo()\nDebug.Print Range("A1").Value\nEnd Sub');
  assert(read.includes('before failure'), read);
  result.checks.push('changes before runtime error remain visible to next invocation');
  const filesOutput = await run('Sub Demo()\nMsgBox "UI message"\nOpen "demo.txt" For Output As #1\nPrint #1, "browser output"\nClose #1\nEnd Sub');
  assert(filesOutput.includes('UI message'), filesOutput);
  const fileContent = await evaluate(`(async()=>{const link=document.querySelector('#vbaDownloads a');
    return link ? (await fetch(link.href)).text() : null;})()`);
  assert(fileContent?.includes('browser output'), String(fileContent));
  result.checks.push('MsgBox and virtual file download are displayed');
  const invalid = await run('Sub Demo()\nEnd Sub', 'Demo', '{}');
  assert(/array/i.test(invalid), invalid);
  result.checks.push('invalid arguments reported without locking controls');
  await evaluate(`document.getElementById('vbaSource').value='Sub Demo()\\nDo\\nLoop\\nEnd Sub';
    document.getElementById('vbaArguments').value='[]'; document.getElementById('vbaRun').click();`);
  await waitUntil(() => evaluate(`!document.getElementById('vbaStop').disabled`));
  await evaluate(`document.getElementById('vbaStop').click()`);
  await waitUntil(() => evaluate(`!document.getElementById('vbaRun').disabled`));
  const recovered = await run('Sub Demo()\nDebug.Print "after stop"\nEnd Sub');
  assert(recovered.includes('after stop'), recovered);
  result.checks.push('Stop cancels worker and subsequent run recovers');
  const created = await run('Sub Demo()\nDim other As Workbook\nSet other = Workbooks.Add\nother.Worksheets(1).Range("A1").Value = "second workbook"\nDebug.Print Workbooks.Count\nEnd Sub');
  assert(!/error|not found/i.test(created), created);
  assert.equal(await evaluate(`document.getElementById('vbaWorkbook').options.length`), 2);
  const choices = await evaluate(`Array.from(document.getElementById('vbaWorkbook').options, o=>({value:o.value,text:o.textContent}))`);
  const other = choices.find(o => o.text !== 'sample.xlsx');
  await evaluate(`document.getElementById('vbaWorkbook').value=${JSON.stringify(other.value)};
    document.getElementById('vbaWorkbook').dispatchEvent(new Event('change'));`);
  await waitUntil(() => evaluate(`document.getElementById('docName').textContent===${JSON.stringify(other.text)}`));
  const second = await run('Sub Demo()\nDebug.Print Range("A1").Value\nEnd Sub');
  assert(second.includes('second workbook'), second);
  result.checks.push('macro-created workbook can be selected and executed');
  await run('Sub Demo()\nThisWorkbook.Close SaveChanges:=False\nEnd Sub');
  // Closing the selected secondary workbook selects the remaining one.
  await waitUntil(() => evaluate(`document.getElementById('docName').textContent==='sample.xlsx'`));
  const restored = await run('Sub Demo()\nDebug.Print Range("B3").Value\nEnd Sub');
  assert(restored.includes('77'), restored);
  await run('Sub Demo()\nThisWorkbook.Close SaveChanges:=False\nEnd Sub');
  assert.equal(await evaluate(`document.querySelector('#canvas #drop') !== null`), true);
  await evaluate(`document.getElementById('vbaDialog').close(); document.getElementById('btnVba').click()`);
  await waitUntil(() => evaluate(`document.getElementById('docName').textContent==='sample.xlsx'`));
  assert.equal(await evaluate(`document.getElementById('vbaProcedure').value`), 'Demo');
  result.checks.push('closing all workbooks resets the view and permits a fresh project');
  // Exercise the upload input and embedded project, with an optional local fixture.
  if (process.argv[4]) {
    const fixture = path.resolve(process.argv[4]);
    const before = hash(fs.readFileSync(fixture));
    await call('Page.reload');
    await waitUntil(() => evaluate(`document.getElementById('btnTryVba')?.disabled===false`));
    const document = await call('DOM.getDocument');
    const input = await call('DOM.querySelector', {nodeId: document.root.nodeId, selector: '#fileIn'});
    await call('DOM.setFileInputFiles', {nodeId: input.nodeId, files: [fixture]});
    await waitUntil(() => evaluate(`document.getElementById('docName').textContent===${JSON.stringify(path.basename(fixture))}`));
    await evaluate(`document.getElementById('btnVba').click()`);
    await waitUntil(() => evaluate(`document.getElementById('vbaDialog').open`));
    const imported = await evaluate(`document.getElementById('vbaSource').value`);
    assert(imported.includes('Twice = amount * 2'));
    const edited = imported.replace('Twice = amount * 2', 'Twice = amount * 3');
    const computed = await run(edited, 'Twice', '[7]');
    assert(computed.includes('21'), computed);
    await evaluate(`document.getElementById('vbaDialog').close(); document.getElementById('btnSave').click()`);
    const destination = path.join(downloads, path.basename(fixture));
    await waitUntil(() => fs.existsSync(destination));
    const uploadedSaved = fs.readFileSync(destination);
    const recoveredSource = await evaluate(`(async()=>{const w=await import('./oxidocs_wasm.js');
      return w.spreadsheet_vba_source(new Uint8Array(${JSON.stringify([...uploadedSaved])}));})()`);
    assert(recoveredSource.includes('Twice = amount * 3'));
    assert.equal(hash(fs.readFileSync(fixture)), before);
    result.checks.push('uploaded .xlsm code loads, executes, and survives download without changing original');
  }
  await evaluate(`document.getElementById('btnVba').click()`);
  await waitUntil(() => evaluate(`document.getElementById('vbaDialog').open`));
  const screenshot = await call('Page.captureScreenshot', {format: 'png'});
  fs.writeFileSync(path.join(output, 'vba-ui.png'), Buffer.from(screenshot.data, 'base64'));
  assert.equal(result.exceptions.length, 0, JSON.stringify(result.exceptions));
  for (const [file, expected] of Object.entries(result.files)) assert.equal(hash(fs.readFileSync(file)), expected, file);
  result.passed = true;
} catch (error) { result.error = String(error); result.stack = error.stack; }
finally {
  if (browserCall && !exited) { try { await browserCall('Browser.close'); } catch {} }
  pageSocket?.close(); browserSocket?.close();
  if (exit && !exited) await exit;
  if (server) { server.closeAllConnections(); await new Promise(resolve => server.close(resolve)); }
  if (exited && fs.existsSync(profile)) {
    assert.equal(path.dirname(fs.realpathSync(profile)), fs.realpathSync(output));
    assert.equal(path.basename(profile), 'owned-edge-profile');
    fs.rmSync(profile, {recursive: true, maxRetries: 12, retryDelay: 250});
    result.profile_removed = true;
  }
  result.passed &&= exited && result.browser_exit === 0 && result.profile_removed;
  result.completed = true;
  fs.writeFileSync(path.join(output, 'result.json'), JSON.stringify(result, null, 2) + '\n');
}
console.log(JSON.stringify({passed: result.passed, checks: result.checks, error: result.error, output}));
process.exitCode = result.passed ? 0 : 1;
