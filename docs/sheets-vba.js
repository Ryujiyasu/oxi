// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

import {runVbaInWorker, cancelActiveVbaExecution, closeVbaProject,
  parseVbaArguments, bindVbaWorkbookPicker, prepareVbaSaveInWorker,
  recordVbaSourceChange} from './vba-runner.js';
import {bindVbaFileDownloads} from './vba-file-downloads.js';

const EXAMPLE = `Sub Demo()
    Range("A1").Value = "Hello from VBA"
    Range("B3").Value = 42
    Debug.Print Range("A1").Value
End Sub
`;

// Workbook state and ZIP serialization belong to the sheet editor. The worker
// owns the VBA project and is never started by opening a file.
export function bindSheetVba({engine, capture: captureWorkbook, apply, closed, ensureWorkbook, language}) {
  const dialog = document.createElement('dialog');
  dialog.id = 'vbaDialog';
  dialog.setAttribute('aria-labelledby', 'vbaTitle');
  dialog.innerHTML = `<form method="dialog" class="vba-header"><h2 id="vbaTitle">Run VBA</h2><button id="vbaClose">Close</button></form>
    <p id="vbaScope"></p>
    <label><span id="vbaWorkbookLabel">Workbook</span> <select id="vbaWorkbook"></select></label>
    <div class="vba-actions"><button id="vbaLoad">Load workbook code</button><button id="vbaExample">Use example</button><button id="vbaExport">Download .bas</button><button id="vbaAddin">Load add-in (.xlam)</button><button id="vbaAddinCalc" hidden>Recalculate with add-ins</button><input type="file" id="vbaAddinFile" accept=".xlam,.xlsm" hidden></div>
    <p id="vbaAddins" class="vba-addins"></p>
    <label for="vbaSource" id="vbaSourceLabel">VBA source</label><textarea id="vbaSource" spellcheck="false" autocapitalize="off"></textarea>
    <div class="vba-actions"><label><span id="vbaProcedureLabel">Procedure</span> <input id="vbaProcedure" list="vbaProcedures" autocomplete="off" value="Demo"></label><datalist id="vbaProcedures"></datalist>
    <label><span id="vbaArgumentsLabel">Arguments (JSON array)</span> <input id="vbaArguments" value="[]" spellcheck="false"></label></div>
    <div class="vba-actions"><button id="vbaRun">Run</button><button id="vbaStop" disabled>Stop</button></div>
    <pre id="vbaOutput" role="status" aria-live="polite"></pre><div id="vbaDownloads"></div>`;
  document.body.append(dialog);
  const el = id => dialog.querySelector(`#${id}`);
  const editor = el('vbaSource');
  const output = el('vbaOutput');
  const files = bindVbaFileDownloads(el('vbaDownloads'));
  let busy = false;
  let saving = false;
  let bound = false;
  let revision = 0;
  let capturedBytes;
  const capture = () => {
    const state = captureWorkbook();
    if (state.bytes !== capturedBytes) { capturedBytes = state.bytes; revision++; }
    return state;
  };
  let hasRun = false;
  let previousSource = EXAMPLE;
  editor.value = EXAMPLE;
  const report = message => { output.textContent = String(message); };
  // Add-ins (.xlam): their standard modules join the project when code runs,
  // so their functions and macros are callable, and never when the workbook
  // is saved -- an add-in stays in its own file, as it does in Excel.
  let addins = [];
  const HOST_MODULE = 'Attribute VB_Name = "OxiAddinHost"\nPublic Sub OxiCalculateWithAddins()\n    Application.CalculateFull\nEnd Sub\n';
  function standardModules(source) {
    // Modules follow one another, each opening with its Attribute VB_Name line;
    // document and class modules carry a VERSION header and are left behind
    // (the open workbook has its own ThisWorkbook and sheets).
    const parts = source.split(/(?=^(?:VERSION 1\.0 CLASS\r?\n)?Attribute VB_Name = )/m);
    return parts.filter(part => /^Attribute VB_Name = /.test(part)).join('\n');
  }
  const withAddins = () => [editor.value, ...addins.map(addin => addin.source)].join('\n');
  function showAddins() {
    const ja = language() === 'ja';
    el('vbaAddins').textContent = addins.length
      ? (ja ? 'アドイン: ' : 'Add-ins: ') + addins.map(addin => addin.name).join(', ')
      : '';
    el('vbaAddinCalc').hidden = !addins.length;
  }
  function procedures() {
    try {
      const state = capture();
      const found = engine.list_spreadsheet_vba_project_procedures(withAddins(), state.bytes || new Uint8Array());
      el('vbaProcedures').replaceChildren(...found.map(proc => {
        const option = document.createElement('option');
        option.value = proc.name;
        option.label = `${proc.kind}: ${proc.required_parameter_count} required arguments`;
        return option;
      }));
    } catch { el('vbaProcedures').replaceChildren(); }
  }
  function applySelection(selection) {
    if (!selection.package_bytes) throw new Error('This workbook has no save package.');
    const selected = selection.selected;
    const bytes = new Uint8Array(selection.package_bytes);
    const updated = engine.edit_xlsx_from_workbook(bytes, selected.workbook);
    apply({workbook: selected.workbook, bytes: updated, name: selected.name,
      activeSheet: selected.active_sheet, dirty: !selected.saved, selection: true});
    editor.value = selected.source || '';
    bound = typeof engine.spreadsheet_vba_source(bytes) === 'string';
    revision++;
    procedures();
  }
  const picker = bindVbaWorkbookPicker(el('vbaWorkbook'), {
    capture: () => { const s = capture(); return {workbook: s.workbook,
      activeSheet: s.activeSheet, editRevision: revision}; },
    apply: applySelection, busy: () => busy, report,
  });
  function setBusy(value) {
    busy = value;
    editor.disabled = value;
    for (const id of ['vbaRun', 'vbaLoad', 'vbaExample', 'vbaExport', 'vbaProcedure', 'vbaArguments', 'vbaAddin', 'vbaAddinCalc']) el(id).disabled = value;
    el('vbaStop').disabled = !value;
    picker.refresh();
  }
  function accept(execution) {
    hasRun = !execution.workbook_closed || Boolean(execution.workbook_selection);
    files.show(execution);
    picker.update(execution, capture().name);
    if (execution.workbook_closed) {
      if (!execution.workbook_selection) closed();
      return;
    }
    if (execution.workbook_selection) return;
    const state = capture();
    let bytes = engine.edit_xlsx_from_workbook(state.bytes, execution.workbook);
    // Uploaded projects retain edited code; scratch examples are exported as .bas.
    if (bound) bytes = engine.save_spreadsheet_vba_source(bytes, editor.value);
    apply({workbook: execution.workbook, bytes, name: state.name,
      activeSheet: execution.active_sheet, dirty: execution.workbook_changed !== false || bound});
    revision++;
  }
  function describe(execution) {
    const lines = [];
    if (execution.messages?.length) lines.push(...execution.messages.map(m => `${m.title || 'MsgBox'}: ${m.prompt}`));
    if (execution.debug_output?.length) lines.push(...execution.debug_output);
    if (execution.result) lines.push(`Result: ${JSON.stringify(execution.result)}`);
    if (execution.workbook_closed) lines.push('Workbook closed. Open a workbook to start a new project.');
    return lines.join('\n') || 'Completed.';
  }
  el('vbaRun').onclick = async () => {
    if (busy) return;
    try {
      const args = parseVbaArguments(el('vbaArguments').value);
      const state = capture();
      if (!state.workbook) throw new Error('Open a workbook first.');
      if (!el('vbaProcedure').value.trim()) throw new Error('Choose a procedure.');
      setBusy(true);
      report('Running…');
      const execution = await runVbaInWorker(state.workbook, withAddins(),
        el('vbaProcedure').value.trim(), args, state.activeSheet, state.name,
        {errorLanguage: language() === 'ja' ? 'ja-JP' : 'en-US'}, state.bytes);
      accept(execution);
      report(describe(execution));
    } catch (error) {
      if (error.execution) {
        try { accept(error.execution); }
        catch (saveError) { report(`${error.message}\nCould not apply changes: ${saveError.message}`); return; }
      }
      report(`${error.execution ? describe(error.execution) + '\n' : ''}${error.message}`);
    } finally { setBusy(false); }
  };
  el('vbaStop').onclick = () => cancelActiveVbaExecution();
  async function calculateWithAddins() {
    const state = capture();
    if (!state.workbook || busy) return;
    setBusy(true);
    report(language() === 'ja' ? 'アドインの関数で再計算中…' : 'Recalculating with add-in functions…');
    try {
      const execution = await runVbaInWorker(state.workbook, withAddins() + '\n' + HOST_MODULE,
        'OxiCalculateWithAddins', [], state.activeSheet, state.name,
        {errorLanguage: language() === 'ja' ? 'ja-JP' : 'en-US'}, state.bytes);
      accept(execution);
      report(language() === 'ja' ? 'アドインの関数で再計算しました。' : 'Recalculated with add-in functions.');
    } catch (error) {
      if (error.execution) { try { accept(error.execution); } catch { /* reported below */ } }
      report(error.message);
    } finally { setBusy(false); }
  }
  el('vbaAddin').onclick = () => el('vbaAddinFile').click();
  el('vbaAddinFile').onchange = async () => {
    const file = el('vbaAddinFile').files?.[0];
    el('vbaAddinFile').value = '';
    if (!file) return;
    try {
      const source = engine.spreadsheet_vba_source(new Uint8Array(await file.arrayBuffer()));
      const modules = typeof source === 'string' ? standardModules(source) : '';
      if (!modules) { report(language() === 'ja' ? 'このファイルには標準モジュールがありません。' : 'That file holds no standard modules.'); return; }
      addins = addins.filter(addin => addin.name !== file.name).concat([{name: file.name, source: modules}]);
      showAddins();
      procedures();
      await calculateWithAddins();
    } catch (error) { report(error.message ?? String(error)); }
  };
  el('vbaAddinCalc').onclick = () => calculateWithAddins();
  dialog.addEventListener('close', () => {
    // A queued close event must not cancel a Save started after the editor closed.
    if (!dialog.open && busy && !saving) cancelActiveVbaExecution();
  });
  el('vbaExample').onclick = () => {
    if (editor.value !== EXAMPLE && !confirm('Replace the code in the editor?')) return;
    editor.value = EXAMPLE; bound = false; el('vbaProcedure').value = 'Demo'; procedures();
  };
  el('vbaLoad').onclick = () => {
    if (editor.value && !confirm('Replace the code in the editor with workbook code?')) return;
    const source = engine.spreadsheet_vba_source(capture().bytes);
    if (typeof source !== 'string') { report('This workbook has no VBA project.'); return; }
    editor.value = source; bound = true; procedures();
    el('vbaProcedure').value = el('vbaProcedures').firstChild?.value || '';
  };
  editor.addEventListener('input', () => {
    recordVbaSourceChange(previousSource, editor.value);
    previousSource = editor.value;
    procedures();
  });
  el('vbaExport').onclick = () => {
    const url = URL.createObjectURL(new Blob([editor.value], {type: 'text/plain;charset=utf-8'}));
    const link = document.createElement('a'); link.href = url; link.download = 'Module.bas'; link.click();
    setTimeout(() => URL.revokeObjectURL(url), 4000);
  };
  return {
    async open() {
      await ensureWorkbook();
      const state = capture();
      if (!state.workbook) return;
      const ja = language() === 'ja';
      const labels = {
        vbaTitle: ['Run VBA', 'VBAを実行'], vbaClose: ['Close', '閉じる'],
        vbaWorkbookLabel: ['Workbook', 'ブック'], vbaSourceLabel: ['VBA source', 'VBAコード'],
        vbaProcedureLabel: ['Procedure', 'プロシージャ'], vbaArgumentsLabel: ['Arguments (JSON array)', '引数（JSON配列）'],
        vbaLoad: ['Load workbook code', 'ブックのコードを読み込む'], vbaExample: ['Use example', 'サンプルコード'],
        vbaExport: ['Download .bas', '.basを保存'], vbaRun: ['Run', '実行'], vbaStop: ['Stop', '停止'],
        vbaAddin: ['Load add-in (.xlam)', 'アドインを読み込む (.xlam)'], vbaAddinCalc: ['Recalculate with add-ins', 'アドインで再計算'],
      };
      for (const [id, text] of Object.entries(labels)) el(id).textContent = text[ja ? 1 : 0];
      el('vbaScope').textContent = language() === 'ja'
        ? 'Oxi独自のVBAインタープリターで実行します。OfficeやWindows COMは呼び出しません。互換性は検証中です。UserFormの表示、任意のActiveX、Windows DLLの実行には対応していません。実行中は停止できます。セルの変更は「コピーを保存」で保存できます。サンプルのコードは .bas として保存してください。'
        : 'Runs in Oxi’s VBA interpreter, without Office or Windows COM. Compatibility is in progress. UserForm display, arbitrary ActiveX and Windows DLL execution are unavailable here. Stop cancels execution. Save a copy keeps workbook edits; download scratch code separately as .bas.';
      previousSource = editor.value;
      showAddins();
      procedures(); dialog.showModal();
    },
    async prepareSave() {
      if (busy) throw new Error('VBA is running. Stop it before saving.');
      if (!hasRun) return;
      const state = capture();
      saving = true;
      setBusy(true);
      try {
        const execution = await prepareVbaSaveInWorker(state.workbook, editor.value,
          state.activeSheet, state.name, {}, state.bytes);
        accept(execution);
      } catch (error) {
        if (error.execution) accept(error.execution);
        throw error;
      } finally { saving = false; setBusy(false); }
    },
    reset(name, bytes) {
      closeVbaProject(); files.clear(); picker.reset(name); report(''); revision++; hasRun = false;
      const source = engine.spreadsheet_vba_source(bytes);
      bound = typeof source === 'string';
      editor.value = bound ? source : EXAMPLE;
      previousSource = editor.value;
      el('vbaProcedure').value = bound ? '' : 'Demo';
      procedures();
      if (bound) el('vbaProcedure').value = el('vbaProcedures').firstChild?.value || '';
    },
    get busy() { return busy; },
  };
}
