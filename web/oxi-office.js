// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! A stand-in for office.js: the Office and Excel JavaScript API an add-in
//! programs against, answered by Oxi's workbook instead of by Excel.
//!
//! An Office add-in is a web page that loads office.js and talks to the
//! workbook through `Excel.run(async context => { ... await context.sync(); })`.
//! The proxy objects it gets back queue what the add-in asks for -- a value
//! written, a property to `load()` -- and `context.sync()` carries the queue to
//! the host in one go and fills in what was loaded. Reading a property that was
//! not loaded and synced throws `PropertyNotLoaded`, as Excel's does; it is
//! the commonest mistake in add-in code, and an add-in that works here should
//! not then fail in Excel.
//!
//! The host is attached with `OxiOffice.attach(adapter)`:
//!
//!   sheetNames()                 -> string[]
//!   activeSheet()                -> index;   activate(index)
//!   addSheet(name)               -> index
//!   read(sheet, row, col)        -> { value, formula, numberFormat, text,
//!                                     fill, fontColor, bold, italic, fontName, fontSize }
//!   write(edits)                 -> Promise|void; edits are
//!                                   { sheet, row, col, value?, formula?, numberFormat?,
//!                                     fill?, fontColor?, bold?, italic?, clear? }
//!   usedRange(sheet)             -> { row, col, rows, cols }
//!   selection()                  -> { sheet, row, col, rows, cols }
//!
//! Rows and columns count from zero. Values are numbers, strings, booleans or
//! '' for an empty cell, as Excel's JavaScript API hands them over.

(function (root) {
  'use strict';

  let host = null;
  const readyWaiters = [];

  class OfficeExtensionError extends Error {
    constructor(code, message) {
      super(message);
      this.name = 'RichApi.Error';
      this.code = code;
      this.debugInfo = { code, message };
    }
  }

  // ---- addresses ----------------------------------------------------------
  function columnName(index) {
    let out = '';
    for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) {
      out = String.fromCharCode(65 + ((n - 1) % 26)) + out;
    }
    return out;
  }
  function columnIndex(letters) {
    let n = 0;
    for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n - 1;
  }
  /** "A1:B2" / "$A$1" / "C:C" / "3:3" -> { row, col, rows, cols } */
  function parseArea(text) {
    const [a, b = a] = text.replace(/\$/g, '').split(':');
    const one = (part, end) => {
      const m = /^([A-Za-z]*)(\d*)$/.exec(part);
      if (!m) throw new OfficeExtensionError('InvalidArgument', `The address ${text} is not valid.`);
      const col = m[1] ? columnIndex(m[1]) : (end ? 16383 : 0);
      const row = m[2] ? Number(m[2]) - 1 : (end ? 1048575 : 0);
      return { row, col };
    };
    const p = one(a, false);
    const q = one(b, true);
    const r0 = Math.min(p.row, q.row), r1 = Math.max(p.row, q.row);
    const c0 = Math.min(p.col, q.col), c1 = Math.max(p.col, q.col);
    return { row: r0, col: c0, rows: r1 - r0 + 1, cols: c1 - c0 + 1 };
  }
  function addressOf(sheetName, area) {
    const quote = /^[A-Za-z_][A-Za-z0-9_]*$/.test(sheetName) ? sheetName : `'${sheetName.replace(/'/g, "''")}'`;
    const first = `${columnName(area.col)}${area.row + 1}`;
    const last = `${columnName(area.col + area.cols - 1)}${area.row + area.rows}`;
    return `${quote}!${first}${area.rows === 1 && area.cols === 1 ? '' : ':' + last}`;
  }

  // ---- the request context -------------------------------------------------
  class RequestContext {
    constructor() {
      this._writes = [];
      this._loads = [];
      this.workbook = new Workbook(this);
      this.trackedObjects = { add() {}, remove() {} };
    }
    async sync() {
      if (!host) throw new OfficeExtensionError('GeneralException', 'No workbook is attached.');
      const writes = this._writes.splice(0);
      for (const w of writes) {
        if (w.kind === 'cells') await host.write(w.edits);
        else if (w.kind === 'activate') host.activate(w.sheet);
        else if (w.kind === 'addSheet') w.done(host.addSheet(w.name));
      }
      const loads = this._loads.splice(0);
      for (const fill of loads) fill();
    }
    load(object, props) {
      object.load(props);
      return object;
    }
  }

  /** Base for proxies: load() queues a fill-in; reading before it throws. */
  class Proxy {
    constructor(context) {
      Object.defineProperty(this, '_ctx', { value: context });
      Object.defineProperty(this, '_loaded', { value: new Map(), writable: true });
    }
    _get(name) {
      if (!this._loaded.has(name)) {
        throw new OfficeExtensionError('PropertyNotLoaded',
          `The property '${name}' is not available. Before reading the property's value, call the load method on the containing object and call "context.sync()" on the associated request context.`);
      }
      return this._loaded.get(name);
    }
    load(props) {
      const names = props == null ? this._defaultProps()
        : Array.isArray(props) ? props
        : typeof props === 'string' ? props.split(',').map(s => s.trim())
        : props.select ? props.select.split(',').map(s => s.trim())
        : Object.keys(props);
      this._ctx._loads.push(() => {
        for (const name of names) {
          const [head, ...rest] = name.split('/');
          const child = Object.prototype.hasOwnProperty.call(this, head) ? this[head] : null;
          if (rest.length && child && typeof child._fill === 'function') {
            child._fill(rest.join('/'));
          } else {
            this._fill(head);
          }
        }
      });
      return this;
    }
    _defaultProps() { return []; }
    _fill() {}
    toJSON() { return Object.fromEntries(this._loaded); }
  }

  // ---- workbook and sheets -------------------------------------------------
  class Workbook extends Proxy {
    constructor(context) {
      super(context);
      this.worksheets = new WorksheetCollection(context);
    }
    getSelectedRange() {
      const s = host.selection();
      return new Range(this._ctx, s.sheet, { row: s.row, col: s.col, rows: s.rows, cols: s.cols });
    }
    getActiveCell() {
      const s = host.selection();
      return new Range(this._ctx, s.sheet, { row: s.row, col: s.col, rows: 1, cols: 1 });
    }
  }

  class WorksheetCollection extends Proxy {
    getActiveWorksheet() { return new Worksheet(this._ctx, () => host.activeSheet()); }
    getItem(name) {
      return new Worksheet(this._ctx, () => {
        const i = host.sheetNames().findIndex(n => n.toLowerCase() === String(name).toLowerCase());
        if (i < 0) throw new OfficeExtensionError('ItemNotFound', `The requested resource doesn't exist.`);
        return i;
      });
    }
    getItemOrNullObject(name) {
      const sheet = this.getItem(name);
      Object.defineProperty(sheet, 'isNullObject', {
        get: () => !host.sheetNames().some(n => n.toLowerCase() === String(name).toLowerCase()),
      });
      return sheet;
    }
    getFirst() { return new Worksheet(this._ctx, () => 0); }
    getCount() {
      const result = { value: undefined };
      this._ctx._loads.push(() => { result.value = host.sheetNames().length; });
      return result;
    }
    add(name) {
      let index = null;
      const sheet = new Worksheet(this._ctx, () => {
        if (index === null) throw new OfficeExtensionError('InvalidOperation', 'Call context.sync() first.');
        return index;
      });
      this._ctx._writes.push({ kind: 'addSheet', name: name ?? `Sheet${host.sheetNames().length + 1}`, done: i => { index = i; } });
      return sheet;
    }
    _defaultProps() { return ['items']; }
    _fill(name) {
      if (name === 'items' || name.startsWith('items')) {
        this._loaded.set('items', host.sheetNames().map((_, i) => {
          const ws = new Worksheet(this._ctx, () => i);
          ws._fill('name'); ws._fill('position');
          return ws;
        }));
      }
    }
    get items() { return this._get('items'); }
  }

  class Worksheet extends Proxy {
    constructor(context, index) {
      super(context);
      Object.defineProperty(this, '_index', { value: index });
    }
    get name() { return this._get('name'); }
    get position() { return this._get('position'); }
    get id() { return this._get('id'); }
    _defaultProps() { return ['name', 'position', 'id']; }
    _fill(prop) {
      const i = this._index();
      if (prop === 'name') this._loaded.set('name', host.sheetNames()[i]);
      else if (prop === 'position') this._loaded.set('position', i);
      else if (prop === 'id') this._loaded.set('id', `{oxi-sheet-${i}}`);
    }
    getRange(address) {
      if (!address) {
        return new Range(this._ctx, this._index, null);
      }
      const text = String(address);
      const area = parseArea(text.includes('!') ? text.split('!').pop() : text);
      return new Range(this._ctx, this._index, area);
    }
    getRangeByIndexes(row, col, rows, cols) {
      return new Range(this._ctx, this._index, { row, col, rows, cols });
    }
    getCell(row, col) { return new Range(this._ctx, this._index, { row, col, rows: 1, cols: 1 }); }
    getUsedRange() {
      return new Range(this._ctx, this._index, () => {
        const u = host.usedRange(this._index());
        return { row: u.row, col: u.col, rows: Math.max(1, u.rows), cols: Math.max(1, u.cols) };
      });
    }
    activate() { this._ctx._writes.push({ kind: 'activate', sheet: this._index() }); }
  }

  // ---- ranges ------------------------------------------------------------
  const CELL_PROPS = ['address', 'values', 'formulas', 'numberFormat', 'text', 'rowCount', 'columnCount', 'rowIndex', 'columnIndex', 'valueTypes'];

  class Range extends Proxy {
    constructor(context, sheetIndex, area) {
      super(context);
      Object.defineProperty(this, '_sheet', { value: typeof sheetIndex === 'function' ? sheetIndex : () => sheetIndex });
      Object.defineProperty(this, '_areaOf', {
        value: typeof area === 'function' ? area : () => area ?? { row: 0, col: 0, rows: 1048576, cols: 16384 },
      });
      this.format = new RangeFormat(context, this);
    }
    _area() { return this._areaOf(); }
    _cells(fn) {
      const a = this._area();
      const out = [];
      for (let r = 0; r < a.rows; r++) {
        const line = [];
        for (let c = 0; c < a.cols; c++) line.push(fn(host.read(this._sheet(), a.row + r, a.col + c), r, c));
        out.push(line);
      }
      return out;
    }
    _defaultProps() { return CELL_PROPS; }
    _fill(prop) {
      const a = this._area();
      switch (prop) {
        case 'address': this._loaded.set('address', addressOf(host.sheetNames()[this._sheet()], a)); break;
        case 'values': this._loaded.set('values', this._cells(c => c?.value ?? '')); break;
        case 'formulas': this._loaded.set('formulas', this._cells(c => c?.formula ? (c.formula.startsWith('=') ? c.formula : '=' + c.formula) : (c?.value ?? ''))); break;
        case 'numberFormat': this._loaded.set('numberFormat', this._cells(c => c?.numberFormat ?? 'General')); break;
        case 'text': this._loaded.set('text', this._cells(c => c?.text ?? String(c?.value ?? ''))); break;
        case 'valueTypes': this._loaded.set('valueTypes', this._cells(c => {
          const v = c?.value;
          return v === '' || v == null ? 'Empty' : typeof v === 'number' ? 'Double' : typeof v === 'boolean' ? 'Boolean' : String(v).startsWith('#') ? 'Error' : 'String';
        })); break;
        case 'rowCount': this._loaded.set('rowCount', a.rows); break;
        case 'columnCount': this._loaded.set('columnCount', a.cols); break;
        case 'rowIndex': this._loaded.set('rowIndex', a.row); break;
        case 'columnIndex': this._loaded.set('columnIndex', a.col); break;
        default: break;
      }
    }
    get address() { return this._get('address'); }
    get rowCount() { return this._get('rowCount'); }
    get columnCount() { return this._get('columnCount'); }
    get rowIndex() { return this._get('rowIndex'); }
    get columnIndex() { return this._get('columnIndex'); }
    get text() { return this._get('text'); }
    get valueTypes() { return this._get('valueTypes'); }
    get values() { return this._get('values'); }
    set values(grid) { this._queueGrid(grid, (value) => (typeof value === 'string' && value.startsWith('=') ? { formula: value } : { value })); }
    get formulas() { return this._get('formulas'); }
    set formulas(grid) { this._queueGrid(grid, (value) => (typeof value === 'string' && value.startsWith('=') ? { formula: value } : { value })); }
    get numberFormat() { return this._get('numberFormat'); }
    set numberFormat(grid) { this._queueGrid(grid, (numberFormat) => ({ numberFormat })); }
    _queueGrid(grid, make) {
      const a = this._area();
      // A single value fills the whole range, as the API allows.
      const at = (r, c) => (Array.isArray(grid) ? grid[r]?.[c] : grid);
      if (Array.isArray(grid) && (grid.length !== a.rows || grid.some(line => !Array.isArray(line) || line.length !== a.cols))) {
        throw new OfficeExtensionError('InvalidArgument',
          `The number of rows or columns in the input array doesn't match the size or dimensions of the range.`);
      }
      const edits = [];
      for (let r = 0; r < a.rows; r++) {
        for (let c = 0; c < a.cols; c++) {
          edits.push({ sheet: this._sheet(), row: a.row + r, col: a.col + c, ...make(at(r, c)) });
        }
      }
      this._ctx._writes.push({ kind: 'cells', edits });
    }
    getCell(row, col) {
      const a = this._area();
      return new Range(this._ctx, this._sheet, { row: a.row + row, col: a.col + col, rows: 1, cols: 1 });
    }
    getOffsetRange(rows, cols) {
      const a = this._area();
      return new Range(this._ctx, this._sheet, { ...a, row: a.row + rows, col: a.col + cols });
    }
    getResizedRange(rows, cols) {
      const a = this._area();
      return new Range(this._ctx, this._sheet, { ...a, rows: a.rows + rows, cols: a.cols + cols });
    }
    getRow(i) { const a = this._area(); return new Range(this._ctx, this._sheet, { ...a, row: a.row + i, rows: 1 }); }
    getColumn(i) { const a = this._area(); return new Range(this._ctx, this._sheet, { ...a, col: a.col + i, cols: 1 }); }
    getLastRow() { const a = this._area(); return this.getRow(a.rows - 1); }
    getLastColumn() { const a = this._area(); return this.getColumn(a.cols - 1); }
    clear(applyTo) {
      const a = this._area();
      const edits = [];
      for (let r = 0; r < a.rows; r++) for (let c = 0; c < a.cols; c++) {
        edits.push({ sheet: this._sheet(), row: a.row + r, col: a.col + c, clear: applyTo ?? 'All' });
      }
      this._ctx._writes.push({ kind: 'cells', edits });
    }
    select() {}
  }

  class RangeFormat extends Proxy {
    constructor(context, range) {
      super(context);
      Object.defineProperty(this, '_range', { value: range });
      this.fill = new RangeFill(context, range);
      this.font = new RangeFont(context, range);
    }
    autofitColumns() {}
    autofitRows() {}
  }

  class RangeFill extends Proxy {
    constructor(context, range) { super(context); Object.defineProperty(this, '_range', { value: range }); }
    get color() { return this._get('color'); }
    set color(value) { this._range._queueGrid(value, fill => ({ fill })); }
    clear() { this._range._queueGrid(null, () => ({ fill: null })); }
    _defaultProps() { return ['color']; }
    _fill(prop) {
      if (prop === 'color') {
        const top = this._range._cells(c => c?.fill ?? null);
        const flat = top.flat();
        this._loaded.set('color', flat.every(v => v === flat[0]) ? (flat[0] ?? '#FFFFFF') : null);
      }
    }
  }

  class RangeFont extends Proxy {
    constructor(context, range) { super(context); Object.defineProperty(this, '_range', { value: range }); }
    _defaultProps() { return ['bold', 'italic', 'color', 'name', 'size']; }
    _fill(prop) {
      const key = { bold: 'bold', italic: 'italic', color: 'fontColor', name: 'fontName', size: 'fontSize' }[prop];
      if (!key) return;
      const flat = this._range._cells(c => c?.[key] ?? null).flat();
      const fallback = { bold: false, italic: false, color: '#000000', name: 'Calibri', size: 11 }[prop];
      this._loaded.set(prop, flat.every(v => v === flat[0]) ? (flat[0] ?? fallback) : null);
    }
    get bold() { return this._get('bold'); }
    set bold(v) { this._range._queueGrid(v, bold => ({ bold })); }
    get italic() { return this._get('italic'); }
    set italic(v) { this._range._queueGrid(v, italic => ({ italic })); }
    get color() { return this._get('color'); }
    set color(v) { this._range._queueGrid(v, fontColor => ({ fontColor })); }
    get name() { return this._get('name'); }
    get size() { return this._get('size'); }
  }

  // ---- Office and Excel namespaces -------------------------------------------
  const Excel = {
    run(batch) {
      const context = new RequestContext();
      return Promise.resolve()
        .then(() => batch(context))
        .then(async result => { await context.sync(); return result; });
    },
    RequestContext,
    ErrorCodes: {
      invalidArgument: 'InvalidArgument',
      itemNotFound: 'ItemNotFound',
      generalException: 'GeneralException',
      propertyNotLoaded: 'PropertyNotLoaded',
    },
    ClearApplyTo: { all: 'All', formats: 'Formats', contents: 'Contents' },
  };

  const Office = {
    HostType: { Excel: 'Excel', Word: 'Word', PowerPoint: 'PowerPoint' },
    PlatformType: { OfficeOnline: 'OfficeOnline' },
    initialize: undefined,
    onReady(callback) {
      const info = { host: 'Excel', platform: 'OfficeOnline' };
      const promise = new Promise(resolve => {
        if (host) resolve(info);
        else readyWaiters.push(() => resolve(info));
      });
      if (callback) promise.then(callback);
      return promise;
    },
    context: {
      host: 'Excel',
      platform: 'OfficeOnline',
      diagnostics: { host: 'Excel', platform: 'OfficeOnline', version: '16.0' },
      requirements: { isSetSupported: (name, version) => name === 'ExcelApi' && Number(version ?? 1) <= 1.4 },
      displayLanguage: 'ja-JP',
    },
  };

  const OxiOffice = {
    attach(adapter) {
      host = adapter;
      if (typeof Office.initialize === 'function') Office.initialize('');
      for (const wake of readyWaiters.splice(0)) wake();
    },
    detach() { host = null; },
    columnName,
    parseArea,
  };

  // ---- remote host: an add-in page in an iframe, the workbook in its parent --
  //
  // An add-in runs on its own origin, so it cannot touch the editor's objects.
  // Each sync posts the queued writes to the parent and gets back a mirror of
  // the workbook -- sheet names, the active sheet, the selection and every
  // cell -- which answers the reads that follow until the next sync.
  function remoteAdapter(target) {
    let mirror = { names: [], active: 0, selection: { sheet: 0, row: 0, col: 0, rows: 1, cols: 1 }, sheets: [] };
    let pending = [];
    let seq = 0;
    const waiting = new Map();
    root.addEventListener('message', event => {
      const data = event.data;
      if (!data || data.oxiOffice !== 'reply' || !waiting.has(data.id)) return;
      const resolve = waiting.get(data.id);
      waiting.delete(data.id);
      resolve(data);
    });
    const call = (op, payload) => new Promise(resolve => {
      const id = ++seq;
      waiting.set(id, resolve);
      target.postMessage({ oxiOffice: 'request', id, op, ...payload }, '*');
    });
    const adopt = snapshot => {
      mirror = snapshot;
      mirror.cells = mirror.sheets.map(sheet => new Map(sheet.cells.map(([r, c, cell]) => [`${r},${c}`, cell])));
    };
    return {
      async refresh() { adopt((await call('snapshot', {})).snapshot); },
      sheetNames: () => mirror.names,
      activeSheet: () => mirror.active,
      activate: index => { pending.push({ op: 'activate', index }); },
      addSheet: name => { pending.push({ op: 'addSheet', name }); mirror.names = [...mirror.names, name]; return mirror.names.length - 1; },
      read: (s, r, c) => mirror.cells?.[s]?.get(`${r},${c}`) ?? { value: '' },
      write: edits => { pending.push({ op: 'write', edits }); },
      usedRange: s => mirror.sheets[s]?.used ?? { row: 0, col: 0, rows: 1, cols: 1 },
      selection: () => mirror.selection,
      async flush() {
        const batch = pending.splice(0);
        const reply = await call('sync', { batch });
        if (reply.error) throw new OfficeExtensionError(reply.error.code ?? 'GeneralException', reply.error.message ?? 'The host refused the change.');
        adopt(reply.snapshot);
      },
    };
  }
  // A remote host flushes once per sync, after the writes are queued and
  // before the loads read the mirror.
  const plainSync = RequestContext.prototype.sync;
  RequestContext.prototype.sync = async function () {
    if (host && typeof host.flush === 'function') {
      const writes = this._writes.splice(0);
      for (const w of writes) {
        if (w.kind === 'cells') host.write(w.edits);
        else if (w.kind === 'activate') host.activate(w.sheet);
        else if (w.kind === 'addSheet') w.done(host.addSheet(w.name));
      }
      await host.flush();
      const loads = this._loads.splice(0);
      for (const fill of loads) fill();
      return;
    }
    return plainSync.call(this);
  };
  if (typeof window !== 'undefined' && window.parent && window.parent !== window) {
    const remote = remoteAdapter(window.parent);
    remote.refresh().then(() => OxiOffice.attach(remote));
  }

  const OfficeExtension = { Error: OfficeExtensionError };
  root.Office = Office;
  root.Excel = Excel;
  root.OfficeExtension = OfficeExtension;
  root.OxiOffice = OxiOffice;
  if (typeof module === 'object' && module.exports) module.exports = { Office, Excel, OfficeExtension, OxiOffice };
})(typeof globalThis !== 'undefined' ? globalThis : this);
