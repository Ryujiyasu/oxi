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
        else if (typeof host.op === 'function') {
          const answer = await host.op(settle(w));
          if (w.done) w.done(answer);
        }
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

  // A queued operation with its late-bound parts (a sheet index, a range's
  // address) worked out, ready for the host: sync time is when they exist.
  function settle(w) {
    const out = {};
    for (const [k, v] of Object.entries(w)) {
      if (k === 'done') continue;
      out[k] = typeof v === 'function' ? v() : v;
    }
    return out;
  }
  function absoluteAddress(sheetName, a) {
    const quoted = /^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheetName) ? sheetName : `'${sheetName.replace(/'/g, "''")}'`;
    const one = (r, c) => `$${columnName(c)}$${r + 1}`;
    const cells = a.rows === 1 && a.cols === 1 ? one(a.row, a.col) : `${one(a.row, a.col)}:${one(a.row + a.rows - 1, a.col + a.cols - 1)}`;
    return `${quoted}!${cells}`;
  }
  // "=Sheet1!$A$1:$B$2" -> { sheet, area } on this workbook, or null.
  function referenceOf(formula) {
    const text = String(formula ?? '').replace(/^=/, '');
    const bang = text.lastIndexOf('!');
    if (bang < 0) return null;
    const name = text.slice(0, bang).replace(/^'|'$/g, '').replace(/''/g, "'");
    const sheet = host.sheetNames().findIndex(n => n.toLowerCase() === name.toLowerCase());
    if (sheet < 0) return null;
    return { sheet, area: parseArea(text.slice(bang + 1).replace(/\$/g, '')) };
  }
  function notFound() {
    return new OfficeExtensionError('ItemNotFound', `The requested resource doesn't exist.`);
  }

  // ---- events --------------------------------------------------------------
  // `worksheet.onChanged.add(handler)`: the host tells of edits made in the
  // workbook (by the user, not by this add-in's own writes) and of selection
  // moves; each handler gets Excel's event arguments.
  class EventSource {
    constructor(context, type, sheetIndex) {
      Object.defineProperty(this, '_ctx', { value: context });
      Object.defineProperty(this, '_type', { value: type });
      Object.defineProperty(this, '_sheet', { value: sheetIndex });
    }
    add(handler) {
      const off = typeof host?.on === 'function'
        ? host.on(this._type, event => {
          const here = this._sheet ? this._sheet() : null;
          if (here !== null && event.sheet !== here) return undefined;
          const args = {
            type: this._type === 'changed' ? 'WorksheetChanged' : 'WorksheetSelectionChanged',
            address: event.address,
            worksheetId: `{oxi-sheet-${event.sheet}}`,
            source: 'Local',
          };
          if (this._type === 'changed') args.changeType = 'RangeEdited';
          return handler(args);
        })
        : () => {};
      return { remove: () => off() };
    }
  }

  // ---- names ---------------------------------------------------------------
  class NamedItemCollection extends Proxy {
    add(name, reference, comment) {
      const formula = typeof reference === 'string'
        ? () => (reference.startsWith('=') ? reference : '=' + reference)
        : () => '=' + absoluteAddress(host.sheetNames()[reference._sheet()], reference._area());
      this._ctx._writes.push({ kind: 'name', name, formula, comment: comment ?? '' });
      return new NamedItem(this._ctx, name);
    }
    getItem(name) { return new NamedItem(this._ctx, name); }
    getItemOrNullObject(name) {
      const item = this.getItem(name);
      Object.defineProperty(item, 'isNullObject', { get: () => !item._find(false) });
      return item;
    }
    getCount() {
      const result = { value: undefined };
      this._ctx._loads.push(() => { result.value = (host.names?.() ?? []).length; });
      return result;
    }
    _defaultProps() { return ['items']; }
    _fill(name) {
      if (name === 'items' || name.startsWith('items')) {
        this._loaded.set('items', (host.names?.() ?? []).map(n => {
          const item = new NamedItem(this._ctx, n.name);
          for (const p of item._defaultProps()) item._fill(p);
          return item;
        }));
      }
    }
    get items() { return this._get('items'); }
  }

  class NamedItem extends Proxy {
    constructor(context, name) {
      super(context);
      Object.defineProperty(this, '_name', { value: name });
    }
    _find(strict = true) {
      const found = (host.names?.() ?? []).find(n => n.name.toLowerCase() === String(this._name).toLowerCase());
      if (!found && strict) throw notFound();
      return found;
    }
    _defaultProps() { return ['name', 'formula', 'value', 'type', 'visible']; }
    _fill(prop) {
      const n = this._find();
      if (prop === 'name') this._loaded.set('name', n.name);
      else if (prop === 'formula') this._loaded.set('formula', n.formula);
      else if (prop === 'value') this._loaded.set('value', n.formula.replace(/^=/, ''));
      else if (prop === 'type') this._loaded.set('type', referenceOf(n.formula) ? 'Range' : 'String');
      else if (prop === 'visible') this._loaded.set('visible', n.visible !== false);
    }
    get name() { return this._get('name'); }
    get formula() { return this._get('formula'); }
    get value() { return this._get('value'); }
    get type() { return this._get('type'); }
    get visible() { return this._get('visible'); }
    getRange() {
      const at = () => {
        const ref = referenceOf(this._find().formula);
        if (!ref) throw new OfficeExtensionError('InvalidArgument', 'The name does not refer to a range.');
        return ref;
      };
      return new Range(this._ctx, () => at().sheet, () => at().area);
    }
    delete() { this._ctx._writes.push({ kind: 'deleteName', name: this._name }); }
  }

  // ---- tables --------------------------------------------------------------
  class TableCollection extends Proxy {
    constructor(context, sheetIndex) {
      super(context);
      Object.defineProperty(this, '_sheet', { value: sheetIndex });
    }
    _list() {
      const all = host.tables?.() ?? [];
      return this._sheet ? all.filter(t => t.sheet === this._sheet()) : all;
    }
    add(address, hasHeaders) {
      let made = null;
      let target;
      if (typeof address === 'string') {
        const text = address.replace(/^=/, '');
        const bang = text.lastIndexOf('!');
        const sheet = bang < 0
          ? () => (this._sheet ? this._sheet() : host.activeSheet())
          : () => {
            const name = text.slice(0, bang).replace(/^'|'$/g, '');
            const i = host.sheetNames().findIndex(n => n.toLowerCase() === name.toLowerCase());
            if (i < 0) throw notFound();
            return i;
          };
        target = { sheet, area: () => parseArea((bang < 0 ? text : text.slice(bang + 1)).replace(/\$/g, '')) };
      } else {
        target = { sheet: () => address._sheet(), area: () => address._area() };
      }
      this._ctx._writes.push({ kind: 'table', sheet: target.sheet, area: target.area, hasHeaders: !!hasHeaders, done: name => { made = name; } });
      return new Table(this._ctx, () => {
        if (made === null) throw new OfficeExtensionError('InvalidOperation', 'Call context.sync() first.');
        return made;
      });
    }
    getItem(name) { return new Table(this._ctx, () => name); }
    getItemOrNullObject(name) {
      const table = this.getItem(name);
      Object.defineProperty(table, 'isNullObject', { get: () => !table._find(false) });
      return table;
    }
    getCount() {
      const result = { value: undefined };
      this._ctx._loads.push(() => { result.value = this._list().length; });
      return result;
    }
    _defaultProps() { return ['items']; }
    _fill(name) {
      if (name === 'items' || name.startsWith('items')) {
        this._loaded.set('items', this._list().map(t => {
          const table = new Table(this._ctx, () => t.name);
          for (const p of table._defaultProps()) table._fill(p);
          return table;
        }));
      }
    }
    get items() { return this._get('items'); }
  }

  class Table extends Proxy {
    constructor(context, name) {
      super(context);
      Object.defineProperty(this, '_name', { value: name });
      this.rows = new TableRowCollection(context, this);
    }
    _find(strict = true) {
      let wanted;
      try { wanted = this._name(); } catch (e) { if (strict) throw e; return null; }
      const found = (host.tables?.() ?? []).find(t => t.name.toLowerCase() === String(wanted).toLowerCase());
      if (!found && strict) throw notFound();
      return found;
    }
    _defaultProps() { return ['name', 'showHeaders', 'showTotals', 'style', 'id']; }
    _fill(prop) {
      const t = this._find();
      if (prop === 'name') this._loaded.set('name', t.name);
      else if (prop === 'id') this._loaded.set('id', t.name);
      else if (prop === 'showHeaders') this._loaded.set('showHeaders', t.headerRows > 0);
      else if (prop === 'showTotals') this._loaded.set('showTotals', t.totalsRows > 0);
      else if (prop === 'style') this._loaded.set('style', t.style ?? 'TableStyleMedium2');
    }
    get name() { return this._get('name'); }
    get id() { return this._get('id'); }
    get showHeaders() { return this._get('showHeaders'); }
    get showTotals() { return this._get('showTotals'); }
    get style() { return this._get('style'); }
    _part(top, bottom) {
      return new Range(this._ctx, () => this._find().sheet, () => {
        const t = this._find();
        const a = t.area;
        const head = top ? 0 : t.headerRows;
        const foot = bottom ? 0 : t.totalsRows;
        return { row: a.row + head, col: a.col, rows: Math.max(1, a.rows - head - foot), cols: a.cols };
      });
    }
    getRange() { return this._part(true, true); }
    getDataBodyRange() { return this._part(false, false); }
    getHeaderRowRange() {
      return new Range(this._ctx, () => this._find().sheet, () => {
        const a = this._find().area;
        return { row: a.row, col: a.col, rows: 1, cols: a.cols };
      });
    }
    getTotalRowRange() {
      return new Range(this._ctx, () => this._find().sheet, () => {
        const a = this._find().area;
        return { row: a.row + a.rows - 1, col: a.col, rows: 1, cols: a.cols };
      });
    }
  }

  class TableRowCollection extends Proxy {
    constructor(context, table) {
      super(context);
      Object.defineProperty(this, '_table', { value: table });
    }
    // add(index, values): rows go on at the end (index null or past the last
    // row, as most add-ins call it) and the table grows to take them.
    add(index, values) {
      const rows = Array.isArray(values?.[0]) ? values : [values];
      this._ctx._writes.push({ kind: 'tableRows', name: () => this._table._name(), values: rows });
    }
    getCount() {
      const result = { value: undefined };
      this._ctx._loads.push(() => {
        const t = this._table._find();
        result.value = t.area.rows - t.headerRows - t.totalsRows;
      });
      return result;
    }
  }

  // ---- workbook and sheets -------------------------------------------------
  class Workbook extends Proxy {
    constructor(context) {
      super(context);
      this.worksheets = new WorksheetCollection(context);
      this.names = new NamedItemCollection(context);
      this.tables = new TableCollection(context, null);
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
      this.tables = new TableCollection(context, index);
      this.onChanged = new EventSource(context, 'changed', index);
      this.onSelectionChanged = new EventSource(context, 'selectionChanged', index);
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
    _defaultProps() { return ['horizontalAlignment', 'verticalAlignment', 'wrapText']; }
    _fill(prop) {
      const key = { horizontalAlignment: 'hAlign', verticalAlignment: 'vAlign', wrapText: 'wrap', columnWidth: 'columnWidth', rowHeight: 'rowHeight' }[prop];
      if (!key) return;
      const fallback = { horizontalAlignment: 'General', verticalAlignment: 'Bottom', wrapText: false, columnWidth: 64, rowHeight: 15 }[prop];
      const flat = this._range._cells(c => c?.[key] ?? null).flat();
      this._loaded.set(prop, flat.every(v => v === flat[0]) ? (flat[0] ?? fallback) : null);
    }
    // Alignment as Excel's JavaScript API names it: General, Left, Center,
    // Right, Fill, Justify, CenterAcrossSelection, Distributed; Top, Center,
    // Bottom, Justify, Distributed.
    get horizontalAlignment() { return this._get('horizontalAlignment'); }
    set horizontalAlignment(v) { this._range._queueGrid(v, hAlign => ({ hAlign })); }
    get verticalAlignment() { return this._get('verticalAlignment'); }
    set verticalAlignment(v) { this._range._queueGrid(v, vAlign => ({ vAlign })); }
    get wrapText() { return this._get('wrapText'); }
    set wrapText(v) { this._range._queueGrid(v, wrap => ({ wrap })); }
    // Widths and heights in points, as the API states them.
    get columnWidth() { return this._get('columnWidth'); }
    set columnWidth(v) { this._range._queueGrid(v, columnWidth => ({ columnWidth })); }
    get rowHeight() { return this._get('rowHeight'); }
    set rowHeight(v) { this._range._queueGrid(v, rowHeight => ({ rowHeight })); }
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
    set name(v) { this._range._queueGrid(v, fontName => ({ fontName })); }
    get size() { return this._get('size'); }
    set size(v) { this._range._queueGrid(v, fontSize => ({ fontSize })); }
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
    const listeners = { changed: new Set(), selectionChanged: new Set() };
    root.addEventListener('message', event => {
      const data = event.data;
      // The workbook tells of the user's edits and selection moves.
      if (data && data.oxiOffice === 'event' && listeners[data.type]) {
        for (const listen of listeners[data.type]) {
          try { listen(data); } catch (e) { console.error(e); }
        }
        return;
      }
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
      names: () => mirror.definedNames ?? [],
      tables: () => mirror.tables ?? [],
      // Names, tables and the rest go over with the next sync; a table's name
      // is the one the workbook will give it (Table1, Table2, ...).
      op: w => {
        if (w.kind === 'table') {
          const taken = new Set((mirror.tables ?? []).map(t => t.name.toLowerCase()));
          let n = 1;
          while (taken.has(`table${n}`)) n++;
          const name = `Table${n}`;
          pending.push({ op: 'op', w: { ...w, name } });
          mirror.tables = [...(mirror.tables ?? []), { name, sheet: w.sheet, area: w.area, headerRows: w.hasHeaders ? 1 : 0, totalsRows: 0 }];
          return name;
        }
        pending.push({ op: 'op', w });
        return undefined;
      },
      on: (type, listen) => {
        listeners[type]?.add(listen);
        if (listeners[type]?.size === 1) target.postMessage({ oxiOffice: 'request', op: 'listen', type }, '*');
        return () => listeners[type]?.delete(listen);
      },
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
        else if (typeof host.op === 'function') {
          const answer = host.op(settle(w));
          if (w.done) w.done(answer);
        }
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
