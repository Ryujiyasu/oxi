// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! oxi-office.js against the way add-ins use office.js: queued writes carried
//! by context.sync(), load()-then-read, PropertyNotLoaded for reading early,
//! sheets listed and added, formats, and the size check on a written array.
//! The host is a Map, so what is tested is the API's own bookkeeping.
//! Run with `node web/oxi-office.test.mjs`.

import assert from 'node:assert/strict';
import './oxi-office.js';

const { Excel, Office, OxiOffice } = globalThis;
const sheets = [{ name: 'Sheet1', cells: new Map() }, { name: 'Data', cells: new Map() }];
let active = 0;
const key = (r, c) => `${r},${c}`;
OxiOffice.attach({
  sheetNames: () => sheets.map(s => s.name),
  activeSheet: () => active,
  activate: i => { active = i; },
  addSheet: name => sheets.push({ name, cells: new Map() }) - 1,
  read: (s, r, c) => sheets[s].cells.get(key(r, c)) ?? { value: '' },
  write: edits => {
    for (const e of edits) {
      const cells = sheets[e.sheet].cells;
      const cell = { value: '', ...(cells.get(key(e.row, e.col)) ?? {}) };
      if (e.clear) { cells.delete(key(e.row, e.col)); continue; }
      if ('value' in e) { cell.value = e.value; delete cell.formula; }
      if ('formula' in e) { cell.formula = e.formula; cell.value = 0; }
      for (const k of ['numberFormat', 'fill', 'fontColor', 'bold', 'italic', 'fontName', 'fontSize', 'hAlign', 'vAlign', 'wrap', 'columnWidth', 'rowHeight']) if (k in e) cell[k] = e[k];
      cells.set(key(e.row, e.col), cell);
    }
  },
  usedRange: s => {
    const keys = [...sheets[s].cells.keys()].map(k => k.split(',').map(Number));
    if (!keys.length) return { row: 0, col: 0, rows: 1, cols: 1 };
    const rs = keys.map(k => k[0]), cs = keys.map(k => k[1]);
    return { row: Math.min(...rs), col: Math.min(...cs), rows: Math.max(...rs) - Math.min(...rs) + 1, cols: Math.max(...cs) - Math.min(...cs) + 1 };
  },
  selection: () => ({ sheet: active, row: 1, col: 1, rows: 1, cols: 1 }),
  names: () => definedNames,
  tables: () => tables,
  op: w => {
    if (w.kind === 'name') definedNames.push({ name: w.name, formula: w.formula });
    else if (w.kind === 'deleteName') definedNames = definedNames.filter(n => n.name !== w.name);
    else if (w.kind === 'table') {
      const name = `Table${tables.length + 1}`;
      tables.push({ name, sheet: w.sheet, area: w.area, headerRows: w.hasHeaders ? 1 : 0, totalsRows: 0 });
      return name;
    } else if (w.kind === 'tableRows') {
      const t = tables.find(x => x.name === w.name);
      for (const row of w.values) {
        row.forEach((v, j) => sheets[t.sheet].cells.set(key(t.area.row + t.area.rows, t.area.col + j), { value: v }));
        t.area = { ...t.area, rows: t.area.rows + 1 };
      }
    }
    return undefined;
  },
  on: (type, listen) => { listeners.push([type, listen]); return () => { listeners = listeners.filter(l => l[1] !== listen); }; },
});
let definedNames = [];
const tables = [];
let listeners = [];

let passed = 0;
async function check(name, fn) { await fn(); passed++; console.log('ok -', name); }

await check('Office.onReady answers Excel', async () => {
  const info = await Office.onReady();
  assert.equal(info.host, 'Excel');
});

await check('values written in one sync come back after load', async () => {
  await Excel.run(async ctx => {
    const range = ctx.workbook.worksheets.getActiveWorksheet().getRange('A1:B2');
    range.values = [[1, 'two'], [true, '=A1*2']];
    await ctx.sync();
    range.load('values, formulas, address, rowCount');
    await ctx.sync();
    assert.deepEqual(range.formulas, [[1, 'two'], [true, '=A1*2']]);
    assert.equal(range.address, 'Sheet1!A1:B2');
    assert.equal(range.rowCount, 2);
  });
});

await check('reading before load throws PropertyNotLoaded', async () => {
  await Excel.run(async ctx => {
    const range = ctx.workbook.worksheets.getActiveWorksheet().getRange('A1');
    assert.throws(() => range.values, e => e.code === 'PropertyNotLoaded');
    range.load('values');
    assert.throws(() => range.values, e => e.code === 'PropertyNotLoaded');
    await ctx.sync();
    assert.deepEqual(range.values, [[1]]);
  });
});

await check('a written array must match the range', async () => {
  await Excel.run(async ctx => {
    const range = ctx.workbook.worksheets.getActiveWorksheet().getRange('A1:B2');
    assert.throws(() => { range.values = [[1, 2, 3]]; }, e => e.code === 'InvalidArgument');
  });
});

await check('formats are written and read back', async () => {
  await Excel.run(async ctx => {
    const range = ctx.workbook.worksheets.getItem('Sheet1').getRange('A1:B1');
    range.format.fill.color = '#FFFF00';
    range.format.font.bold = true;
    await ctx.sync();
    range.format.fill.load('color');
    range.format.font.load('bold');
    await ctx.sync();
    assert.equal(range.format.fill.color, '#FFFF00');
    assert.equal(range.format.font.bold, true);
  });
});

await check('sheets are listed and added', async () => {
  await Excel.run(async ctx => {
    const added = ctx.workbook.worksheets.add('Report');
    await ctx.sync();
    added.getRange('C3').values = [[42]];
    const list = ctx.workbook.worksheets;
    list.load('items/name');
    await ctx.sync();
    assert.deepEqual(list.items.map(s => s.name), ['Sheet1', 'Data', 'Report']);
    const back = ctx.workbook.worksheets.getItem('report').getRange('C3');
    back.load('values');
    await ctx.sync();
    assert.deepEqual(back.values, [[42]]);
  });
});

await check('getUsedRange spans what was written', async () => {
  await Excel.run(async ctx => {
    const used = ctx.workbook.worksheets.getItem('Sheet1').getUsedRange();
    used.load('address');
    await ctx.sync();
    assert.equal(used.address, 'Sheet1!A1:B2');
  });
});

await check('missing sheets fail with ItemNotFound', async () => {
  await assert.rejects(Excel.run(async ctx => {
    const r = ctx.workbook.worksheets.getItem('Nope').getRange('A1');
    r.load('values');
    await ctx.sync();
  }), e => e.code === 'ItemNotFound');
});


await check('alignment, wrapping, font name and size are written and read back', async () => {
  await Excel.run(async ctx => {
    const range = ctx.workbook.worksheets.getItem('Data').getRange('C3:D3');
    range.format.horizontalAlignment = 'Center';
    range.format.wrapText = true;
    range.format.font.name = 'Arial';
    range.format.font.size = 14;
    range.format.rowHeight = 30;
    await ctx.sync();
    range.format.load('horizontalAlignment, wrapText, rowHeight');
    range.format.font.load('name, size');
    await ctx.sync();
    assert.equal(range.format.horizontalAlignment, 'Center');
    assert.equal(range.format.wrapText, true);
    assert.equal(range.format.rowHeight, 30);
    assert.equal(range.format.font.name, 'Arial');
    assert.equal(range.format.font.size, 14);
  });
});

await check('names are added, listed and resolve to their range', async () => {
  await Excel.run(async ctx => {
    const sheet = ctx.workbook.worksheets.getItem('Sheet1');
    ctx.workbook.names.add('Rate', sheet.getRange('A1'));
    ctx.workbook.names.add('Label', '="hello"');
    await ctx.sync();
    const names = ctx.workbook.names.load('items');
    const rate = ctx.workbook.names.getItem('rate').getRange().load('values, address');
    await ctx.sync();
    assert.deepEqual(names.items.map(n => [n.name, n.formula, n.type]), [['Rate', '=Sheet1!$A$1', 'Range'], ['Label', '="hello"', 'String']]);
    assert.equal(rate.address, 'Sheet1!A1');
    assert.deepEqual(rate.values, [[1]]);
    const missing = ctx.workbook.names.getItemOrNullObject('nope');
    await ctx.sync();
    assert.equal(missing.isNullObject, true);
  });
});

await check('a table is made from a range, its parts read, and rows added', async () => {
  await Excel.run(async ctx => {
    const sheet = ctx.workbook.worksheets.getItem('Data');
    sheet.getRange('A1:B3').values = [['Item', 'Qty'], ['Pen', 2], ['Ink', 5]];
    const table = sheet.tables.add('A1:B3', true);
    table.load('name');
    await ctx.sync();
    assert.equal(table.name, 'Table1');
    table.rows.add(null, [['Pad', 7]]);
    await ctx.sync();
    const body = table.getDataBodyRange().load('values, address');
    const head = table.getHeaderRowRange().load('values');
    const count = table.rows.getCount();
    const all = ctx.workbook.tables.load('items');
    await ctx.sync();
    assert.deepEqual(head.values, [['Item', 'Qty']]);
    assert.deepEqual(body.values, [['Pen', 2], ['Ink', 5], ['Pad', 7]]);
    assert.equal(body.address, 'Data!A2:B4');
    assert.equal(count.value, 3);
    assert.deepEqual(all.items.map(t => t.name), ['Table1']);
  });
});

await check('onChanged and onSelectionChanged hand over Excel-shaped arguments', async () => {
  const seen = [];
  let off;
  await Excel.run(async ctx => {
    const sheet = ctx.workbook.worksheets.getItem('Data');
    off = sheet.onChanged.add(args => { seen.push(args); });
    sheet.onSelectionChanged.add(args => { seen.push(args); });
    await ctx.sync();
  });
  for (const [type, listen] of listeners) {
    listen({ type, sheet: 1, address: 'B2' });
    listen({ type, sheet: 0, address: 'A1' }); // another sheet: not told
  }
  assert.deepEqual(seen.map(a => [a.type, a.address, a.changeType ?? null]),
    [['WorksheetChanged', 'B2', 'RangeEdited'], ['WorksheetSelectionChanged', 'B2', null]]);
  off.remove();
  assert.equal(listeners.length, 1);
});

await check('reading a table before sync throws PropertyNotLoaded', async () => {
  await Excel.run(async ctx => {
    const table = ctx.workbook.tables.getItem('Table1');
    assert.throws(() => table.name, e => e.code === 'PropertyNotLoaded');
  });
});

console.log(`${passed} passed`);
