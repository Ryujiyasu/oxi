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
      for (const k of ['numberFormat', 'fill', 'fontColor', 'bold', 'italic']) if (k in e) cell[k] = e[k];
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
});

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

console.log(`${passed} passed`);
