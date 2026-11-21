// Builds tests/fixtures/xlsx_addin/TaxTools.xlsm with Oxi's own VBA project writer
// (the .xlam is that file with the add-in content type; Excel opens it as an
// add-in and its AddTax / StampTotals give 110, 216, 330 and C2 = 55).
// Run: node tools/metrics/_xlsx_make_xlam.mjs [wasm pkg dir]
import { readFileSync, writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
const root = join(dirname(fileURLToPath(import.meta.url)), '..', '..');
const pkg = (process.argv[2] ?? join(root, 'crates/oxidocs-wasm/pkg')) + '/';
const mod = await import('file:///' + pkg + 'oxidocs_wasm.js');
mod.initSync({ module: readFileSync(pkg + 'oxidocs_wasm_bg.wasm') });
const src = `Attribute VB_Name = "TaxTools"
Public Function AddTax(amount As Double, Optional rate As Double = 0.1) As Double
    AddTax = amount * (1 + rate)
End Function

Public Sub StampTotals()
    ActiveSheet.Range("C1").Value = "Stamped by add-in"
    ActiveSheet.Range("C2").Value = AddTax(ActiveSheet.Range("B2").Value)
End Sub
`;
const base = readFileSync(join(root, 'tests/fixtures/vba/cell_width.xlsm'));
try {
  const out = mod.save_spreadsheet_vba_source(base, src, null);
  writeFileSync(join(root, 'tests/fixtures/xlsx_addin/TaxTools.xlsm'), out);
  console.log('written', out.length, 'source back:', JSON.stringify((mod.spreadsheet_vba_source(out, null) || '').slice(0, 60)));
} catch (e) { console.log('failed:', e.message ?? e); }
