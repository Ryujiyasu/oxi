// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.
import assert from 'node:assert/strict';
import {readFile} from 'node:fs/promises';
import path from 'node:path';
import {fileURLToPath} from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const pkg = path.resolve(process.argv[2] || path.join(root, 'crates/oxidocs-wasm/pkg'));
const binding = 'data:text/javascript;base64,' + Buffer.from(await readFile(path.join(pkg, 'oxidocs_wasm.js'))).toString('base64');
const wasm = await import(binding);
await wasm.default({module_or_path: await readFile(path.join(pkg, 'oxidocs_wasm_bg.wasm'))});
const original = await readFile(path.join(root, 'tests/fixtures/vba/row_height_state.xlsm'));
const source = `Public Function AddRule(ByVal insertCell As Boolean) As String
Range("A1:A3").Clear
Range("A1").Validation.Add 1, 1, 1, 1, 9
If insertCell Then Range("A2").Insert xlShiftDown, 0
If insertCell Then AddRule = CStr(Range("A2").Validation.Type) Else AddRule = CStr(Range("A1").Validation.Type)
End Function
Public Function ReadRule(ByVal insertCell As Boolean) As String
If insertCell Then ReadRule = CStr(Range("A2").Validation.Type) Else ReadRule = CStr(Range("A1").Validation.Type)
End Function
Public Function AddBandFormat(ByVal axis As Long) As Long
If axis = 0 Then Rows(1).Interior.Color = 255: Rows(2).Insert Else Columns(1).Interior.Color = 255: Columns(2).Insert
AddBandFormat = ReadBandFormat(axis)
End Function
Public Function ReadBandFormat(ByVal axis As Long) As Long
If axis = 0 Then ReadBandFormat = Range("Z2").Interior.Color Else ReadBandFormat = Range("B100").Interior.Color
End Function
Public Function ReadExistingRule() As String
ReadExistingRule = CStr(Range("B1").Validation.Type) & "|" & Range("B1").Validation.Formula1 & "|" & CStr(Range("B1").Validation.IgnoreBlank) & "|" & CStr(Range("B1").Validation.InCellDropdown) & "|" & CStr(Range("B1").Validation.ShowInput) & "|" & CStr(Range("B1").Validation.ShowError)
End Function
Public Function AddNewSheetFormat() As String
Worksheets.Add
Rows(1).Interior.Color = 255: Rows(2).Insert
Columns(1).Interior.Color = 65280: Columns(2).Insert
Range("C10").Validation.Add 1, 1, 1, 1, 9
AddNewSheetFormat = ReadNewSheetFormat()
End Function
Public Function ReadNewSheetFormat() As String
ReadNewSheetFormat = CStr(Range("Z2").Interior.Color) & "|" & CStr(Range("B100").Interior.Color) & "|" & CStr(Range("C10").Validation.Type)
End Function`;
const failures = [];
for (const inserted of [false, true]) {
    const before = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(original), source, 'AddRule', [inserted], 0, 'rows.xlsx', {});
    assert.equal(before.error, undefined, 'before save');
    assert.deepEqual(before.result, {type: 'string', value: '1'}, 'before save');
    const reopened = wasm.parse_spreadsheet(wasm.edit_xlsx_from_workbook(original, JSON.parse(JSON.stringify(before.workbook))));
    const after = wasm.run_spreadsheet_vba_with_outcome(reopened, source, 'ReadRule', [inserted], 0, 'rows.xlsx', {});
    if (after.error || after.result?.type !== 'string' || after.result?.value !== '1') {
        failures.push({case: 'validation', inserted, result: after.result, error: after.error});
    }
}
for (const axis of [0, 1]) {
    const before = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(original), source, 'AddBandFormat', [axis], 0, 'rows.xlsx', {});
    assert.equal(before.error, undefined, 'band before save');
    assert.deepEqual(before.result, {type: 'integer', value: 255}, 'band before save');
    const reopened = wasm.parse_spreadsheet(wasm.edit_xlsx_from_workbook(original, JSON.parse(JSON.stringify(before.workbook))));
    const after = wasm.run_spreadsheet_vba_with_outcome(reopened, source, 'ReadBandFormat', [axis], 0, 'rows.xlsx', {});
    if (after.error || after.result?.type !== 'integer' || after.result?.value !== 255) {
        failures.push({case: 'band_format', axis, result: after.result, error: after.error});
    }
}
// Excel 16.0 build 20527 reads this fixture's existing list validation as
// type 3, with the four flags true and Formula1 equal to a,b,c.
const existingFile = await readFile(path.join(root, 'crates/oxicells-core/tests/fixtures/unsupported_bits.xlsx'));
const existing = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(existingFile), source, 'ReadExistingRule', [], 0, 'existing.xlsx', {});
if (existing.error || existing.result?.type !== 'string' || existing.result?.value !== '3|a,b,c|True|True|True|True') {
    failures.push({case: 'existing_validation', result: existing.result, error: existing.error});
}
const newSheet = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(original), source, 'AddNewSheetFormat', [], 0, 'rows.xlsx', {});
assert.equal(newSheet.error, undefined, 'new sheet before save');
assert.deepEqual(newSheet.result, {type: 'string', value: '255|65280|1'}, 'new sheet before save');
const newSheetReopened = wasm.parse_spreadsheet(wasm.edit_xlsx_from_workbook(original, JSON.parse(JSON.stringify(newSheet.workbook))));
const newSheetRead = wasm.run_spreadsheet_vba_with_outcome(newSheetReopened, source, 'ReadNewSheetFormat', [], 0, 'rows.xlsx', {});
if (newSheetRead.error || newSheetRead.result?.type !== 'string' || newSheetRead.result?.value !== '255|65280|1') {
    failures.push({case: 'new_sheet_format', result: newSheetRead.result, error: newSheetRead.error});
}
// SaveCopyAs runs while the macro is still active, before normal outcome sync.
const group = JSON.parse(await readFile(path.join(root, 'tests/fixtures/vba/save_copy_native_cases.json'), 'utf8')).groups[0];
const copySource = `Public Function SaveBand() As Long
Rows(1).Interior.Color = 255: Rows(2).Insert
ThisWorkbook.SaveCopyAs "${group.cases[0].target}"
SaveBand = Range("Z2").Interior.Color
End Function
Public Function ReadBand() As Long
ReadBand = Range("Z2").Interior.Color
End Function`;
const copy = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(original), copySource, 'SaveBand', [], 0, 'input.xlsm', {
    applicationState: {displayAlerts: false},
    fileContext: {packageBytes: original, fullName: group.full_name, readOnly: group.read_only,
        folders: group.folders, files: group.initial_files},
});
assert.equal(copy.error, undefined, 'SaveCopyAs with row format');
assert.equal(copy.files?.length, 1, 'SaveCopyAs output');
const copyRead = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(new Uint8Array(copy.files[0].bytes)), copySource, 'ReadBand', [], 0, 'copy.xlsm', {});
if (copyRead.error || copyRead.result?.type !== 'integer' || copyRead.result?.value !== 255) {
    failures.push({case: 'save_copy_band_format', result: copyRead.result, error: copyRead.error});
}
// Excel 16.0 build 20527: later row/column formatting wins at the
// intersection while unrelated properties of the other band remain intact.
for (const [name, operation, expected] of [
    ['row_then_col', 'Rows(1).Interior.Color=255: Columns(1).Interior.Color=65280', '65280|255|65280|False|False|False'],
    ['col_then_row', 'Columns(1).Interior.Color=65280: Rows(1).Interior.Color=255', '255|255|65280|False|False|False'],
    ['cell_then_row', 'Range("A1").Interior.Color=16711680: Rows(1).Interior.Color=255', '255|255|16777215|False|False|False'],
    ['row_then_cell', 'Rows(1).Interior.Color=255: Range("A1").Interior.Color=16711680', '16711680|255|16777215|False|False|False'],
    ['row_bold_then_col_color', 'Rows(1).Font.Bold=True: Columns(1).Interior.Color=65280', '65280|16777215|65280|True|True|False'],
    ['col_bold_then_row_color', 'Columns(1).Font.Bold=True: Rows(1).Interior.Color=255', '255|255|16777215|True|False|True'],
]) {
    const crossSource = `Public Function MakeCross() As String
Worksheets.Add
${operation}
MakeCross = ReadCross()
End Function
Public Function ReadCross() As String
ReadCross = CStr(Range("A1").Interior.Color) & "|" & CStr(Range("Z1").Interior.Color) & "|" & CStr(Range("A100").Interior.Color) & "|" & CStr(Range("A1").Font.Bold) & "|" & CStr(Range("Z1").Font.Bold) & "|" & CStr(Range("A100").Font.Bold)
End Function`;
    const before = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(original), crossSource, 'MakeCross', [], 0, 'rows.xlsm', {});
    if (before.error || before.result?.value !== expected) {
        failures.push({case: name, phase: 'runtime', result: before.result, error: before.error});
        continue;
    }
    const saved = wasm.edit_xlsx_from_workbook(original, JSON.parse(JSON.stringify(before.workbook)));
    const after = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), crossSource, 'ReadCross', [], 0, 'rows.xlsm', {});
    if (after.error || after.result?.value !== expected) failures.push({case: name, phase: 'saved', result: after.result, error: after.error});
}
for (const [name, operation, expected] of [
    ['clear_col', 'Rows(1).Interior.Color=255: Columns(1).Interior.Color=65280: Columns(1).ClearFormats', '16777215|255|16777215'],
    ['clear_row', 'Rows(1).Interior.Color=255: Columns(1).Interior.Color=65280: Rows(1).ClearFormats', '16777215|16777215|65280'],
    ['clear_cell', 'Rows(1).Interior.Color=255: Columns(1).Interior.Color=65280: Range("A1").ClearFormats', '16777215|255|65280'],
    ['clear_col_after_row', 'Rows(1).Interior.Color=255: Columns(1).ClearFormats', '16777215|255|16777215'],
    ['clear_row_after_col', 'Columns(1).Interior.Color=65280: Rows(1).ClearFormats', '16777215|16777215|65280'],
    ['clear_cell_after_col', 'Columns(1).Interior.Color=65280: Range("A1").ClearFormats', '16777215|16777215|65280'],
    ['clear_cell_after_row', 'Rows(1).Interior.Color=255: Range("A1").ClearFormats', '16777215|255|16777215'],
]) {
    const clearSource = `Public Function MakeClear() As String
Worksheets.Add
${operation}
MakeClear = ReadClear()
End Function
Public Function ReadClear() As String
ReadClear = CStr(Range("A1").Interior.Color) & "|" & CStr(Range("Z1").Interior.Color) & "|" & CStr(Range("A100").Interior.Color)
End Function`;
    const before = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(original), clearSource, 'MakeClear', [], 0, 'rows.xlsm', {});
    if (before.error || before.result?.value !== expected) {
        failures.push({case: name, phase: 'runtime', result: before.result, error: before.error});
        continue;
    }
    const saved = wasm.edit_xlsx_from_workbook(original, JSON.parse(JSON.stringify(before.workbook)));
    const after = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), clearSource, 'ReadClear', [], 0, 'rows.xlsm', {});
    if (after.error || after.result?.value !== expected) failures.push({case: name, phase: 'saved', result: after.result, error: after.error});
}
// Excel 16.0 build 20527: clearing part of a sheet-wide format must leave
// the rest of the sheet formatted, including after a save/reopen.
for (const [name, operation, expected] of [
    ['sheet_clear_row', 'Rows(1).ClearFormats', '16777215|16777215|255|255'],
    ['sheet_clear_col', 'Columns(1).ClearFormats', '16777215|255|16777215|255'],
    ['sheet_clear_cell', 'Range("A1").ClearFormats', '16777215|255|255|255'],
]) {
    const sheetSource = `Public Function MakeSheetClear() As String
Worksheets.Add
Cells.Interior.Color=255
${operation}
MakeSheetClear = ReadSheetClear()
End Function
Public Function ReadSheetClear() As String
ReadSheetClear = CStr(Range("A1").Interior.Color) & "|" & CStr(Range("Z1").Interior.Color) & "|" & CStr(Range("A100").Interior.Color) & "|" & CStr(Range("Z100").Interior.Color)
End Function`;
    const before = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(original), sheetSource, 'MakeSheetClear', [], 0, 'rows.xlsm', {});
    if (before.error || before.result?.value !== expected) {
        failures.push({case: name, phase: 'runtime', result: before.result, error: before.error});
        continue;
    }
    const saved = wasm.edit_xlsx_from_workbook(original, JSON.parse(JSON.stringify(before.workbook)));
    const after = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), sheetSource, 'ReadSheetClear', [], 0, 'rows.xlsm', {});
    if (after.error || after.result?.value !== expected) failures.push({case: name, phase: 'saved', result: after.result, error: after.error});
}
const sheetCrossSource = `Public Function MakeSheetCross() As String
Worksheets.Add
Cells.Interior.Color=255: Rows(1).Font.Bold=True: Columns(1).Interior.Color=65280
MakeSheetCross = ReadSheetCross()
End Function
Public Function ReadSheetCross() As String
ReadSheetCross = CStr(Range("A1").Interior.Color) & "|" & CStr(Range("A1").Font.Bold) & "|" & CStr(Range("Z1").Interior.Color) & "|" & CStr(Range("Z1").Font.Bold) & "|" & CStr(Range("A100").Interior.Color) & "|" & CStr(Range("A100").Font.Bold) & "|" & CStr(Range("Z100").Interior.Color) & "|" & CStr(Range("Z100").Font.Bold)
End Function
Public Function ChangeSecondRow() As String
Rows(2).Font.Bold=True
ChangeSecondRow = ReadSecondRow()
End Function
Public Function ReadSecondRow() As String
ReadSecondRow = CStr(Range("A2").Interior.Color) & "|" & CStr(Range("A2").Font.Bold) & "|" & CStr(Range("Z2").Interior.Color) & "|" & CStr(Range("Z2").Font.Bold)
End Function`;
const sheetCrossExpected = '65280|True|255|True|65280|False|255|False';
const sheetCrossBefore = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(original), sheetCrossSource, 'MakeSheetCross', [], 0, 'rows.xlsm', {});
if (sheetCrossBefore.error || sheetCrossBefore.result?.value !== sheetCrossExpected) {
    failures.push({case: 'sheet_cross', phase: 'runtime', result: sheetCrossBefore.result, error: sheetCrossBefore.error});
} else {
    const saved = wasm.edit_xlsx_from_workbook(original, JSON.parse(JSON.stringify(sheetCrossBefore.workbook)));
    const reopened = wasm.parse_spreadsheet(saved);
    const after = wasm.run_spreadsheet_vba_with_outcome(reopened, sheetCrossSource, 'ReadSheetCross', [], 0, 'rows.xlsm', {});
    if (after.error || after.result?.value !== sheetCrossExpected) failures.push({case: 'sheet_cross', phase: 'saved', result: after.result, error: after.error});
    const bands = reopened.sheets[0].band_styles;
    if (!bands?.sheet || (bands.columns?.length || 0) !== 1) failures.push({case: 'sheet_cross', phase: 'compact', columns: bands?.columns?.length});
    const changed = wasm.run_spreadsheet_vba_with_outcome(reopened, sheetCrossSource, 'ChangeSecondRow', [], 0, 'rows.xlsm', {});
    const count = changed.workbook.sheets[0].rows.reduce((sum, row) => sum + row.cells.length, 0);
    if (changed.error || changed.result?.value !== '65280|True|255|True' || count > 2) failures.push({case: 'sheet_cross', phase: 'second_row', result: changed.result, cells: count, error: changed.error});
    const savedAgain = wasm.edit_xlsx_from_workbook(saved, JSON.parse(JSON.stringify(changed.workbook)));
    const afterAgain = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(savedAgain), sheetCrossSource, 'ReadSecondRow', [], 0, 'rows.xlsm', {});
    if (afterAgain.error || afterAgain.result?.value !== '65280|True|255|True') failures.push({case: 'sheet_cross', phase: 'second_save', result: afterAgain.result, error: afterAgain.error});
}
// Excel 16.0 build 20527 resolves relative references from the first cell of
// the validation range, including when a range is split by deleting a rule.
for (const [name, deletion, expected] of [
    ['relative_validation', '', '7:=D1+$E$2|7:=D2+$E$2|7:=D3+$E$2'],
    ['relative_validation_delete_first', 'Range("A1").Validation.Delete', 'none|7:=D2+$E$2|7:=D3+$E$2'],
    ['relative_validation_delete_middle', 'Range("A2").Validation.Delete', '7:=D1+$E$2|none|7:=D3+$E$2'],
]) {
    const relativeSource = `Public Function MakeRelative() As String
Worksheets.Add
Range("A1:A3").Validation.Add 7, 1, 1, "=D1+$E$2"
${deletion}
MakeRelative = ReadRelative()
End Function
Public Function ReadRelative() As String
ReadRelative = OneRule("A1") & "|" & OneRule("A2") & "|" & OneRule("A3")
End Function
Public Function ReadSharedRelative() As String
ReadSharedRelative = CStr(Range("A1:A3").Validation.Type) & ":" & Range("A1:A3").Validation.Formula1
End Function
Public Function OneRule(ByVal address As String) As String
On Error Resume Next
Dim kind As Long
kind = Range(address).Validation.Type
If Err.Number <> 0 Then Err.Clear: OneRule = "none": Exit Function
OneRule = CStr(kind) & ":" & Range(address).Validation.Formula1
End Function`;
    const before = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(original), relativeSource, 'MakeRelative', [], 0, 'rows.xlsm', {});
    if (before.error || before.result?.value !== expected) {
        failures.push({case: name, phase: 'runtime', result: before.result, error: before.error});
        continue;
    }
    const saved = wasm.edit_xlsx_from_workbook(original, JSON.parse(JSON.stringify(before.workbook)));
    const after = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), relativeSource, 'ReadRelative', [], 0, 'rows.xlsm', {});
    if (after.error || after.result?.value !== expected) failures.push({case: name, phase: 'saved', result: after.result, error: after.error});
    if (!deletion) {
        const shared = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), relativeSource, 'ReadSharedRelative', [], 0, 'rows.xlsm', {});
        if (shared.error || shared.result?.value !== '7:=D1+$E$2') failures.push({case: name, phase: 'shared_range', result: shared.result, error: shared.error});
    }
}
for (const [name, operation, expected] of [
    ['validation_insert_cell_down', 'Range("A2").Insert xlShiftDown', '7:=D1+$E$2|7:=D2+$E$2|7:=D2+$E$2|7:=D3+$E$2|none'],
    ['validation_insert_row', 'Rows(2).Insert', '7:=D1+$E$3|7:=D2+$E$3|7:=D3+$E$3|7:=D4+$E$3|none'],
    ['validation_insert_cell_right', 'Range("A2").Insert xlShiftToRight', '7:=D1+$F$2|none|7:=D3+$F$2|none|7:=E2+$F$2'],
    ['validation_delete_row', 'Rows(2).Delete', '7:=D1+#REF!|7:=D2+#REF!|none|none|none'],
]) {
    const shiftSource = `Public Function MakeShift() As String
Worksheets.Add
Range("A1:A3").Validation.Add 7, 1, 1, "=D1+$E$2"
${operation}
MakeShift = ReadShift()
End Function
Public Function ReadShift() As String
ReadShift = OneRule("A1") & "|" & OneRule("A2") & "|" & OneRule("A3") & "|" & OneRule("A4") & "|" & OneRule("B2")
End Function
Public Function OneRule(ByVal address As String) As String
On Error Resume Next
Dim kind As Long
kind = Range(address).Validation.Type
If Err.Number <> 0 Then Err.Clear: OneRule = "none": Exit Function
OneRule = CStr(kind) & ":" & Range(address).Validation.Formula1
End Function`;
    const before = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(original), shiftSource, 'MakeShift', [], 0, 'rows.xlsm', {});
    if (before.error || before.result?.value !== expected) {
        failures.push({case: name, phase: 'runtime', result: before.result, error: before.error});
        continue;
    }
    const saved = wasm.edit_xlsx_from_workbook(original, JSON.parse(JSON.stringify(before.workbook)));
    const after = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), shiftSource, 'ReadShift', [], 0, 'rows.xlsm', {});
    if (after.error || after.result?.value !== expected) failures.push({case: name, phase: 'saved', result: after.result, error: after.error});
}
// Excel 16.0 build 20527: alternate copy origins, the first row/column,
// and deleting a cell preserve different validation-reference positions.
for (const [name, operation, expected] of [
    ['cell_down_after', 'Range("A2").Insert xlShiftDown,1', '=D1+$E$2|=D1+$E$2|=D2+$E$2|=D3+$E$2|none|none'],
    ['row_down_after', 'Rows(2).Insert xlShiftDown,1', '=D1+$E$3|=D2+$E$3|=D3+$E$3|=D4+$E$3|none|none'],
    ['cell_top', 'Range("A1").Insert xlShiftDown', 'none|=D1+$E$2|=D2+$E$2|=D3+$E$2|none|none'],
    ['row_top', 'Rows(1).Insert', 'none|=D2+$E$3|=D3+$E$3|=D4+$E$3|none|none'],
    ['cell_delete', 'Range("A2").Delete xlShiftUp', '=D1+$E$2|=D3+$E$2|none|none|none|none'],
    ['col_insert', 'Columns(1).Insert', 'none|none|none|none|=E1+$F$2|=E2+$F$2'],
]) {
    const boundarySource = `Public Function MakeBoundary() As String
Worksheets.Add
Range("A1:A3").Validation.Add 7, 1, 1, "=D1+$E$2"
${operation}
MakeBoundary = ReadBoundary()
End Function
Public Function ReadBoundary() As String
ReadBoundary = OneRule("A1") & "|" & OneRule("A2") & "|" & OneRule("A3") & "|" & OneRule("A4") & "|" & OneRule("B1") & "|" & OneRule("B2")
End Function
Public Function OneRule(ByVal address As String) As String
On Error Resume Next
Dim kind As Long
kind = Range(address).Validation.Type
If Err.Number <> 0 Then Err.Clear: OneRule = "none": Exit Function
OneRule = Range(address).Validation.Formula1
End Function`;
    const before = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(original), boundarySource, 'MakeBoundary', [], 0, 'rows.xlsm', {});
    if (before.error || before.result?.value !== expected) {
        failures.push({case: name, phase: 'runtime', result: before.result, error: before.error});
        continue;
    }
    const saved = wasm.edit_xlsx_from_workbook(original, JSON.parse(JSON.stringify(before.workbook)));
    const after = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), boundarySource, 'ReadBoundary', [], 0, 'rows.xlsm', {});
    if (after.error || after.result?.value !== expected) failures.push({case: name, phase: 'saved', result: after.result, error: after.error});
}
// Both bounds of a numeric validation have their own relative references.
for (const [name, operation, expected] of [
    ['bounds_none', '', '=D1,=E1|=D2,=E2|=D3,=E3|none'],
    ['bounds_cell_down', 'Range("A2").Insert xlShiftDown', '=D1,=E1|=D2,=E2|=D2,=E2|=D3,=E3'],
    ['bounds_row_down', 'Rows(2).Insert', '=D1,=E1|=D2,=E2|=D3,=E3|=D4,=E4'],
    ['bounds_cell_delete', 'Range("A2").Delete xlShiftUp', '=D1,=E1|=D3,=E3|none|none'],
]) {
    const boundsSource = `Public Function MakeBounds() As String
Worksheets.Add
Range("A1:A3").Validation.Add 1, 1, 1, "=D1", "=E1"
${operation}
MakeBounds = ReadBounds()
End Function
Public Function ReadBounds() As String
ReadBounds = OneRule("A1") & "|" & OneRule("A2") & "|" & OneRule("A3") & "|" & OneRule("A4")
End Function
Public Function OneRule(ByVal address As String) As String
On Error Resume Next
Dim kind As Long
kind = Range(address).Validation.Type
If Err.Number <> 0 Then Err.Clear: OneRule = "none": Exit Function
OneRule = Range(address).Validation.Formula1 & "," & Range(address).Validation.Formula2
End Function`;
    const before = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(original), boundsSource, 'MakeBounds', [], 0, 'rows.xlsm', {});
    if (before.error || before.result?.value !== expected) {
        failures.push({case: name, phase: 'runtime', result: before.result, error: before.error});
        continue;
    }
    const saved = wasm.edit_xlsx_from_workbook(original, JSON.parse(JSON.stringify(before.workbook)));
    const after = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), boundsSource, 'ReadBounds', [], 0, 'rows.xlsm', {});
    if (after.error || after.result?.value !== expected) failures.push({case: name, phase: 'saved', result: after.result, error: after.error});
}
const validationValueSource = `Public Function MakeValues() As String
Worksheets.Add
Range("D1").Value2=10: Range("D2").Value2=20: Range("D3").Value2=30
Range("A1").Value2=15: Range("A2").Value2=15: Range("A3").Value2=35
Range("A1:A3").Validation.Add 1,1,5,"=D1"
MakeValues=ReadValues()
End Function
Public Function ReadValues() As String
ReadValues=CStr(Range("A1").Validation.Value) & "|" & CStr(Range("A2").Validation.Value) & "|" & CStr(Range("A3").Validation.Value) & "|" & CStr(Range("A1:A3").Validation.Value)
End Function`;
const valueBefore = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(original), validationValueSource, 'MakeValues', [], 0, 'rows.xlsm', {});
if (valueBefore.error || valueBefore.result?.value !== 'True|False|True|True') {
    failures.push({case: 'relative_validation_value', phase: 'runtime', result: valueBefore.result, error: valueBefore.error});
} else {
    const saved = wasm.edit_xlsx_from_workbook(original, JSON.parse(JSON.stringify(valueBefore.workbook)));
    const after = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), validationValueSource, 'ReadValues', [], 0, 'rows.xlsm', {});
    if (after.error || after.result?.value !== 'True|False|True|True') failures.push({case: 'relative_validation_value', phase: 'saved', result: after.result, error: after.error});
}
const crossSheetSource = `Public Function MakeCrossSheet() As String
Worksheets.Add: ActiveSheet.Name="Source"
Worksheets.Add: ActiveSheet.Name="Target"
Range("A1:A3").Validation.Add 7,1,1,"=Source!D1"
Worksheets("Source").Rows(1).Insert
MakeCrossSheet=ReadCrossSheet()
End Function
Public Function ReadCrossSheet() As String
ReadCrossSheet=Range("A1").Validation.Formula1 & "|" & Range("A2").Validation.Formula1
End Function`;
const crossSheetBefore = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(original), crossSheetSource, 'MakeCrossSheet', [], 0, 'rows.xlsm', {});
if (crossSheetBefore.error || crossSheetBefore.result?.value !== '=Source!D2|=Source!D3') {
    failures.push({case: 'cross_sheet_validation', phase: 'runtime', result: crossSheetBefore.result, error: crossSheetBefore.error});
} else {
    const saved = wasm.edit_xlsx_from_workbook(original, JSON.parse(JSON.stringify(crossSheetBefore.workbook)));
    const after = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), crossSheetSource, 'ReadCrossSheet', [], 0, 'rows.xlsm', {});
    if (after.error || after.result?.value !== '=Source!D2|=Source!D3') failures.push({case: 'cross_sheet_validation', phase: 'saved', result: after.result, error: after.error});
}
const multiValidationSource = `Public Function MakeMulti() As String
Worksheets.Add
Range("A1:A3,C2:C4").Validation.Add Type:=7, AlertStyle:=1, Operator:=1, Formula1:="=D1"
Range("A1:A3,C2:C4").Validation.IgnoreBlank=False
MakeMulti=ReadMulti()
End Function
Public Function ReadMulti() As String
ReadMulti=CStr(Range("A1:A3,C2:C4").Validation.Type) & "|" & Range("A1:A3,C2:C4").Validation.Formula1 & "|" & OneMulti("A1") & "|" & OneMulti("A3") & "|" & OneMulti("C2") & "|" & OneMulti("C4") & "|" & CStr(Range("A1").Validation.IgnoreBlank) & "|" & CStr(Range("C2").Validation.IgnoreBlank) & "|" & Range("A1:A3,C2:C4").Validation.Parent.Address
End Function
Public Function ModifyMulti() As String
Range("A1:A3,C2:C4").Validation.Modify Type:=7, AlertStyle:=1, Operator:=1, Formula1:="=E1"
ModifyMulti=ReadMulti()
End Function
Public Function DeleteMulti() As String
Range("A1:A3,C2:C4").Validation.Delete
DeleteMulti=OneMulti("A1") & "|" & OneMulti("C2")
End Function
Public Function MakeReverse() As String
Worksheets.Add
Range("C2:C4,A1:A3").Validation.Add 7,1,1,"=D3"
MakeReverse=OneMulti("C2") & "|" & OneMulti("C4") & "|" & OneMulti("A1") & "|" & OneMulti("A3")
End Function
Public Function ReadReverse() As String
ReadReverse=OneMulti("C2") & "|" & OneMulti("C4") & "|" & OneMulti("A1") & "|" & OneMulti("A3")
End Function
Public Function MakeOverlap() As String
Worksheets.Add
Range("A1:A3,A2:A4").Validation.Add 7,1,1,"=D1"
MakeOverlap=OneMulti("A1") & "|" & OneMulti("A2") & "|" & OneMulti("A3") & "|" & OneMulti("A4")
End Function
Public Function ReadOverlap() As String
ReadOverlap=OneMulti("A1") & "|" & OneMulti("A2") & "|" & OneMulti("A3") & "|" & OneMulti("A4")
End Function
Public Function MakeAtomic() As String
Worksheets.Add
Range("C2").Validation.Add 7,1,1,"=H1"
On Error Resume Next
Range("A1:A3,C2:C4").Validation.Add 7,1,1,"=D1"
Dim failed As Boolean
failed=(Err.Number<>0): Err.Clear
MakeAtomic=CStr(failed) & "|" & OneMulti("A1") & "|" & OneMulti("C2")
End Function
Public Function ReadAtomic() As String
ReadAtomic=OneMulti("A1") & "|" & OneMulti("C2")
End Function
Public Function MakeModifyAtomic() As String
Worksheets.Add
Range("A1").Validation.Add 7,1,1,"=D1"
On Error Resume Next
Range("A1,C2").Validation.Modify 7,1,1,"=E1"
Dim failed As Boolean
failed=(Err.Number<>0): Err.Clear
MakeModifyAtomic=CStr(failed) & "|" & OneMulti("A1") & "|" & OneMulti("C2")
End Function
Public Function MultiValue() As String
Worksheets.Add
Range("A1,C2").Validation.Add 1,1,5,10
Range("A1").Value2=5: Range("C2").Value2=15
MultiValue=CStr(Range("A1,C2").Validation.Value)
End Function
Public Function OneMulti(ByVal address As String) As String
On Error Resume Next
Dim kind As Long
kind=Range(address).Validation.Type
If Err.Number<>0 Then Err.Clear: OneMulti="none": Exit Function
OneMulti=Range(address).Validation.Formula1
End Function`;
const checkMulti = (outcome, expected, label) => {
    if (outcome.error || outcome.result?.value !== expected) failures.push({case: label, result: outcome.result, error: outcome.error});
};
const multi = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(original), multiValidationSource, 'MakeMulti', [], 0, 'rows.xlsm', {});
checkMulti(multi, '7|=D1|=D1|=D3|=F2|=F4|False|False|$A$1:$A$3,$C$2:$C$4', 'multi_validation_add');
if (!multi.error) {
    let saved = wasm.edit_xlsx_from_workbook(original, JSON.parse(JSON.stringify(multi.workbook)));
    let reopened = wasm.parse_spreadsheet(saved);
    checkMulti(wasm.run_spreadsheet_vba_with_outcome(reopened, multiValidationSource, 'ReadMulti', [], 0, 'rows.xlsm', {}), '7|=D1|=D1|=D3|=F2|=F4|False|False|$A$1:$A$3,$C$2:$C$4', 'multi_validation_saved');
    const modified = wasm.run_spreadsheet_vba_with_outcome(reopened, multiValidationSource, 'ModifyMulti', [], 0, 'rows.xlsm', {});
    checkMulti(modified, '7|=E1|=E1|=E3|=G2|=G4|False|False|$A$1:$A$3,$C$2:$C$4', 'multi_validation_modify');
    if (!modified.error) {
        saved = wasm.edit_xlsx_from_workbook(saved, JSON.parse(JSON.stringify(modified.workbook)));
        reopened = wasm.parse_spreadsheet(saved);
        checkMulti(wasm.run_spreadsheet_vba_with_outcome(reopened, multiValidationSource, 'ReadMulti', [], 0, 'rows.xlsm', {}), '7|=E1|=E1|=E3|=G2|=G4|False|False|$A$1:$A$3,$C$2:$C$4', 'multi_validation_modify_saved');
        const deleted = wasm.run_spreadsheet_vba_with_outcome(reopened, multiValidationSource, 'DeleteMulti', [], 0, 'rows.xlsm', {});
        checkMulti(deleted, 'none|none', 'multi_validation_delete');
        if (!deleted.error) {
            saved = wasm.edit_xlsx_from_workbook(saved, JSON.parse(JSON.stringify(deleted.workbook)));
            const readDeleted = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), multiValidationSource, 'DeleteMulti', [], 0, 'rows.xlsm', {});
            checkMulti(readDeleted, 'none|none', 'multi_validation_delete_saved');
        }
    }
}
for (const [procedure, expected, reader, savedExpected] of [['MakeReverse', '=D3|=D5|=B2|=B4', 'ReadReverse', '=D3|=D5|=B2|=B4'], ['MakeOverlap', '=D1|none|none|=D4', 'ReadOverlap', '=D1|none|none|=D4'], ['MakeAtomic', 'True|none|=H1', 'ReadAtomic', 'none|=H1'], ['MakeModifyAtomic', 'True|=D1|none', 'ReadAtomic', '=D1|none']]) {
    const result = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(original), multiValidationSource, procedure, [], 0, 'rows.xlsm', {});
    checkMulti(result, expected, procedure);
    if (!result.error) {
        const saved = wasm.edit_xlsx_from_workbook(original, JSON.parse(JSON.stringify(result.workbook)));
        checkMulti(wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), multiValidationSource, reader, [], 0, 'rows.xlsm', {}), savedExpected, procedure + '_saved');
    }
}
checkMulti(wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(original), multiValidationSource, 'MultiValue', [], 0, 'rows.xlsm', {}), 'False', 'multi_validation_value');
// Excel 16.0 build 20527: shifting one or both areas must keep each
// relative reference and preserve ambiguity where overlapping areas remain.
for (const [name, area, operation, expected] of [
    ['multi_row_insert', 'A1:A3,C2:C4', 'Rows(2).Insert', '=D1|=D2|=D3|=D4|none|none|none|=F3|=F4|=F5|none|none'],
    ['multi_cell_insert', 'A1:A3,C2:C4', 'Range("A2").Insert xlShiftDown', '=D1|=D2|=D2|=D3|none|none|=F2|=F3|=F4|none|none|none'],
    ['multi_col_insert', 'A1:A3,C2:C4', 'Columns(2).Insert', '=E1|=E2|=E3|none|none|none|none|none|none|none|=G2|=G3'],
    ['multi_row_delete', 'A1:A3,C2:C4', 'Rows(2).Delete', '=D1|=D2|none|none|none|none|=F2|=F3|none|none|none|none'],
    ['multi_overlap_insert', 'A1:A3,A2:A4', 'Rows(2).Insert', '=D1|=D2|none|none|=D5|none|none|none|none|none|none|none'],
]) {
    const source = `Public Function MakeMultiShift() As String
Worksheets.Add
Range("${area}").Validation.Add 7,1,1,"=D1"
${operation}
MakeMultiShift=ReadMultiShift()
End Function
Public Function ReadMultiShift() As String
ReadMultiShift=OneRule("A1") & "|" & OneRule("A2") & "|" & OneRule("A3") & "|" & OneRule("A4") & "|" & OneRule("A5") & "|" & OneRule("C1") & "|" & OneRule("C2") & "|" & OneRule("C3") & "|" & OneRule("C4") & "|" & OneRule("C5") & "|" & OneRule("D2") & "|" & OneRule("D3")
End Function
Public Function OneRule(ByVal address As String) As String
On Error Resume Next
Dim kind As Long
kind=Range(address).Validation.Type
If Err.Number<>0 Then Err.Clear: OneRule="none": Exit Function
OneRule=Range(address).Validation.Formula1
End Function`;
    const before = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(original), source, 'MakeMultiShift', [], 0, 'rows.xlsm', {});
    if (before.error || before.result?.value !== expected) {
        failures.push({case: name, phase: 'runtime', result: before.result, error: before.error});
        continue;
    }
    const saved = wasm.edit_xlsx_from_workbook(original, JSON.parse(JSON.stringify(before.workbook)));
    const after = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), source, 'ReadMultiShift', [], 0, 'rows.xlsm', {});
    if (after.error || after.result?.value !== expected) failures.push({case: name, phase: 'saved', result: after.result, error: after.error});
}
for (const [name, operation, expected] of [
    ['validation_copy', 'Range("A1:A3").Copy Range("C1")', '=F1+$E$2|=F2+$E$2|=F3+$E$2'],
    ['validation_paste', 'Range("A1:A3").Copy: Range("C1").PasteSpecial 6', '=F1+$E$2|=F2+$E$2|=F3+$E$2'],
    ['validation_paste_transpose', 'Range("A1:A3").Copy: Range("C1").PasteSpecial 6,-4142,False,True', '=F1+$E$2|=G1+$E$2|=H1+$E$2'],
]) {
    const positions = name === 'validation_paste_transpose' ? ['C1', 'D1', 'E1'] : ['C1', 'C2', 'C3'];
    const source = `Public Function MakeCopy() As String
Worksheets.Add
Range("A1:A3").Validation.Add 7,1,1,"=D1+$E$2"
${operation}
MakeCopy=ReadCopy()
End Function
Public Function ReadCopy() As String
ReadCopy=Range("${positions[0]}").Validation.Formula1 & "|" & Range("${positions[1]}").Validation.Formula1 & "|" & Range("${positions[2]}").Validation.Formula1
End Function`;
    const before = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(original), source, 'MakeCopy', [], 0, 'rows.xlsm', {});
    if (before.error || before.result?.value !== expected) {
        failures.push({case: name, phase: 'runtime', result: before.result, error: before.error});
        continue;
    }
    const saved = wasm.edit_xlsx_from_workbook(original, JSON.parse(JSON.stringify(before.workbook)));
    const after = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), source, 'ReadCopy', [], 0, 'rows.xlsm', {});
    if (after.error || after.result?.value !== expected) failures.push({case: name, phase: 'saved', result: after.result, error: after.error});
}
for (const [name, formula, operation, destinationSheet, expected] of [
    ['validation_cut_self', '=A1+$A$2', 'Range("A1:A3").Cut Range("C1")', 'Source', '=C1+$C$2|=C2+$C$2|=C3+$C$2|=C1'],
    ['validation_cut_cross', '=D1+$E$2', 'Range("A1:A3").Cut Worksheets("Target").Range("C1")', 'Target', '=F1+$E$2|=F2+$E$2|=F3+$E$2|=Target!C1'],
    ['validation_cut_other_sheet', '=Other!D1', 'Range("A1:A3").Cut Worksheets("Target").Range("C1")', 'Target', '=Other!D1|=Other!D2|=Other!D3|=Target!C1'],
]) {
    const source = `Public Function MakeCut() As String
Worksheets.Add: ActiveSheet.Name="Source"
Worksheets.Add: ActiveSheet.Name="Target"
Worksheets.Add: ActiveSheet.Name="Other"
Worksheets("Source").Activate
Range("A1:A3").Validation.Add 7,1,1,"${formula}"
Range("G1").Validation.Add 7,1,1,"=A1"
${operation}
MakeCut=ReadCut()
End Function
Public Function ReadCut() As String
ReadCut=Worksheets("${destinationSheet}").Range("C1").Validation.Formula1 & "|" & Worksheets("${destinationSheet}").Range("C2").Validation.Formula1 & "|" & Worksheets("${destinationSheet}").Range("C3").Validation.Formula1 & "|" & Worksheets("Source").Range("G1").Validation.Formula1
End Function`;
    const before = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(original), source, 'MakeCut', [], 0, 'rows.xlsm', {});
    if (before.error || before.result?.value !== expected) {
        failures.push({case: name, phase: 'runtime', result: before.result, error: before.error});
        continue;
    }
    const saved = wasm.edit_xlsx_from_workbook(original, JSON.parse(JSON.stringify(before.workbook)));
    const after = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), source, 'ReadCut', [], 0, 'rows.xlsm', {});
    if (after.error || after.result?.value !== expected) failures.push({case: name, phase: 'saved', result: after.result, error: after.error});
}
const wrapSource = `Public Function MakeWrap() As String
Worksheets.Add
Range("A1:B1").Validation.Add 7,1,1,"=XFD1048576"
Range("C3").Validation.Add 7,1,1,"=A1"
Range("C3").Copy Range("A3")
MakeWrap=ReadWrap()
End Function
Public Function ReadWrap() As String
ReadWrap=Range("A1").Validation.Formula1 & "|" & Range("B1").Validation.Formula1 & "|" & Range("A3").Validation.Formula1
End Function`;
const wrapExpected = '=XFD1048576|=A1048576|=XFC1';
const wrapBefore = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(original), wrapSource, 'MakeWrap', [], 0, 'rows.xlsm', {});
if (wrapBefore.error || wrapBefore.result?.value !== wrapExpected) {
    failures.push({case: 'validation_edge_wrap', phase: 'runtime', result: wrapBefore.result, error: wrapBefore.error});
} else {
    const saved = wasm.edit_xlsx_from_workbook(original, JSON.parse(JSON.stringify(wrapBefore.workbook)));
    const after = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), wrapSource, 'ReadWrap', [], 0, 'rows.xlsm', {});
    if (after.error || after.result?.value !== wrapExpected) failures.push({case: 'validation_edge_wrap', phase: 'saved', result: after.result, error: after.error});
}
const overlapSource = `Public Function MakeOverlap() As String
Worksheets.Add
Range("A1").Value=1: Range("A2").Value=2
Range("A1").Validation.Add 7,1,1,"=D1"
Range("A2").Validation.Add 7,1,1,"=F2"
On Error Resume Next
Range("A1:A2").Copy Range("A2:A5")
Range("H1").Value=Err.Number
On Error GoTo 0
MakeOverlap=ReadOverlap()
End Function
Public Function ReadOverlap() As String
ReadOverlap=CStr(Range("H1").Value) & "|" & CStr(Range("A1").Value) & "|" & CStr(Range("A2").Value) & "|" & Range("A2").Validation.Formula1
End Function`;
const overlapExpected = '1004|1|2|=F2';
const overlapBefore = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(original), overlapSource, 'MakeOverlap', [], 0, 'rows.xlsm', {});
if (overlapBefore.error || overlapBefore.result?.value !== overlapExpected) {
    failures.push({case: 'overlapping_repeat_copy', phase: 'runtime', result: overlapBefore.result, error: overlapBefore.error});
} else {
    const saved = wasm.edit_xlsx_from_workbook(original, JSON.parse(JSON.stringify(overlapBefore.workbook)));
    const after = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), overlapSource, 'ReadOverlap', [], 0, 'rows.xlsm', {});
    if (after.error || after.result?.value !== overlapExpected) failures.push({case: 'overlapping_repeat_copy', phase: 'saved', result: after.result, error: after.error});
}
for (const [name, operation, expected] of [
    ['clipboard_live_value', 'Range("A1").Value=2', '0|2|=F1'],
    ['clipboard_live_validation', 'Range("A1").Validation.Delete: Range("A1").Validation.Add 7,1,1,"=F1"', '0|1|=H1'],
    ['clipboard_clear_cancels', 'Range("D1").ClearContents', '1004||'],
]) {
    const source = `Public Function MakeLive() As String
Worksheets.Add
Range("A1").Value=1
Range("A1").Validation.Add 7,1,1,"=D1"
Range("A1").Copy
${operation}
On Error Resume Next
Range("C1").PasteSpecial -4104
Dim n As Long: n=Err.Number: Err.Clear
Dim f As String: f=Range("C1").Validation.Formula1: Err.Clear
MakeLive=CStr(n) & "|" & CStr(Range("C1").Value) & "|" & f
End Function`;
    const result = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(original), source, 'MakeLive', [], 0, 'rows.xlsm', {});
    if (result.error || result.result?.value !== expected) failures.push({case: name, phase: 'runtime', result: result.result, error: result.error});
}
const unevenPasteSource = `Public Function MakeUneven() As String
Worksheets.Add
Range("A1").Value=1: Range("A2").Value=2
Range("C1:D3").Value=9
Range("A1:A2").Copy
Range("C1:D3").PasteSpecial -4163
MakeUneven=ReadUneven()
End Function
Public Function ReadUneven() As String
ReadUneven=CStr(Range("C1").Value) & "|" & CStr(Range("C2").Value) & "|" & CStr(Range("C3").Value) & "|" & CStr(Range("D1").Value)
End Function`;
const unevenExpected = '1|2|9|9';
const unevenBefore = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(original), unevenPasteSource, 'MakeUneven', [], 0, 'rows.xlsm', {});
if (unevenBefore.error || unevenBefore.result?.value !== unevenExpected) {
    failures.push({case: 'uneven_paste', phase: 'runtime', result: unevenBefore.result, error: unevenBefore.error});
} else {
    const saved = wasm.edit_xlsx_from_workbook(original, JSON.parse(JSON.stringify(unevenBefore.workbook)));
    const after = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), unevenPasteSource, 'ReadUneven', [], 0, 'rows.xlsm', {});
    if (after.error || after.result?.value !== unevenExpected) failures.push({case: 'uneven_paste', phase: 'saved', result: after.result, error: after.error});
}
const skipValidationSource = `Public Function MakeSkip() As String
Worksheets.Add
Range("A1").Value=7
Range("C1").Validation.Add 7,1,1,"=H1"
Range("A1").Copy
Range("C1").PasteSpecial 6,-4142,True,False
MakeSkip=ReadSkip()
End Function
Public Function ReadSkip() As String
ReadSkip=Range("C1").Validation.Formula1
End Function`;
const skipBefore = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(original), skipValidationSource, 'MakeSkip', [], 0, 'rows.xlsm', {});
if (skipBefore.error || skipBefore.result?.value !== '=H1') {
    failures.push({case: 'skip_empty_validation', phase: 'runtime', result: skipBefore.result, error: skipBefore.error});
} else {
    const saved = wasm.edit_xlsx_from_workbook(original, JSON.parse(JSON.stringify(skipBefore.workbook)));
    const after = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), skipValidationSource, 'ReadSkip', [], 0, 'rows.xlsm', {});
    if (after.error || after.result?.value !== '=H1') failures.push({case: 'skip_empty_validation', phase: 'saved', result: after.result, error: after.error});
}
for (const [name, operation, expected] of [
    ['overlap_validation_skip', 'Range("A1").Copy: Range("C2").PasteSpecial 6,-4142,True,False', '=H1|=H2|=H3|=H4|none'],
    ['overlap_validation_copy', 'Range("C2").Copy Range("E1")', '=H1|=H2|=H3|=H4|=J1'],
    ['overlap_validation_values', 'Range("A1").Copy: Range("E1").PasteSpecial -4163', '=H1|none|none|=H4|none'],
]) {
    const source = `Public Function MakeOverlapRule() As String
Worksheets.Add
Range("C1:C3,C2:C4").Validation.Add 7,1,1,"=H1"
${operation}
MakeOverlapRule=ReadOverlapRule()
End Function
Public Function ReadOverlapRule() As String
ReadOverlapRule=One("C1") & "|" & One("C2") & "|" & One("C3") & "|" & One("C4") & "|" & One("E1")
End Function
Public Function One(ByVal address As String) As String
On Error Resume Next
Dim n As Long: n=Range(address).Validation.Type
If Err.Number<>0 Then Err.Clear: One="none": Exit Function
One=Range(address).Validation.Formula1
End Function`;
    const before = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(original), source, 'MakeOverlapRule', [], 0, 'rows.xlsm', {});
    if (before.error || before.result?.value !== expected) {
        failures.push({case: name, phase: 'runtime', result: before.result, error: before.error});
        continue;
    }
    const saved = wasm.edit_xlsx_from_workbook(original, JSON.parse(JSON.stringify(before.workbook)));
    const after = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), source, 'ReadOverlapRule', [], 0, 'rows.xlsm', {});
    if (after.error || after.result?.value !== expected) failures.push({case: name, phase: 'saved', result: after.result, error: after.error});
}
const distinctBook = JSON.parse(JSON.stringify(wasm.parse_spreadsheet(original)));
distinctBook.sheets[0].validations = [
    {kind: 'custom', ranges: [[0, 2, 2, 2]], formula1: 'H1'},
    {kind: 'custom', ranges: [[1, 2, 3, 2]], formula1: 'J2'},
];
const distinctFixture = wasm.edit_xlsx_from_workbook(original, distinctBook);
const distinctSource = `Public Function MakeDistinct() As String
Range("A1").Copy
Range("C2").PasteSpecial 6,-4142,True,False
MakeDistinct=ReadDistinct()
End Function
Public Function ReadDistinct() As String
ReadDistinct=Range("C1").Validation.Formula1 & "|" & Range("C2").Validation.Formula1 & "|" & Range("C3").Validation.Formula1 & "|" & Range("C4").Validation.Formula1
End Function`;
const distinctExpected = '=H1|=H2|=H3|=J4';
const distinctBefore = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(distinctFixture), distinctSource, 'MakeDistinct', [], 0, 'rows.xlsm', {});
if (distinctBefore.error || distinctBefore.result?.value !== distinctExpected) {
    failures.push({case: 'distinct_overlapping_rules', phase: 'runtime', result: distinctBefore.result, error: distinctBefore.error});
} else {
    const saved = wasm.edit_xlsx_from_workbook(distinctFixture, JSON.parse(JSON.stringify(distinctBefore.workbook)));
    const parsed = wasm.parse_spreadsheet(saved);
    const coverage = (parsed.sheets[0].validations || []).flatMap(rule => rule.ranges).filter(([r0,c0,r1,c1]) => r0 <= 1 && r1 >= 1 && c0 <= 2 && c1 >= 2).length;
    const after = wasm.run_spreadsheet_vba_with_outcome(parsed, distinctSource, 'ReadDistinct', [], 0, 'rows.xlsm', {});
    if (after.error || after.result?.value !== distinctExpected || coverage !== 2) failures.push({case: 'distinct_overlapping_rules', phase: 'saved', result: after.result, error: after.error, coverage});
}
const loadedConditionSource = `Public Function Probe() As String
Dim c As Object
Set c=Range("A2").FormatConditions(1)
Probe=CStr(Range("A2").FormatConditions.Count) & "|" & CStr(c.Type) & "|" & c.Formula1 & "|" & c.AppliesTo.Address & "|" & CStr(Range("A4").DisplayFormat.Interior.Color)
End Function`;
const loadedConditionBytes = await readFile(path.join(root, 'tests/fixtures/conditional/cell_is.xlsx'));
const loadedConditionBook = wasm.parse_spreadsheet(loadedConditionBytes);
const loadedConditionBefore = wasm.run_spreadsheet_vba_with_outcome(loadedConditionBook, loadedConditionSource, 'Probe', [], 0, 'conditions.xlsx', {});
const loadedConditionExpected = '2|1|=50|$A$2:$A$8|13551615';
if (loadedConditionBefore.error || loadedConditionBefore.result?.value !== loadedConditionExpected) {
    failures.push({case: 'loaded_conditional_format', phase: 'runtime', result: loadedConditionBefore.result, error: loadedConditionBefore.error});
} else {
    const saved = wasm.edit_xlsx_from_workbook(loadedConditionBytes, JSON.parse(JSON.stringify(loadedConditionBefore.workbook)));
    const after = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), loadedConditionSource, 'Probe', [], 0, 'conditions.xlsx', {});
    if (after.error || after.result?.value !== loadedConditionExpected) failures.push({case: 'loaded_conditional_format', phase: 'saved', result: after.result, error: after.error});
}
const conditionColorSource = `Public Function Probe() As String
Probe=CStr(Range("A2").DisplayFormat.Interior.Color) & "|" & CStr(Range("A3").DisplayFormat.Interior.Color) & "|" & CStr(Range("A4").DisplayFormat.Interior.Color) & "|" & CStr(Range("A5").DisplayFormat.Interior.Color)
End Function`;
const existingConditionStyleSource = `Public Function Change() As String
Dim c As Object
Set c=Range("A2").FormatConditions(1)
c.Interior.Color=255: c.Font.Color=65280: c.Font.Bold=True: c.Font.Strikethrough=True
Change=ReadRule()
End Function
Public Function ReadRule() As String
Dim c As Object
Set c=Range("A2").FormatConditions(1)
ReadRule=CStr(c.Interior.Color) & "|" & CStr(c.Font.Color) & "|" & CStr(Range("A4").DisplayFormat.Font.Bold) & "|" & CStr(Range("A4").DisplayFormat.Font.Strikethrough) & "|" & CStr(Range("A4").DisplayFormat.Interior.Color)
End Function`;
const existingConditionStyleExpected = '255|65280|True|True|255';
const existingConditionStyleBefore = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(loadedConditionBytes), existingConditionStyleSource, 'Change', [], 0, 'conditions.xlsx', {});
if (existingConditionStyleBefore.error || existingConditionStyleBefore.result?.value !== existingConditionStyleExpected) {
    failures.push({case: 'existing_condition_style', phase: 'runtime', result: existingConditionStyleBefore.result, error: existingConditionStyleBefore.error});
} else {
    const saved = wasm.edit_xlsx_from_workbook(loadedConditionBytes, JSON.parse(JSON.stringify(existingConditionStyleBefore.workbook)));
    const after = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), existingConditionStyleSource, 'ReadRule', [], 0, 'conditions.xlsx', {});
    if (after.error || after.result?.value !== existingConditionStyleExpected) failures.push({case: 'existing_condition_style', phase: 'saved', result: after.result, error: after.error});
}
const priorityConditionSource = `Public Function Change() As String
Range("A2").FormatConditions(1).Interior.Color=255
Dim c As Object
Set c=Range("A2:A8").FormatConditions.Add(2,,"=A2>0")
c.Interior.Color=65280: c.SetFirstPriority
Change=ReadOrder()
End Function
Public Function ReadOrder() As String
Dim a As Object, b As Object
Set a=Range("A2").FormatConditions(1): Set b=Range("A2").FormatConditions(2)
ReadOrder=a.Formula1 & "|" & CStr(a.Priority) & "|" & CStr(a.Interior.Color) & "|" & b.Formula1 & "|" & CStr(b.Priority) & "|" & CStr(b.Interior.Color)
End Function`;
const priorityConditionExpected = '=A2>0|1|65280|=50|2|255';
const priorityBefore = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(loadedConditionBytes), priorityConditionSource, 'Change', [], 0, 'conditions.xlsx', {});
if (priorityBefore.error || priorityBefore.result?.value !== priorityConditionExpected) {
    failures.push({case: 'conditional_priority', phase: 'runtime', result: priorityBefore.result, error: priorityBefore.error});
} else {
    const saved = wasm.edit_xlsx_from_workbook(loadedConditionBytes, JSON.parse(JSON.stringify(priorityBefore.workbook)));
    const after = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), priorityConditionSource, 'ReadOrder', [], 0, 'conditions.xlsx', {});
    if (after.error || after.result?.value !== priorityConditionExpected) failures.push({case: 'conditional_priority', phase: 'saved', result: after.result, error: after.error});
}
const newTextConditionSource = `Public Function Change() As String
Dim c As Object
Range("B2").Value="   ": Range("C2").Value="x": Range("D2").Value=CVErr(xlErrNA): Range("E2").Value=1
Range("F2").Value="ABC": Range("G2").Value="xyz": Range("H2").Value="abcd": Range("I2").Value="zabc"
Set c=Range("B2:B5").FormatConditions.Add(10): c.Interior.Color=255
Set c=Range("C2:C5").FormatConditions.Add(13): c.Interior.Color=255
Set c=Range("D2:D5").FormatConditions.Add(16): c.Interior.Color=255
Set c=Range("E2:E5").FormatConditions.Add(17): c.Interior.Color=255
Set c=Range("F2:F5").FormatConditions.Add(9,,,,"abc",0): c.Interior.Color=255
Set c=Range("G2:G5").FormatConditions.Add(9,,,,"abc",1): c.Interior.Color=255
Set c=Range("H2:H5").FormatConditions.Add(9,,,,"abc",2): c.Interior.Color=255
Set c=Range("I2:I5").FormatConditions.Add(9,,,,"abc",3): c.Interior.Color=255
Change=ReadRule()
End Function
Public Function ReadRule() As String
Dim c As Object, i As Long
For i=2 To 9
Set c=Cells(2,i).FormatConditions(1)
ReadRule=ReadRule & CStr(c.Type) & "|" & c.Formula1 & "|" & CStr(Cells(2,i).DisplayFormat.Interior.Color) & ";"
Next i
End Function`;
const newTextConditionExpected = '10|=LEN(TRIM(B2))=0|255;13|=LEN(TRIM(C2))>0|255;16|=ISERROR(D2)|255;17|=NOT(ISERROR(E2))|255;9|=NOT(ISERROR(SEARCH("abc",F2)))|255;9|=ISERROR(SEARCH("abc",G2))|255;9|=LEFT(H2,LEN("abc"))="abc"|255;9|=RIGHT(I2,LEN("abc"))="abc"|255;';
const newTextBefore = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(loadedConditionBytes), newTextConditionSource, 'Change', [], 0, 'conditions.xlsx', {});
if (newTextBefore.error || newTextBefore.result?.value !== newTextConditionExpected) {
    failures.push({case: 'new_text_blank_error_conditions', phase: 'runtime', result: newTextBefore.result, error: newTextBefore.error});
} else {
    const saved = wasm.edit_xlsx_from_workbook(loadedConditionBytes, JSON.parse(JSON.stringify(newTextBefore.workbook)));
    const after = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), newTextConditionSource, 'ReadRule', [], 0, 'conditions.xlsx', {});
    if (after.error || after.result?.value !== newTextConditionExpected) failures.push({case: 'new_text_blank_error_conditions', phase: 'saved', result: after.result, error: after.error});
}
for (const [fixture, statement, read, expected] of [
    ['cell_is.xlsx', 'Range("A2").FormatConditions(1).Modify 2,,"=A2>30"', 'CStr(c.Type) & "|" & c.Formula1 & "|" & CStr(c.Priority) & "|" & CStr(c.Interior.Color)', '2|=A2>30|3|13551615'],
    ['expression.xlsx', 'Range("A2").FormatConditions(1).Modify 1,1,5,20', 'CStr(c.Type) & "|" & CStr(c.Operator) & "|" & c.Formula1 & "|" & c.Formula2 & "|" & CStr(c.Priority) & "|" & CStr(c.Interior.Color)', '1|1|=5|=20|3|13551615'],
]) {
    const bytes = await readFile(path.join(root, 'tests/fixtures/conditional', fixture));
    const selector = fixture === 'cell_is.xlsx' ? 2 : 1;
    const source = `Public Function Change() As String
${statement}
Change=ReadRule()
End Function
Public Function ReadRule() As String
Dim c As Object
Set c=Range("A2").FormatConditions(${selector})
ReadRule=${read}
End Function`;
    const before = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(bytes), source, 'Change', [], 0, fixture, {});
    if (before.error || before.result?.value !== expected) {
        failures.push({case: `modify_${fixture}`, phase: 'runtime', result: before.result, error: before.error});
    } else {
        const saved = wasm.edit_xlsx_from_workbook(bytes, JSON.parse(JSON.stringify(before.workbook)));
        const after = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), source, 'ReadRule', [], 0, fixture, {});
        if (after.error || after.result?.value !== expected) failures.push({case: `modify_${fixture}`, phase: 'saved', result: after.result, error: after.error});
    }
}
for (const [fixture, settings, read, expected] of [
    ['top10.xlsx', 'c.Rank=3: c.TopBottom=0: c.Percent=True: c.StopIfTrue=True', 'CStr(c.Rank) & "|" & CStr(c.TopBottom) & "|" & CStr(c.Percent) & "|" & CStr(c.StopIfTrue)', '3|0|True|True'],
    ['duplicates.xlsx', 'c.DupeUnique=0: c.StopIfTrue=True', 'CStr(c.DupeUnique) & "|" & CStr(c.StopIfTrue)', '0|True'],
]) {
    const bytes = await readFile(path.join(root, 'tests/fixtures/conditional', fixture));
    const source = `Public Function Change() As String
Dim c As Object
Set c=Range("A2").FormatConditions(1)
${settings}
Change=${read}
End Function
Public Function ReadRule() As String
Dim c As Object
Set c=Range("A2").FormatConditions(1)
ReadRule=${read}
End Function`;
    const before = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(bytes), source, 'Change', [], 0, fixture, {});
    if (before.error || before.result?.value !== expected) {
        failures.push({case: `settings_${fixture}`, phase: 'runtime', result: before.result, error: before.error});
    } else {
        const saved = wasm.edit_xlsx_from_workbook(bytes, JSON.parse(JSON.stringify(before.workbook)));
        const after = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), source, 'ReadRule', [], 0, fixture, {});
        if (after.error || after.result?.value !== expected) failures.push({case: `settings_${fixture}`, phase: 'saved', result: after.result, error: after.error});
    }
}
for (const [name, statement, expected] of [
    ['delete_one_condition', 'Range("A2").FormatConditions(1).Delete', '1'],
    ['delete_all_conditions', 'Range("A2:A8").FormatConditions.Delete', '0'],
]) {
    const source = `Public Function Remove() As String
${statement}
Remove=ReadCount()
End Function
Public Function ReadCount() As String
ReadCount=CStr(Range("A2").FormatConditions.Count)
End Function`;
    const before = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(loadedConditionBytes), source, 'Remove', [], 0, 'conditions.xlsx', {});
    if (before.error || before.result?.value !== expected) {
        failures.push({case: name, phase: 'runtime', result: before.result, error: before.error});
    } else {
        const saved = wasm.edit_xlsx_from_workbook(loadedConditionBytes, JSON.parse(JSON.stringify(before.workbook)));
        const after = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), source, 'ReadCount', [], 0, 'conditions.xlsx', {});
        if (after.error || after.result?.value !== expected) failures.push({case: name, phase: 'saved', result: after.result, error: after.error});
    }
}
const extendedIconBytes = await readFile(path.join(root, 'tests/fixtures/conditional/x14_icons.xlsx'));
for (const [action, address, expected] of [
    ['Range("A2:A6").Cut Range("E2")', 'E2', '$E$2:$E$6|1|=$B$1'],
    ['Dim original As Object: Set original=ActiveSheet\nWorksheets.Add\noriginal.Range("A2:A6").Cut Range("F2")', 'F2', '$F$2:$F$6|1|=$B$1'],
]) {
    const source = `Public Function Change() As String
Dim c As Object: Set c=Range("A2").FormatConditions(1)
c.IconCriteria(2).Type=4: c.IconCriteria(2).Value="=$B$1"
${action}
Change=ReadRule()
End Function
Public Function ReadRule() As String
Dim c As Object: Set c=Range("${address}").FormatConditions(1)
ReadRule=c.AppliesTo.Address & "|" & CStr(c.Priority) & "|" & CStr(c.IconCriteria(2).Value)
End Function`;
    const before = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(extendedIconBytes), source, 'Change', [], 0, 'icons.xlsx', {});
    if (before.error || before.result?.value !== expected) failures.push({case: 'extended icon cut', action, phase: 'runtime', result: before.result, error: before.error});
    else {
        const saved = wasm.edit_xlsx_from_workbook(extendedIconBytes, JSON.parse(JSON.stringify(before.workbook)));
        const after = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), source, 'ReadRule', [], before.active_sheet, 'icons.xlsx', {});
        if (after.error || after.result?.value !== expected) failures.push({case: 'extended icon cut', action, phase: 'saved', result: after.result, error: after.error});
    }
}
for (const [action, address, expected] of [
    ['Range("A2:A6").Copy Range("E2")', 'E2', '$E$2:$E$6|1|True|True|=$B$1'],
    ['Dim original As Object: Set original=ActiveSheet\nWorksheets.Add\noriginal.Range("A2:A6").Copy Range("F2")', 'F2', '$F$2:$F$6|1|True|True|=$B$1'],
]) {
    const source = `Public Function Change() As String
Dim c As Object: Set c=Range("A2").FormatConditions(1)
c.ShowIconOnly=True: c.ReverseOrder=True
c.IconCriteria(2).Type=4: c.IconCriteria(2).Value="=$B$1"
${action}
Change=ReadRule()
End Function
Public Function ReadRule() As String
Dim c As Object: Set c=Range("${address}").FormatConditions(1)
ReadRule=c.AppliesTo.Address & "|" & CStr(c.Priority) & "|" & CStr(c.ShowIconOnly) & "|" & CStr(c.ReverseOrder) & "|" & CStr(c.IconCriteria(2).Value)
End Function`;
    const before = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(extendedIconBytes), source, 'Change', [], 0, 'icons.xlsx', {});
    if (before.error || before.result?.value !== expected) failures.push({case: 'extended icon copy', action, phase: 'runtime', result: before.result, error: before.error});
    else {
        const saved = wasm.edit_xlsx_from_workbook(extendedIconBytes, JSON.parse(JSON.stringify(before.workbook)));
        const after = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), source, 'ReadRule', [], before.active_sheet, 'icons.xlsx', {});
        if (after.error || after.result?.value !== expected) failures.push({case: 'extended icon copy', action, phase: 'saved', result: after.result, error: after.error});
    }
}
{
    const source = `Public Sub Change()
Dim original As Object: Set original=ActiveSheet
Worksheets.Add
original.Range("A2:A6").Copy Range("F2")
Worksheets.Add
original.Range("A2:A6").Copy Range("F2")
End Sub`;
    const before = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(extendedIconBytes), source, 'Change', [], 0, 'icons.xlsx', {});
    if (before.error) failures.push({case: 'extended icon copy ids', phase: 'runtime', error: before.error});
    else {
        const saved = wasm.edit_xlsx_from_workbook(extendedIconBytes, JSON.parse(JSON.stringify(before.workbook)));
        const book = wasm.parse_spreadsheet(saved);
        const ids = book.sheets.flatMap(sheet => (sheet.conditional_rules || []).filter(rule => rule.visual_extension_only && rule.ranges.some(range => range[1] === 5)).map(rule => rule.visual_extension_id));
        if (ids.length !== 2 || new Set(ids).size !== 2) failures.push({case: 'extended icon copy ids', phase: 'saved', ids});
    }
}
for (const [name, source, expectedChange, expectedRead] of [
    ['extended priorities', `Public Function Change() As String
Dim a As Object, b As Object, c As Object, s As String
Set a=Range("A2").FormatConditions(1): Set b=Range("B2").FormatConditions(1): Set c=Range("C2").FormatConditions(1)
s=ReadRule() & ";"
c.SetFirstPriority: s=s & ReadRule() & ";"
a.SetLastPriority: s=s & ReadRule() & ";"
b.Priority=1
Change=s & ReadRule()
End Function
Public Function ReadRule() As String
ReadRule=CStr(Range("A2").FormatConditions(1).Priority) & "," & CStr(Range("B2").FormatConditions(1).Priority) & "," & CStr(Range("C2").FormatConditions(1).Priority)
End Function`, '1,2,3;2,3,1;4,3,1;5,1,2', '5,1,2'],
    ['mixed priorities', `Public Function Change() As String
Dim b As Object, c As Object
Set c=Range("B2").FormatConditions(1)
c.ReverseOrder=True: c.IconCriteria(2).Type=4: c.IconCriteria(2).Value="=$B$1"
Set b=Range("A2:A6").FormatConditions.AddDatabar
b.SetFirstPriority
Rows(1).Insert
Change=ReadRule()
End Function
Public Function ReadRule() As String
Dim a As Object, b As Object, c As Object
Set a=Range("A3").FormatConditions(1): Set b=Range("A3").FormatConditions(2): Set c=Range("B3").FormatConditions(1)
ReadRule=CStr(a.Type) & ":" & CStr(a.Priority) & "|" & CStr(b.Type) & ":" & CStr(b.Priority) & "|" & CStr(c.Priority) & "|" & CStr(c.ReverseOrder) & "|" & CStr(c.IconCriteria(2).Value)
End Function`, '4:1|6:2|3|True|=$B$2', '4:1|6:2|3|True|=$B$2'],
]) {
    const before = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(extendedIconBytes), source, 'Change', [], 0, 'icons.xlsx', {});
    if (before.error || before.result?.value !== expectedChange) failures.push({case: name, phase: 'runtime', result: before.result, error: before.error});
    else {
        const saved = wasm.edit_xlsx_from_workbook(extendedIconBytes, JSON.parse(JSON.stringify(before.workbook)));
        const after = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), source, 'ReadRule', [], 0, 'icons.xlsx', {});
        if (after.error || after.result?.value !== expectedRead) failures.push({case: name, phase: 'saved', result: after.result, error: after.error});
    }
}
{
    const source = `Public Function Change() As String
Dim c As Object
Set c=Range("A2").FormatConditions(1)
Rows(2).Delete
On Error Resume Next
c.Delete
Change=CStr(Err.Number) & "|" & CStr(Range("A2").FormatConditions.Count)
End Function`;
    const result = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(extendedIconBytes), source, 'Change', [], 0, 'icons.xlsx', {});
    if (result.error || result.result?.value !== '424|1') failures.push({case: 'condition deletion after source removed', result: result.result, error: result.error});
}
for (const [scope, inserted, expected] of [
    ['A2', 1, '0,0,1,1,1,1,'], ['A3:A4', 4, '1,0,0,0,1,1,'], ['A2:A6', 4, '0,0,0,0,0,0,'],
]) {
    const source = `Public Function Change() As String
Dim c As Object
Set c=Range("${scope}").FormatConditions(1)
Rows(${inserted}).Insert
c.Delete
Change=ReadRule()
End Function
Public Function ReadRule() As String
Dim i As Long
For i=2 To 7
ReadRule=ReadRule & CStr(Cells(i,1).FormatConditions.Count) & ","
Next i
End Function`;
    const before = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(extendedIconBytes), source, 'Change', [], 0, 'icons.xlsx', {});
    if (before.error || before.result?.value !== expected) failures.push({case: 'condition deletion after insertion', scope, phase: 'runtime', result: before.result, error: before.error});
    else {
        const saved = wasm.edit_xlsx_from_workbook(extendedIconBytes, JSON.parse(JSON.stringify(before.workbook)));
        const after = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), source, 'ReadRule', [], 0, 'icons.xlsx', {});
        if (after.error || after.result?.value !== expected) failures.push({case: 'condition deletion after insertion', scope, phase: 'saved', result: after.result, error: after.error});
    }
}
for (const [scope, expected] of [
    ['A2', '1,2,2,2,2,'], ['A3:A4', '2,1,1,2,2,'],
    ['A1:A3', '1,1,2,2,2,'], ['A2:A6', '1,1,1,1,1,'],
]) {
    const source = `Public Function Change() As String
Range("A2:A6").FormatConditions.AddDatabar
Range("${scope}").FormatConditions(1).Delete
Change=ReadRule()
End Function
Public Function ReadRule() As String
Dim i As Long
For i=2 To 6
ReadRule=ReadRule & CStr(Cells(i,1).FormatConditions.Count) & ","
Next i
End Function`;
    const before = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(extendedIconBytes), source, 'Change', [], 0, 'icons.xlsx', {});
    if (before.error || before.result?.value !== expected) failures.push({case: 'scoped condition deletion', scope, phase: 'runtime', result: before.result, error: before.error});
    else {
        const saved = wasm.edit_xlsx_from_workbook(extendedIconBytes, JSON.parse(JSON.stringify(before.workbook)));
        const after = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), source, 'ReadRule', [], 0, 'icons.xlsx', {});
        if (after.error || after.result?.value !== expected) failures.push({case: 'scoped condition deletion', scope, phase: 'saved', result: after.result, error: after.error});
    }
}
{
    const source = `Public Function Change() As String
Dim c As Object
Set c=Range("A2").FormatConditions(1)
c.ModifyAppliesToRange Range("D2:D6")
Range("B2").FormatConditions.Delete
Range("C2:C6").FormatConditions.Delete
Change=ReadRule()
End Function
Public Function ReadRule() As String
Dim c As Object
ReadRule=CStr(Range("A2").FormatConditions.Count) & "|" & CStr(Range("B2").FormatConditions.Count) & "|" & CStr(Range("B3").FormatConditions.Count) & "|" & CStr(Range("C3").FormatConditions.Count) & "|"
Set c=Range("D2").FormatConditions(1)
ReadRule=ReadRule & c.AppliesTo.Address & "|" & CStr(c.Priority) & "|" & CStr(c.IconCriteria.Count)
End Function`;
    const expected = '0|0|1|0|$D$2:$D$6|1|3';
    const before = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(extendedIconBytes), source, 'Change', [], 0, 'icons.xlsx', {});
    if (before.error || before.result?.value !== expected) failures.push({case: 'extended icon range and deletion', phase: 'runtime', result: before.result, error: before.error});
    else {
        const saved = wasm.edit_xlsx_from_workbook(extendedIconBytes, JSON.parse(JSON.stringify(before.workbook)));
        const after = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), source, 'ReadRule', [], 0, 'icons.xlsx', {});
        if (after.error || after.result?.value !== expected) failures.push({case: 'extended icon range and deletion', phase: 'saved', result: after.result, error: after.error});
    }
}
for (const kind of [0, 3, 4, 5]) {
    const source = `Public Function Change() As String
Dim c As Object
Set c=Range("A2").FormatConditions(1)
c.IconCriteria(2).Type=${kind}: c.IconCriteria(2).Value="=$B$1"
Range("B2").FormatConditions(1).IconCriteria(2).Value=Null
Rows(1).Insert
Change=ReadRule()
End Function
Public Function ReadRule() As String
Dim c As Object
Set c=Range("A3").FormatConditions(1)
ReadRule=c.AppliesTo.Address & "|" & CStr(c.IconCriteria(2).Value)
End Function`;
    const expected = '$A$3:$A$7|=$B$2';
    const before = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(extendedIconBytes), source, 'Change', [], 0, 'icons.xlsx', {});
    if (before.error || before.result?.value !== expected) failures.push({case: 'extended icon reference shift', kind, phase: 'runtime', result: before.result, error: before.error});
    else {
        const saved = wasm.edit_xlsx_from_workbook(extendedIconBytes, JSON.parse(JSON.stringify(before.workbook)));
        const after = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), source, 'ReadRule', [], 0, 'icons.xlsx', {});
        if (after.error || after.result?.value !== expected) failures.push({case: 'extended icon reference shift', kind, phase: 'saved', result: after.result, error: after.error});
    }
}
{
    const source = `Public Function ReadRule() As String
Dim c As Object, p As Object, col As Long
For col=1 To 3
Set c=Cells(2,col).FormatConditions(1)
ReadRule=ReadRule & CStr(c.Type) & ":" & CStr(c.IconCriteria.Count) & ":"
For Each p In c.IconCriteria
ReadRule=ReadRule & CStr(p.Value) & ","
Next p
ReadRule=ReadRule & ";"
Next col
End Function
Public Function Change() As String
Range("F1").Value=42
Range("D2:D6").FormatConditions.AddDatabar
Change=ReadRule()
End Function
Public Function EditIcon() As String
Dim c As Object
Set c=Range("A2").FormatConditions(1)
c.IconCriteria(2).Type=0: c.IconCriteria(2).Value=42: c.IconCriteria(2).Operator=5
c.IconCriteria(3).Type=4: c.IconCriteria(3).Value="=$B$1"
c.ShowIconOnly=True: c.ReverseOrder=True
EditIcon=ReadIcon()
End Function
Public Function ReadIcon() As String
Dim c As Object, p As Object
Set c=Range("A2").FormatConditions(1)
ReadIcon=CStr(c.ShowIconOnly) & "|" & CStr(c.ReverseOrder) & "|"
For Each p In c.IconCriteria
ReadIcon=ReadIcon & CStr(p.Type) & "," & CStr(p.Value) & "," & CStr(p.Operator) & ";"
Next p
End Function`;
    const expected = '6:3:0,33,67,;6:3:0,33,67,;6:5:0,20,40,60,80,;';
    const before = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(extendedIconBytes), source, 'Change', [], 0, 'icons.xlsx', {});
    if (before.error || before.result?.value !== expected) failures.push({case: 'extended icon read', phase: 'runtime', result: before.result, error: before.error});
    else {
        const saved = wasm.edit_xlsx_from_workbook(extendedIconBytes, JSON.parse(JSON.stringify(before.workbook)));
        const after = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), source, 'ReadRule', [], 0, 'icons.xlsx', {});
        if (after.error || after.result?.value !== expected) failures.push({case: 'extended icon read', phase: 'saved', result: after.result, error: after.error});
        const edited = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), source, 'EditIcon', [], 0, 'icons.xlsx', {});
        const expectedEdit = 'True|True|3,0,7;0,42,5;4,=$B$1,7;';
        if (edited.error || edited.result?.value !== expectedEdit) failures.push({case: 'extended icon edit', phase: 'runtime', result: edited.result, error: edited.error});
        else {
            const savedEdit = wasm.edit_xlsx_from_workbook(saved, JSON.parse(JSON.stringify(edited.workbook)));
            const reloaded = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(savedEdit), source, 'ReadIcon', [], 0, 'icons.xlsx', {});
            if (reloaded.error || reloaded.result?.value !== expectedEdit) failures.push({case: 'extended icon edit', phase: 'saved', result: reloaded.result, error: reloaded.error});
        }
    }
}
const visualConditionBytes = await readFile(path.join(root, 'tests/fixtures/conditional/visual_rules.xlsx'));
{
    const source = `Public Function Change() As String
Cells.FormatConditions.Delete
Dim c As Object, p As Object, v As Variant, s As String
Set c=Range("A2:A6").FormatConditions.AddIconSetCondition
Set p=c.IconCriteria(2)
p.Type=False: s=CStr(p.Type) & ";"
p.Type=3: p.Type=v: s=s & CStr(p.Type) & ";"
Set c=Range("B2:B6").FormatConditions.AddColorScale(3)
Set p=c.ColorScaleCriteria(2)
p.Type=False: s=s & CStr(p.Type) & ";"
p.Type=3: p.Type=v: s=s & CStr(p.Type)
Change=s
End Function`;
    const result = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(visualConditionBytes), source, 'Change', [], 0, 'visual.xlsx', {});
    if (result.error || result.result?.value !== '0;0;0;0') failures.push({case: 'icon and scale type coercion', result: result.result, error: result.error});
}
{
    const source = `Public Function Change() As String
Cells.FormatConditions.Delete
Dim c As Object, p As Object, s As String
Set c=Range("C2:C6").FormatConditions.AddIconSetCondition
On Error Resume Next
c.IconCriteria(1).Type=3: s=CStr(Err.Number<>0) & ";": Err.Clear
c.IconCriteria(1).Value=5: s=s & CStr(Err.Number<>0) & ";": Err.Clear
c.IconCriteria(2).Operator=6: s=s & CStr(Err.Number<>0) & ";": Err.Clear
On Error GoTo 0
Set p=c.IconCriteria(2): p.Type=0: p.Value=23: p.Type=3
s=s & CStr(p.Value) & ";"
p.Value="=$B$1": p.Type=0: s=s & CStr(p.Value) & ";"
p.Type=4: p.Value="=$B$1": p.Value=12: s=s & CStr(p.Value)
Change=s
End Function`;
    const result = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(visualConditionBytes), source, 'Change', [], 0, 'visual.xlsx', {});
    if (result.error || result.result?.value !== 'True;True;True;33;0;12') failures.push({case: 'icon criterion validation', result: result.result, error: result.error});
}
for (const setup of ['', 'Cells.FormatConditions.Delete\nRange("C2:C6").FormatConditions.AddIconSetCondition\n']) {
    const source = `Public Function Change() As String
${setup}Dim c As Object, s As String
Set c=Range("C2").FormatConditions(1)
s=ReadRule() & "|"
c.IconCriteria(1).Operator=5
c.IconCriteria(2).Type=0: c.IconCriteria(2).Value=42: c.IconCriteria(2).Operator=5
c.IconCriteria(3).Type=4: c.IconCriteria(3).Value="=$B$1"
Rows(1).Insert
Change=s & ReadMoved()
End Function
Public Function ReadRule() As String
Dim c As Object, p As Object
Set c=Range("C2").FormatConditions(1)
For Each p In c.IconCriteria
ReadRule=ReadRule & CStr(p.Type) & "," & CStr(p.Value) & "," & CStr(p.Operator) & ";"
Next p
End Function
Public Function ReadMoved() As String
Dim c As Object, p As Object
Set c=Range("C3").FormatConditions(1)
For Each p In c.IconCriteria
ReadMoved=ReadMoved & CStr(p.Type) & "," & CStr(p.Value) & "," & CStr(p.Operator) & ";"
Next p
End Function`;
    const before = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(visualConditionBytes), source, 'Change', [], 0, 'visual.xlsx', {});
    if (before.error || before.result?.value !== '3,0,7;3,33,7;3,67,7;|3,0,5;0,42,5;4,=$B$2,7;') failures.push({case: 'icon criteria', setup, phase: 'runtime', result: before.result, error: before.error});
    else {
        const saved = wasm.edit_xlsx_from_workbook(visualConditionBytes, JSON.parse(JSON.stringify(before.workbook)));
        const after = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), source, 'ReadMoved', [], 0, 'visual.xlsx', {});
        if (after.error || after.result?.value !== '3,0,5;0,42,5;4,=$B$2,7;') failures.push({case: 'icon criteria', setup, phase: 'saved', result: after.result, error: after.error});
    }
}
for (const setup of ['', 'Cells.FormatConditions.Delete\nRange("B2:B6").FormatConditions.AddDatabar\nRange("C2:C6").FormatConditions.AddIconSetCondition\n']) {
    const source = `Public Function Change() As String
${setup}Dim b As Object, i As Object, s As String
Set b=Range("B2").FormatConditions(1): Set i=Range("C2").FormatConditions(1)
s=ReadRule() & ";"
b.ShowValue=False: i.ShowIconOnly=True: i.ReverseOrder=True
s=s & ReadRule() & ";"
b.ShowValue="0": i.ShowIconOnly=2
Change=s & ReadRule()
End Function
Public Function ResetRule() As String
Range("B2").FormatConditions(1).ShowValue=True
Range("C2").FormatConditions(1).ShowIconOnly=False
Range("C2").FormatConditions(1).ReverseOrder=False
ResetRule=ReadRule()
End Function
Public Function ReadRule() As String
ReadRule=CStr(Range("B2").FormatConditions(1).ShowValue) & "|" & CStr(Range("C2").FormatConditions(1).ShowIconOnly) & "|" & CStr(Range("C2").FormatConditions(1).ReverseOrder)
End Function`;
    const before = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(visualConditionBytes), source, 'Change', [], 0, 'visual.xlsx', {});
    if (before.error || before.result?.value !== 'True|False|False;False|True|True;False|True|True') failures.push({case: 'visual condition visibility', setup, phase: 'runtime', result: before.result, error: before.error});
    else {
        const saved = wasm.edit_xlsx_from_workbook(visualConditionBytes, JSON.parse(JSON.stringify(before.workbook)));
        const read = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), source, 'ReadRule', [], 0, 'visual.xlsx', {});
        if (read.error || read.result?.value !== 'False|True|True') failures.push({case: 'visual condition visibility', setup, phase: 'saved', result: read.result, error: read.error});
        const reset = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), source, 'ResetRule', [], 0, 'visual.xlsx', {});
        if (reset.error || reset.result?.value !== 'True|False|False') failures.push({case: 'visual condition visibility', setup, phase: 'reset', result: reset.result, error: reset.error});
        else {
            const savedReset = wasm.edit_xlsx_from_workbook(saved, JSON.parse(JSON.stringify(reset.workbook)));
            const after = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(savedReset), source, 'ReadRule', [], 0, 'visual.xlsx', {});
            if (after.error || after.result?.value !== 'True|False|False') failures.push({case: 'visual condition visibility', setup, phase: 'saved reset', result: after.result, error: after.error});
        }
    }
}
{
    const source = `Public Function Change() As String
Cells.FormatConditions.Delete
Dim p As Object, v As Variant, s As String
Set p=Range("A2:A6").FormatConditions.AddColorScale(3).ColorScaleCriteria(2)
For Each v In Array("3", "3.4", "3.5", 4.5, 0.5, False)
p.Type=v
s=s & CStr(p.Type) & ";"
Next v
On Error Resume Next
p.Type=True: s=s & CStr(Err.Number<>0) & ";": Err.Clear
p.Type=Null: s=s & CStr(Err.Number<>0)
Change=s
End Function`;
    const result = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(visualConditionBytes), source, 'Change', [], 0, 'visual.xlsx', {});
    if (result.error || result.result?.value !== '3;3;4;4;0;0;True;True') failures.push({case: 'scale criterion type coercion', result: result.result, error: result.error});
}
for (const kind of [0, 3, 4, 5]) {
    const source = `Public Function Change() As String
Cells.FormatConditions.Delete
Dim p As Object: Set p=Range("A2:A6").FormatConditions.AddColorScale(3).ColorScaleCriteria(2)
p.Value=Null
p.Type=${kind}
Change=CStr(p.Type) & "|" & CStr(p.Value)
End Function
Public Function ReadRule() As String
Dim p As Object: Set p=Range("A2").FormatConditions(1).ColorScaleCriteria(2)
ReadRule=CStr(p.Type) & "|" & CStr(p.Value)
End Function`;
    const before = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(visualConditionBytes), source, 'Change', [], 0, 'visual.xlsx', {});
    if (before.error || before.result?.value !== `${kind}|=`) failures.push({case: 'null scale type change', kind, phase: 'runtime', result: before.result, error: before.error});
    else {
        const saved = wasm.edit_xlsx_from_workbook(visualConditionBytes, JSON.parse(JSON.stringify(before.workbook)));
        const after = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), source, 'ReadRule', [], 0, 'visual.xlsx', {});
        if (after.error || after.result?.value !== `${kind}|0`) failures.push({case: 'null scale type change', kind, phase: 'saved', result: after.result, error: after.error});
    }
}
for (const [action, address] of [
    ['Rows(1).Insert', 'A3'],
    ['Range("A2:A6").Copy Range("D2")', 'D2'],
    ['c.ModifyAppliesToRange Range("B2:B6")', 'B2'],
]) for (const kind of [0, 3, 4, 5]) {
    const source = `Public Function Change() As String
Cells.FormatConditions.Delete
Dim c As Object
Set c=Range("A2:A6").FormatConditions.AddColorScale(3)
c.ColorScaleCriteria(2).Type=${kind}
c.ColorScaleCriteria(2).Value=Null
${action}
Change=ReadRule()
End Function
Public Function ReadRule() As String
ReadRule=CStr(Range("${address}").FormatConditions(1).ColorScaleCriteria(2).Value)
End Function`;
    const before = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(visualConditionBytes), source, 'Change', [], 0, 'visual.xlsx', {});
    if (before.error || before.result?.value !== '=') failures.push({case: 'null scale structural change', action, kind, phase: 'runtime', result: before.result, error: before.error});
    else {
        const saved = wasm.edit_xlsx_from_workbook(visualConditionBytes, JSON.parse(JSON.stringify(before.workbook)));
        const after = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), source, 'ReadRule', [], 0, 'visual.xlsx', {});
        if (after.error || after.result?.value !== '0') failures.push({case: 'null scale structural change', action, kind, phase: 'saved', result: after.result, error: after.error});
    }
}
{
    const source = `Public Function Change() As String
Cells.FormatConditions.Delete
Dim p As Object, v As Variant, s As String
Set p=Range("A2:A6").FormatConditions.AddColorScale(3).ColorScaleCriteria(2)
p.Value=True: s=CStr(p.Value) & "|"
p.Value=42: s=s & CStr(p.Value) & "|"
p.Value=False: s=s & CStr(p.Value) & "|"
On Error Resume Next
p.Value=v: s=s & CStr(Err.Number) & "|" & CStr(p.Value)
Err.Clear: On Error GoTo 0
Change=s
End Function
Public Function ClearWithNull() As String
Dim p As Object: Set p=Range("A2").FormatConditions(1).ColorScaleCriteria(2)
p.Value=Null
ClearWithNull=CStr(p.Value)
End Function
Public Function ReadRule() As String
ReadRule=CStr(Range("A2").FormatConditions(1).ColorScaleCriteria(2).Value)
End Function`;
    const before = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(visualConditionBytes), source, 'Change', [], 0, 'visual.xlsx', {});
    if (before.error || before.result?.value !== '=TRUE|=TRUE|=FALSE|1004|=FALSE') failures.push({case: 'scale threshold boolean and null', phase: 'runtime', result: before.result, error: before.error});
    else {
        const saved = wasm.edit_xlsx_from_workbook(visualConditionBytes, JSON.parse(JSON.stringify(before.workbook)));
        const read = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), source, 'ReadRule', [], 0, 'visual.xlsx', {});
        if (read.error || read.result?.value !== '=FALSE') failures.push({case: 'scale threshold boolean', phase: 'saved', result: read.result, error: read.error});
        const cleared = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), source, 'ClearWithNull', [], 0, 'visual.xlsx', {});
        if (cleared.error || cleared.result?.value !== '=') failures.push({case: 'scale threshold null', phase: 'runtime', result: cleared.result, error: cleared.error});
        else {
            const savedNull = wasm.edit_xlsx_from_workbook(saved, JSON.parse(JSON.stringify(cleared.workbook)));
            const after = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(savedNull), source, 'ReadRule', [], 0, 'visual.xlsx', {});
            if (after.error || after.result?.value !== '0') failures.push({case: 'scale threshold null', phase: 'saved', result: after.result, error: after.error});
        }
    }
}
{
    const source = `Public Function Change() As String
Cells.FormatConditions.Delete
Dim p As Object, s As String, v As Variant
Set p=Range("A2:A6").FormatConditions.AddColorScale(3).ColorScaleCriteria(2)
On Error Resume Next
For Each v In Array("=SUM(", "=IF(1", "=ABS(1,2", "=SUM(SUM(1)", "=1+SUM(1")
Err.Clear
p.Value=v
s=s & CStr(Err.Number<>0) & ";"
Next v
On Error GoTo 0
p.Value="=sum($B$1"
Change=s & CStr(p.Value)
End Function
Public Function ReadRule() As String
ReadRule=CStr(Range("A2").FormatConditions(1).ColorScaleCriteria(2).Value)
End Function`;
    const before = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(visualConditionBytes), source, 'Change', [], 0, 'visual.xlsx', {});
    if (before.error || before.result?.value !== 'True;True;True;True;True;=SUM($B$1)') failures.push({case: 'scale threshold function completion', phase: 'runtime', result: before.result, error: before.error});
    else {
        const saved = wasm.edit_xlsx_from_workbook(visualConditionBytes, JSON.parse(JSON.stringify(before.workbook)));
        const after = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), source, 'ReadRule', [], 0, 'visual.xlsx', {});
        if (after.error || after.result?.value !== '=SUM($B$1)') failures.push({case: 'scale threshold function completion', phase: 'saved', result: after.result, error: after.error});
    }
}
{
    const source = `Public Function Change() As String
Cells.FormatConditions.Delete
Dim p As Object, s As String, v As Variant
Set p=Range("A2:A6").FormatConditions.AddColorScale(3).ColorScaleCriteria(2)
p.Value="=$B$1"
On Error Resume Next
For Each v In Array("=1+", "=1<", "=1+*2", "=+")
Err.Clear
p.Value=v
s=s & CStr(Err.Number<>0) & "," & CStr(p.Value) & ";"
Next v
On Error GoTo 0
p.Value="=SUM($B$1:$B$2)"
Change=s & CStr(p.Value)
End Function`;
    const result = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(visualConditionBytes), source, 'Change', [], 0, 'visual.xlsx', {});
    if (result.error || result.result?.value !== 'True,=$B$1;True,=$B$1;True,=$B$1;True,=$B$1;=SUM($B$1:$B$2)') failures.push({case: 'scale threshold invalid operators', result: result.result, error: result.error});
}
{
    const source = `Public Function Change() As String
Cells.FormatConditions.Delete
Dim p As Object, s As String
Set p=Range("A2:A6").FormatConditions.AddColorScale(3).ColorScaleCriteria(2)
On Error Resume Next
p.Value=""
s=CStr(Err.Number<>0) & "|"
Err.Clear
On Error GoTo 0
p.Value="=": s=s & CStr(p.Value) & "|"
p.Value="true": s=s & CStr(p.Value) & "|"
p.Value="1e3": s=s & CStr(p.Value)
Change=s
End Function
Public Function ReadRule() As String
ReadRule=CStr(Range("A2").FormatConditions(1).ColorScaleCriteria(2).Value)
End Function`;
    const before = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(visualConditionBytes), source, 'Change', [], 0, 'visual.xlsx', {});
    if (before.error || before.result?.value !== 'True|="="|=TRUE|=1000') failures.push({case: 'scale threshold text coercion', phase: 'runtime', result: before.result, error: before.error});
    else {
        const saved = wasm.edit_xlsx_from_workbook(visualConditionBytes, JSON.parse(JSON.stringify(before.workbook)));
        const after = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), source, 'ReadRule', [], 0, 'visual.xlsx', {});
        if (after.error || after.result?.value !== '1000') failures.push({case: 'scale threshold text coercion', phase: 'saved', result: after.result, error: after.error});
    }
}
for (const [name, setup, address, expected] of [
    ['existing scale threshold', 'Dim c As Object: Set c=Range("A2").FormatConditions(1)', 'A3', '0,=$B$2;3,=$B$2;5,90;'],
    ['new scale threshold', 'Cells.FormatConditions.Delete\nDim c As Object: Set c=Range("A2:A6").FormatConditions.AddColorScale(3)', 'A3', '0,12.5;3,=$B$2;5,90;'],
]) {
    const source = `Public Function Change() As String
${setup}
c.ColorScaleCriteria(1).Type=0
c.ColorScaleCriteria(1).Value=12.5
c.ColorScaleCriteria(2).Type=3
c.ColorScaleCriteria(2).Value="=$B$1"
c.ColorScaleCriteria(3).Type=5
c.ColorScaleCriteria(3).Value=90
Rows(1).Insert
Change=ReadRule()
End Function
Public Function ReadRule() As String
Dim c As Object, p As Object
Set c=Range("${address}").FormatConditions(1)
For Each p In c.ColorScaleCriteria
ReadRule=ReadRule & CStr(p.Type) & "," & CStr(p.Value) & ";"
Next p
End Function`;
    const before = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(visualConditionBytes), source, 'Change', [], 0, 'visual.xlsx', {});
    if (before.error || before.result?.value !== expected) failures.push({case: name, phase: 'runtime', result: before.result, error: before.error});
    else {
        const saved = wasm.edit_xlsx_from_workbook(visualConditionBytes, JSON.parse(JSON.stringify(before.workbook)));
        const after = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), source, 'ReadRule', [], 0, 'visual.xlsx', {});
        if (after.error || after.result?.value !== expected) failures.push({case: name, phase: 'saved', result: after.result, error: after.error});
    }
}
{
    const source = `Public Function Change() As String
Dim c As Object: Set c=Range("A2").FormatConditions(1)
c.ColorScaleCriteria(1).FormatColor.Color=65280
Change=ReadRule()
End Function
Public Function ReadRule() As String
Dim c As Object, p As Object
Set c=Range("A2").FormatConditions(1)
ReadRule=CStr(c.ColorScaleCriteria.Count) & "|"
For Each p In c.ColorScaleCriteria
ReadRule=ReadRule & CStr(p.Type) & "," & CStr(p.Value) & "," & CStr(p.FormatColor.Color) & ";"
Next p
End Function`;
    const expected = '3|4,=$B$1,65280;5,50,8711167;2,0,8109667;';
    const before = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(visualConditionBytes), source, 'Change', [], 0, 'visual.xlsx', {});
    if (before.error || before.result?.value !== expected) failures.push({case: 'color scale criteria', phase: 'runtime', result: before.result, error: before.error});
    else {
        const saved = wasm.edit_xlsx_from_workbook(visualConditionBytes, JSON.parse(JSON.stringify(before.workbook)));
        const after = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), source, 'ReadRule', [], 0, 'visual.xlsx', {});
        if (after.error || after.result?.value !== expected) failures.push({case: 'color scale criteria', phase: 'saved', result: after.result, error: after.error});
    }
}
for (const [caseName, actions, address, expected, active] of [
    ['visual copy', 'Range("A2:A6").Copy Range("E2")\nRange("B2:B6").Copy Range("F2")', 'F2', '4|1|$F$2:$F$6', false],
    ['visual cross-sheet copy', 'Dim original As Object: Set original=ActiveSheet\nWorksheets.Add\noriginal.Range("B2:B6").Copy Range("F2")', 'F2', '4|1|$F$2:$F$6', true],
]) {
    const source = `Public Function Change() As String
${actions}
Change=ReadRule()
End Function
Public Function ReadRule() As String
Dim c As Object: Set c=Range("${address}").FormatConditions(1)
ReadRule=CStr(c.Type) & "|" & CStr(c.Priority) & "|" & c.AppliesTo.Address
End Function`;
    const before = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(visualConditionBytes), source, 'Change', [], 0, 'visual.xlsx', {});
    if (before.error || before.result?.value !== expected) {
        failures.push({case: caseName, phase: 'runtime', result: before.result, error: before.error});
    } else {
        try {
            const saved = wasm.edit_xlsx_from_workbook(visualConditionBytes, JSON.parse(JSON.stringify(before.workbook)));
            const workbook = wasm.parse_spreadsheet(saved);
            const sheet = active ? before.active_sheet : 0;
            const rule = workbook.sheets[sheet].conditional_rules.find(rule => rule.kind === 'dataBar' && rule.ranges.some(range => range[1] === 5));
            if (!rule?.visual_xml?.includes('FF00FF00') || !rule?.visual_extension_xml?.includes('axisPosition="middle"') || !rule?.visual_extension_xml?.includes('<xm:f>$B$1</xm:f>')) {
                failures.push({case: caseName, phase: 'saved visual settings', rule});
            }
            const after = wasm.run_spreadsheet_vba_with_outcome(workbook, source, 'ReadRule', [], sheet, 'visual.xlsx', {});
            if (after.error || after.result?.value !== expected) failures.push({case: caseName, phase: 'saved', result: after.result, error: after.error});
        } catch (error) { failures.push({case: caseName, phase: 'save', error: String(error)}); }
    }
}
{
    const source = 'Public Sub Change()\nRows(1).Insert\nEnd Sub';
    const before = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(visualConditionBytes), source, 'Change', [], 0, 'visual.xlsx', {});
    if (before.error) failures.push({case: 'visual formula threshold shift', phase: 'runtime', error: before.error});
    else {
        try {
            const saved = wasm.edit_xlsx_from_workbook(visualConditionBytes, JSON.parse(JSON.stringify(before.workbook)));
            const rules = wasm.parse_spreadsheet(saved).sheets[0].conditional_rules;
            const scale = rules.find(rule => rule.kind === 'colorScale');
            const bar = rules.find(rule => rule.kind === 'dataBar');
            if (!scale?.visual_xml?.includes('val="$B$2"') || !bar?.visual_xml?.includes('val="$B$2"') || !bar?.visual_extension_xml?.includes('<xm:f>$B$2</xm:f>')) {
                failures.push({case: 'visual formula threshold shift', phase: 'saved', rules});
            }
        } catch (error) { failures.push({case: 'visual formula threshold shift', phase: 'save', error: String(error)}); }
    }
}
for (const [target, expected] of [
    ['B2', '1|0|1|4|$B$3:$B$6|2'],
    ['B2:B6', '1|0|1|0'],
]) {
    const source = `Public Function Change() As String
Range("${target}").FormatConditions.Delete
Change=ReadRule()
End Function
Public Function ReadRule() As String
Dim c As Object
ReadRule=CStr(Range("A2").FormatConditions.Count) & "|" & CStr(Range("B2").FormatConditions.Count) & "|" & CStr(Range("C2").FormatConditions.Count) & "|" & CStr(Range("B3").FormatConditions.Count)
If Range("B3").FormatConditions.Count > 0 Then
Set c=Range("B3").FormatConditions(1)
ReadRule=Left(ReadRule,Len(ReadRule)-1) & CStr(c.Type) & "|" & c.AppliesTo.Address & "|" & CStr(c.Priority)
End If
End Function`;
    const before = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(visualConditionBytes), source, 'Change', [], 0, 'visual.xlsx', {});
    if (before.error || before.result?.value !== expected) {
        failures.push({case: `delete visual condition ${target}`, phase: 'runtime', result: before.result, error: before.error});
    } else {
        try {
            const saved = wasm.edit_xlsx_from_workbook(visualConditionBytes, JSON.parse(JSON.stringify(before.workbook)));
            const after = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), source, 'ReadRule', [], 0, 'visual.xlsx', {});
            if (after.error || after.result?.value !== expected) failures.push({case: `delete visual condition ${target}`, phase: 'saved', result: after.result, error: after.error});
        } catch (error) { failures.push({case: `delete visual condition ${target}`, phase: 'save', error: String(error)}); }
    }
}
const visualConditionSource = `Public Function Change() As String
Rows(3).Insert
Change=ReadRule()
End Function
Public Function ReadRule() As String
Dim a As Object, b As Object, c As Object
Set a=Range("A2").FormatConditions(1): Set b=Range("B2").FormatConditions(1): Set c=Range("C2").FormatConditions(1)
ReadRule=CStr(a.Type) & "|" & a.AppliesTo.Address & "|" & CStr(a.Priority) & ";" & CStr(b.Type) & "|" & b.AppliesTo.Address & "|" & CStr(b.Priority) & ";" & CStr(c.Type) & "|" & c.AppliesTo.Address & "|" & CStr(c.Priority)
End Function`;
const visualExpected = '3|$A$2:$A$7|1;4|$B$2:$B$7|2;6|$C$2:$C$7|3';
const visualBefore = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(visualConditionBytes), visualConditionSource, 'Change', [], 0, 'visual.xlsx', {});
if (visualBefore.error || visualBefore.result?.value !== visualExpected) {
    failures.push({case: 'visual conditional rules', phase: 'runtime', result: visualBefore.result, error: visualBefore.error});
} else {
    try {
        const saved = wasm.edit_xlsx_from_workbook(visualConditionBytes, JSON.parse(JSON.stringify(visualBefore.workbook)));
        const after = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), visualConditionSource, 'ReadRule', [], 0, 'visual.xlsx', {});
        if (after.error || after.result?.value !== visualExpected) failures.push({case: 'visual conditional rules', phase: 'saved', result: after.result, error: after.error});
    } catch (error) { failures.push({case: 'visual conditional rules', phase: 'save', error: String(error)}); }
}
const blankConditionBytes = await readFile(path.join(root, 'tests/fixtures/conditional/blanks.xlsx'));
for (const [action, address, column, row, applies] of [
    ['Rows(1).Delete', 'A1', 'A', 1, '$A$1:$A$5'],
    ['Rows(2).Delete', 'A2', 'A', 2, '$A$2:$A$5'],
    ['Rows(3).Insert', 'A2', 'A', 2, '$A$2:$A$7'],
    ['Columns(1).Insert', 'B2', 'B', 2, '$B$2:$B$6'],
]) {
    const expected = `=LEN(TRIM(${column}${row}))=0|=LEN(TRIM(${column}${row}))>0|${applies}|1,2`;
    const source = `Public Function Change() As String
${action}
Change=ReadRule()
End Function
Public Function ReadRule() As String
Dim a As Object, b As Object
Set a=Range("${address}").FormatConditions(1): Set b=Range("${address}").FormatConditions(2)
ReadRule=a.Formula1 & "|" & b.Formula1 & "|" & a.AppliesTo.Address & "|" & CStr(a.Priority) & "," & CStr(b.Priority)
End Function`;
    const before = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(blankConditionBytes), source, 'Change', [], 0, 'conditions.xlsx', {});
    if (before.error || before.result?.value !== expected) {
        failures.push({case: `blank condition ${action}`, phase: 'runtime', result: before.result, error: before.error});
    } else {
        const saved = wasm.edit_xlsx_from_workbook(blankConditionBytes, JSON.parse(JSON.stringify(before.workbook)));
        const after = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), source, 'ReadRule', [], 0, 'conditions.xlsx', {});
        if (after.error || after.result?.value !== expected) failures.push({case: `blank condition ${action}`, phase: 'saved', result: after.result, error: after.error});
    }
}
const textConditionBytes = await readFile(path.join(root, 'tests/fixtures/conditional/text_rules.xlsx'));
for (const [action, address, column, row, applies] of [
    ['Rows(1).Delete', 'A1', 'A', 1, '$A$1:$A$6'],
    ['Rows(2).Delete', 'A2', 'A', 2, '$A$2:$A$6'],
    ['Rows(3).Insert', 'A2', 'A', 2, '$A$2:$A$8'],
    ['Columns(1).Insert', 'B2', 'B', 2, '$B$2:$B$7'],
    ['Range("A3").Insert xlShiftDown', 'A2', 'A', 2, '$A$2:$A$8'],
]) {
    const expected = `=NOT(ISERROR(SEARCH("apple",${column}${row})))|=LEFT(${column}${row},2)="ba"|=RIGHT(${column}${row},2)="pe"|${applies}|1,2,3`;
    const source = `Public Function Change() As String
${action}
Change=ReadRule()
End Function
Public Function ReadRule() As String
Dim a As Object, b As Object, c As Object
Set a=Range("${address}").FormatConditions(1): Set b=Range("${address}").FormatConditions(2): Set c=Range("${address}").FormatConditions(3)
ReadRule=a.Formula1 & "|" & b.Formula1 & "|" & c.Formula1 & "|" & a.AppliesTo.Address & "|" & CStr(a.Priority) & "," & CStr(b.Priority) & "," & CStr(c.Priority)
End Function`;
    const before = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(textConditionBytes), source, 'Change', [], 0, 'conditions.xlsx', {});
    if (before.error || before.result?.value !== expected) {
        failures.push({case: `text condition ${action}`, phase: 'runtime', result: before.result, error: before.error});
    } else {
        const saved = wasm.edit_xlsx_from_workbook(textConditionBytes, JSON.parse(JSON.stringify(before.workbook)));
        const after = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), source, 'ReadRule', [], 0, 'conditions.xlsx', {});
        if (after.error || after.result?.value !== expected) failures.push({case: `text condition ${action}`, phase: 'saved', result: after.result, error: after.error});
    }
}
const partialConditionBytes = await readFile(path.join(root, 'tests/fixtures/conditional/expression.xlsx'));
{
    const source = `Public Function Change() As String
Cells.FormatConditions.Delete
Dim c As Object
Set c=Range("F2:F6").FormatConditions.AddColorScale(2)
c.ColorScaleCriteria(1).FormatColor.Color=16711680
Change=ReadRule()
End Function
Public Function ReadRule() As String
Dim c As Object, p As Object
Set c=Range("F2").FormatConditions(1)
ReadRule=CStr(c.ColorScaleCriteria.Count) & "|"
For Each p In c.ColorScaleCriteria
ReadRule=ReadRule & CStr(p.Type) & "," & CStr(p.Value) & "," & CStr(p.FormatColor.Color) & ";"
Next p
End Function`;
    const expected = '2|1,0,16711680;2,0,10285055;';
    const before = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(partialConditionBytes), source, 'Change', [], 0, 'conditions.xlsx', {});
    if (before.error || before.result?.value !== expected) failures.push({case: 'new color scale criteria', phase: 'runtime', result: before.result, error: before.error});
    else {
        const saved = wasm.edit_xlsx_from_workbook(partialConditionBytes, JSON.parse(JSON.stringify(before.workbook)));
        const after = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), source, 'ReadRule', [], 0, 'conditions.xlsx', {});
        if (after.error || after.result?.value !== expected) failures.push({case: 'new color scale criteria', phase: 'saved', result: after.result, error: after.error});
    }
}
const newVisualSource = `Public Function Change() As String
Cells.FormatConditions.Delete
Range("A2:A6").FormatConditions.AddColorScale 2
Range("B2:B6").FormatConditions.AddColorScale 3
Range("C2:C6").FormatConditions.AddDatabar
Range("D2:D6").FormatConditions.AddIconSetCondition
Change=ReadRule()
End Function
Public Function ReadRule() As String
Dim c As Object, i As Long
For i=1 To 4
Set c=Cells(2,i).FormatConditions(1)
ReadRule=ReadRule & CStr(c.Type) & "|" & c.AppliesTo.Address & "|" & CStr(c.Priority) & "|" & CStr(c.StopIfTrue) & ";"
Next i
End Function`;
const newVisualExpected = '3|$A$2:$A$6|1|False;3|$B$2:$B$6|2|False;4|$C$2:$C$6|3|False;6|$D$2:$D$6|4|False;';
const newVisualBefore = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(partialConditionBytes), newVisualSource, 'Change', [], 0, 'conditions.xlsx', {});
if (newVisualBefore.error || newVisualBefore.result?.value !== newVisualExpected) {
    failures.push({case: 'new visual conditional rules', phase: 'runtime', result: newVisualBefore.result, error: newVisualBefore.error});
} else {
    try {
        const saved = wasm.edit_xlsx_from_workbook(partialConditionBytes, JSON.parse(JSON.stringify(newVisualBefore.workbook)));
        const after = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), newVisualSource, 'ReadRule', [], 0, 'conditions.xlsx', {});
        if (after.error || after.result?.value !== newVisualExpected) failures.push({case: 'new visual conditional rules', phase: 'saved', result: after.result, error: after.error});
    } catch (error) { failures.push({case: 'new visual conditional rules', phase: 'save', error: String(error)}); }
}
for (const [action, address, column, row, applies] of [
    ['Rows(1).Delete', 'A1', 'A', 1, '$A$1:$A$5'],
    ['Rows(2).Delete', 'A2', 'A', 2, '$A$2:$A$5'],
    ['Rows(3).Insert', 'A2', 'A', 2, '$A$2:$A$7'],
    ['Columns(1).Insert', 'B2', 'B', 2, '$B$2:$B$6'],
]) {
    const expected = `=ISERROR(${column}${row})|=NOT(ISERROR(${column}${row}))|${applies}|1,2`;
    const source = `Public Function Change() As String
Cells.FormatConditions.Delete
Range("A2:A6").FormatConditions.Add 16
Range("A2:A6").FormatConditions.Add 17
${action}
Change=ReadRule()
End Function
Public Function ReadRule() As String
Dim a As Object, b As Object
Set a=Range("${address}").FormatConditions(1): Set b=Range("${address}").FormatConditions(2)
ReadRule=a.Formula1 & "|" & b.Formula1 & "|" & a.AppliesTo.Address & "|" & CStr(a.Priority) & "," & CStr(b.Priority)
End Function`;
    const before = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(partialConditionBytes), source, 'Change', [], 0, 'conditions.xlsx', {});
    if (before.error || before.result?.value !== expected) {
        failures.push({case: `error condition ${action}`, phase: 'runtime', result: before.result, error: before.error});
    } else {
        const saved = wasm.edit_xlsx_from_workbook(partialConditionBytes, JSON.parse(JSON.stringify(before.workbook)));
        const after = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), source, 'ReadRule', [], 0, 'conditions.xlsx', {});
        if (after.error || after.result?.value !== expected) failures.push({case: `error condition ${action}`, phase: 'saved', result: after.result, error: after.error});
    }
}
for (const [action, other, expected] of [
    ['Range("B2:B3").Insert xlShiftToRight', 'G2', '=A2>$E$2|4|=A2>1|5|=F2>0|$F$4,$G$2:$G$3|2'],
    ['Range("B2:B3").Delete xlShiftToLeft', 'E2', '=A2>$C$2|4|=A2>1|5|=E2>0|$F$4,$E$2:$E$3|2'],
]) {
    const source = `Public Function Change() As String
Cells.FormatConditions.Delete
Range("A2:C4").FormatConditions.Add 2,,"=A2>$D$2"
Range("F2:F4").FormatConditions.Add 2,,"=F2>0"
Range("A2:C4").FormatConditions.Add 2,,"=A2>1"
${action}
Change=ReadRule()
End Function
Public Function ReadRule() As String
Dim a As Object, b As Object, c As Object
Set a=Range("A2").FormatConditions(1): Set b=Range("A2").FormatConditions(2): Set c=Range("${other}").FormatConditions(1)
ReadRule=a.Formula1 & "|" & CStr(a.Priority) & "|" & b.Formula1 & "|" & CStr(b.Priority) & "|" & c.Formula1 & "|" & c.AppliesTo.Address & "|" & CStr(c.Priority)
End Function`;
    const before = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(partialConditionBytes), source, 'Change', [], 0, 'conditions.xlsx', {});
    if (before.error || before.result?.value !== expected) {
        failures.push({case: `separate condition ${action}`, phase: 'runtime', result: before.result, error: before.error});
    } else {
        const saved = wasm.edit_xlsx_from_workbook(partialConditionBytes, JSON.parse(JSON.stringify(before.workbook)));
        const after = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), source, 'ReadRule', [], 0, 'conditions.xlsx', {});
        if (after.error || after.result?.value !== expected) failures.push({case: `separate condition ${action}`, phase: 'saved', result: after.result, error: after.error});
    }
}
for (const [action, expected] of [
    ['Range("B2:B3").Insert xlShiftToRight', '=A2>$E$2|3|=A2>1|4|$A$4:$C$4,$A$2:$D$3'],
    ['Range("B2:B3").Delete xlShiftToLeft', '=A2>$C$2|3|=A2>1|4|$A$4:$C$4,$A$2:$B$3'],
]) {
    const source = `Public Function Change() As String
Cells.FormatConditions.Delete
Range("A2:C4").FormatConditions.Add 2,,"=A2>$D$2"
Range("A2:C4").FormatConditions.Add 2,,"=A2>1"
${action}
Change=ReadRule()
End Function
Public Function ReadRule() As String
Dim a As Object, b As Object
Set a=Range("A2").FormatConditions(1): Set b=Range("A2").FormatConditions(2)
ReadRule=a.Formula1 & "|" & CStr(a.Priority) & "|" & b.Formula1 & "|" & CStr(b.Priority) & "|" & a.AppliesTo.Address
End Function`;
    const before = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(partialConditionBytes), source, 'Change', [], 0, 'conditions.xlsx', {});
    if (before.error || before.result?.value !== expected) {
        failures.push({case: `shared condition ${action}`, phase: 'runtime', result: before.result, error: before.error});
    } else {
        const saved = wasm.edit_xlsx_from_workbook(partialConditionBytes, JSON.parse(JSON.stringify(before.workbook)));
        const after = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), source, 'ReadRule', [], 0, 'conditions.xlsx', {});
        if (after.error || after.result?.value !== expected) failures.push({case: `shared condition ${action}`, phase: 'saved', result: after.result, error: after.error});
    }
}
for (const [action, expected] of [
    ['Range("B3").Insert xlShiftToRight', '=A2>$D$2|$A$2:$C$2,$A$4:$C$4,$A$3:$D$3|1'],
    ['Range("B3").Delete xlShiftToLeft', '=A2>$D$2|$A$2:$C$2,$A$4:$C$4,$A$3:$B$3|1'],
    ['Range("B2:B3").Insert xlShiftToRight', '=A2>$E$2|$A$4:$C$4,$A$2:$D$3|2'],
    ['Range("B2:B3").Delete xlShiftToLeft', '=A2>$C$2|$A$4:$C$4,$A$2:$B$3|2'],
    ['Range("C3").Delete xlShiftToLeft', '=A2>$D$2|$A$2:$C$4|1'],
    ['Range("B3:C3").Delete xlShiftToLeft', '=A2>$D$2|$A$2:$C$4|1'],
    ['Range("A4").Delete xlShiftUp', '=A2>$D$2|$A$2:$C$4|1'],
    ['Range("A3:A4").Delete xlShiftUp', '=A2>$D$2|$A$2:$C$4|1'],
]) {
    const source = `Public Function Change() As String
Cells.FormatConditions.Delete
Range("A2:C4").FormatConditions.Add 2,,"=A2>$D$2"
${action}
Change=ReadRule()
End Function
Public Function ReadRule() As String
Dim c As Object: Set c=Range("A2").FormatConditions(1)
ReadRule=c.Formula1 & "|" & c.AppliesTo.Address & "|" & CStr(c.Priority)
End Function`;
    const before = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(partialConditionBytes), source, 'Change', [], 0, 'conditions.xlsx', {});
    if (before.error || before.result?.value !== expected) {
        failures.push({case: `horizontal condition ${action}`, phase: 'runtime', result: before.result, error: before.error});
    } else {
        const saved = wasm.edit_xlsx_from_workbook(partialConditionBytes, JSON.parse(JSON.stringify(before.workbook)));
        const after = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), source, 'ReadRule', [], 0, 'conditions.xlsx', {});
        if (after.error || after.result?.value !== expected) failures.push({case: `horizontal condition ${action}`, phase: 'saved', result: after.result, error: after.error});
    }
}
for (const [action, expected] of [
    ['Range("A3").Insert xlShiftDown', '=A2>$C$2|$A$2:$B$2,$B$3:$B$5,$A$3:$A$6|1'],
    ['Range("A3").Delete xlShiftUp', '=A2>$C$2|$A$2:$B$2,$B$3:$B$5,$A$3:$A$4|1'],
]) {
    const source = `Public Function Change() As String
Cells.FormatConditions.Delete
Range("A2:B5").FormatConditions.Add 2,,"=A2>$C$2"
${action}
Change=ReadRule()
End Function
Public Function ReadRule() As String
Dim c As Object: Set c=Range("B2").FormatConditions(1)
ReadRule=c.Formula1 & "|" & c.AppliesTo.Address & "|" & CStr(c.Priority)
End Function`;
    const before = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(partialConditionBytes), source, 'Change', [], 0, 'conditions.xlsx', {});
    if (before.error || before.result?.value !== expected) {
        failures.push({case: `partial condition ${action}`, phase: 'runtime', result: before.result, error: before.error});
    } else {
        const saved = wasm.edit_xlsx_from_workbook(partialConditionBytes, JSON.parse(JSON.stringify(before.workbook)));
        const after = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), source, 'ReadRule', [], 0, 'conditions.xlsx', {});
        if (after.error || after.result?.value !== expected) failures.push({case: `partial condition ${action}`, phase: 'saved', result: after.result, error: after.error});
    }
}
for (const [action, expected] of [
    ['Rows(3).Insert', '=A2>$B$2|$A$2:$A$6|1|13551615'],
    ['Rows(2).Delete', '=A2>#REF!|$A$2:$A$4|3|13551615'],
    ['Columns(2).Insert', '=A2>$C$2|$A$2:$B$5|1|13551615'],
    ['Range("A3").Insert xlShiftDown', '=A2>$B$2|$A$2:$A$6|1|13551615'],
    ['Range("A3").Delete xlShiftUp', '=A2>$B$2|$A$2:$A$4|1|13551615'],
    ['Range("A2").Delete xlShiftUp', '=A2>$B$2|$A$2:$A$4|1|13551615'],
]) {
    const source = `Public Function Change() As String
${action}
Change=ReadRule()
End Function
Public Function ReadRule() As String
Dim c As Object: Set c=Range("A2").FormatConditions(1)
ReadRule=c.Formula1 & "|" & c.AppliesTo.Address & "|" & CStr(c.Priority) & "|" & CStr(c.Interior.Color)
End Function`;
    const before = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(partialConditionBytes), source, 'Change', [], 0, 'conditions.xlsx', {});
    if (before.error || before.result?.value !== expected) {
        failures.push({case: action, phase: 'runtime', result: before.result, error: before.error});
    } else {
        const saved = wasm.edit_xlsx_from_workbook(partialConditionBytes, JSON.parse(JSON.stringify(before.workbook)));
        const after = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), source, 'ReadRule', [], 0, 'conditions.xlsx', {});
        if (after.error || after.result?.value !== expected) failures.push({case: action, phase: 'saved', result: after.result, error: after.error});
    }
}
for (const [name, action, expected] of [
    ['cut_condition_all', 'Range("A2:A5").Cut Range("C3")', '0|1|=C3>$B$2|$C$3:$C$6|1|13551615'],
    ['cut_condition_part', 'Range("A3:A4").Cut Range("C3")', '1|1|=A2>$B$2|$A$2,$A$5,$C$3:$C$4|1|13551615'],
]) {
    const source = `Public Function Change() As String
${action}
Change=ReadRule()
End Function
Public Function ReadRule() As String
Dim c As Object
Set c=Range("C3").FormatConditions(1)
ReadRule=CStr(Range("A2").FormatConditions.Count) & "|" & CStr(Range("C3").FormatConditions.Count) & "|" & c.Formula1 & "|" & c.AppliesTo.Address & "|" & CStr(c.Priority) & "|" & CStr(c.Interior.Color)
End Function`;
    const before = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(partialConditionBytes), source, 'Change', [], 0, 'conditions.xlsx', {});
    if (before.error || before.result?.value !== expected) {
        failures.push({case: name, phase: 'runtime', result: before.result, error: before.error});
    } else {
        const saved = wasm.edit_xlsx_from_workbook(partialConditionBytes, JSON.parse(JSON.stringify(before.workbook)));
        const after = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), source, 'ReadRule', [], 0, 'conditions.xlsx', {});
        const reopenedExpected = expected.replace('$A$2,$A$5,$C$3:$C$4', '$A$2,$C$3:$C$4,$A$5');
        if (after.error || after.result?.value !== reopenedExpected) failures.push({case: name, phase: 'saved', result: after.result, error: after.error});
    }
}
const cutReferenceSource = `Public Function Change() As String
Range("B2").Cut Range("D2")
Change=ReadRule()
End Function
Public Function ReadRule() As String
Dim a As Object, b As Object
Set a=Range("A2").FormatConditions(1): Set b=Range("D2").FormatConditions(1)
ReadRule=a.Formula1 & "|" & CStr(a.Priority) & "|" & b.Formula1 & "|" & b.AppliesTo.Address & "|" & CStr(b.Priority)
End Function`;
const cutReferenceExpected = '=A2>$D$2|3|=MOD(ROW(),2)=0|$B$3:$B$5,$D$2|2';
const cutReferenceBefore = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(partialConditionBytes), cutReferenceSource, 'Change', [], 0, 'conditions.xlsx', {});
if (cutReferenceBefore.error || cutReferenceBefore.result?.value !== cutReferenceExpected) {
    failures.push({case: 'cut_conditional_reference', phase: 'runtime', result: cutReferenceBefore.result, error: cutReferenceBefore.error});
} else {
    const saved = wasm.edit_xlsx_from_workbook(partialConditionBytes, JSON.parse(JSON.stringify(cutReferenceBefore.workbook)));
    const after = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), cutReferenceSource, 'ReadRule', [], 0, 'conditions.xlsx', {});
    const reopenedExpected = cutReferenceExpected.replace('$B$3:$B$5,$D$2', '$D$2,$B$3:$B$5');
    if (after.error || after.result?.value !== reopenedExpected) failures.push({case: 'cut_conditional_reference', phase: 'saved', result: after.result, error: after.error});
}
for (const [name, action, target, expected] of [
    ['paste_formats', 'Range("A2:A5").Copy: Range("C3:C10").PasteSpecial -4122', 'C3:C10', '1|=C3>$B$2|$C$3:$C$10|1|13551615'],
    ['copy_destination', 'Range("A2:A5").Copy Range("C3:C10")', 'C3:C10', '1|=C3>$B$2|$C$3:$C$10|1|13551615'],
    ['paste_transposed', 'Range("A2:A5").Copy: Range("C3").PasteSpecial -4122,-4142,False,True', 'C3:F3', '1|=C3>$B$2|$C$3:$F$3|1|13551615'],
]) {
    const source = `Public Function Change() As String
${action}
Change=ReadRule()
End Function
Public Function ReadRule() As String
Dim c As Object
Set c=Range("C3").FormatConditions(1)
ReadRule=CStr(Range("${target}").FormatConditions.Count) & "|" & c.Formula1 & "|" & c.AppliesTo.Address & "|" & CStr(c.Priority) & "|" & CStr(c.Interior.Color)
End Function`;
    const before = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(partialConditionBytes), source, 'Change', [], 0, 'conditions.xlsx', {});
    if (before.error || before.result?.value !== expected) {
        failures.push({case: name, phase: 'runtime', result: before.result, error: before.error});
    } else {
        const saved = wasm.edit_xlsx_from_workbook(partialConditionBytes, JSON.parse(JSON.stringify(before.workbook)));
        const after = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), source, 'ReadRule', [], 0, 'conditions.xlsx', {});
        if (after.error || after.result?.value !== expected) failures.push({case: name, phase: 'saved', result: after.result, error: after.error});
    }
}
const modifyRangeSource = `Public Function Change() As String
Dim c As Object
Set c=Range("A2").FormatConditions(1)
ThisWorkbook.Saved=True
c.ModifyAppliesToRange Range("A2:A5")
Dim wasSaved As Boolean: wasSaved=ThisWorkbook.Saved
c.ModifyAppliesToRange Union(Range("E2:E4"),Range("C5:C6"))
Change=CStr(wasSaved) & "|" & ReadRule()
End Function
Public Function ReadRule() As String
Dim c As Object
Set c=Range("E2").FormatConditions(1)
ReadRule=c.Formula1 & "|" & c.AppliesTo.Address & "|" & CStr(c.Priority) & "|" & CStr(c.Interior.Color) & "|" & CStr(Range("A2").FormatConditions.Count)
End Function`;
const modifyRangeExpected = '=C2>$B$2|$E$2:$E$4,$C$5:$C$6|1|13551615|0';
const modifyRangeBefore = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(partialConditionBytes), modifyRangeSource, 'Change', [], 0, 'conditions.xlsx', {});
if (modifyRangeBefore.error || modifyRangeBefore.result?.value !== `True|${modifyRangeExpected}`) {
    failures.push({case: 'modify_conditional_range', phase: 'runtime', result: modifyRangeBefore.result, error: modifyRangeBefore.error});
} else {
    const saved = wasm.edit_xlsx_from_workbook(partialConditionBytes, JSON.parse(JSON.stringify(modifyRangeBefore.workbook)));
    const after = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), modifyRangeSource, 'ReadRule', [], 0, 'conditions.xlsx', {});
    if (after.error || after.result?.value !== modifyRangeExpected) failures.push({case: 'modify_conditional_range', phase: 'saved', result: after.result, error: after.error});
}
const partialConditionSource = `Public Function Remove() As String
Range("A2").FormatConditions.Delete
Range("A4").FormatConditions.Delete
Remove=ReadRule()
End Function
Public Function ReadRule() As String
Dim c As Object
Set c=Range("A3").FormatConditions(1)
ReadRule=c.Formula1 & "|" & c.AppliesTo.Address & "|" & CStr(Range("A3").FormatConditions.Count) & "|" & CStr(Range("A4").FormatConditions.Count)
End Function`;
const partialConditionExpected = '=A3>$B$2|$A$3,$A$5|1|0';
const partialBefore = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(partialConditionBytes), partialConditionSource, 'Remove', [], 0, 'conditions.xlsx', {});
if (partialBefore.error || partialBefore.result?.value !== partialConditionExpected) {
    failures.push({case: 'partial_condition', phase: 'runtime', result: partialBefore.result, error: partialBefore.error});
} else {
    const saved = wasm.edit_xlsx_from_workbook(partialConditionBytes, JSON.parse(JSON.stringify(partialBefore.workbook)));
    const after = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), partialConditionSource, 'ReadRule', [], 0, 'conditions.xlsx', {});
    if (after.error || after.result?.value !== partialConditionExpected) failures.push({case: 'partial_condition', phase: 'saved', result: after.result, error: after.error});
}
for (const addSheet of [false, true]) {
    const source = `Public Function Change() As String
${addSheet ? 'Worksheets.Add' : ''}
Range("C1").Value=10: Range("C2").Value=30
Dim c As Object
Set c=Range("C1:C2").FormatConditions.Add(1,5,20)
c.Interior.Color=255: c.Font.Bold=True: c.Font.Strikethrough=True
Change=ReadRule()
End Function
Public Function ReadRule() As String
Dim c As Object
Set c=Range("C1").FormatConditions(1)
ReadRule=CStr(Range("C1").DisplayFormat.Interior.Color) & "|" & CStr(Range("C2").DisplayFormat.Interior.Color) & "|" & CStr(Range("C2").DisplayFormat.Font.Bold) & "|" & CStr(Range("C2").DisplayFormat.Font.Strikethrough) & "|" & c.Formula1 & "|" & CStr(c.StopIfTrue)
End Function`;
    const expected = '16777215|255|True|True|=20|True';
    const before = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(loadedConditionBytes), source, 'Change', [], 0, 'conditions.xlsx', {});
    if (before.error || before.result?.value !== expected) {
        failures.push({case: `new_condition_${addSheet}`, phase: 'runtime', result: before.result, error: before.error});
    } else {
        const saved = wasm.edit_xlsx_from_workbook(loadedConditionBytes, JSON.parse(JSON.stringify(before.workbook)));
        const after = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), source, 'ReadRule', [], before.active_sheet, 'conditions.xlsx', {});
        if (after.error || after.result?.value !== expected) failures.push({case: `new_condition_${addSheet}`, phase: 'saved', result: after.result, error: after.error});
    }
}
for (const [fixture, expected] of [
    ['top10.xlsx', '13561798|16777215|13561798|13551615'],
    ['duplicates.xlsx', '13551615|13551615|13551615|13561798'],
]) {
    const bytes = await readFile(path.join(root, 'tests/fixtures/conditional', fixture));
    const before = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(bytes), conditionColorSource, 'Probe', [], 0, fixture, {});
    if (before.error || before.result?.value !== expected) {
        failures.push({case: fixture, phase: 'runtime', result: before.result, error: before.error});
    } else {
        const saved = wasm.edit_xlsx_from_workbook(bytes, JSON.parse(JSON.stringify(before.workbook)));
        const after = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(saved), conditionColorSource, 'Probe', [], 0, fixture, {});
        if (after.error || after.result?.value !== expected) failures.push({case: fixture, phase: 'saved', result: after.result, error: after.error});
    }
}
const aboveAverageSource = `Public Function Probe() As String
Worksheets.Add
Range("A1:A3").Value=2
Dim c As Object
Set c=Range("A1:A3").FormatConditions.AddAboveAverage()
c.Interior.Color=255
Probe=CStr(c.AboveBelow) & "|" & CStr(c.NumStdDev) & "|"
Dim i As Long
For i=0 To 5: c.AboveBelow=i: Probe=Probe & CStr(Range("A1").DisplayFormat.Interior.Color) & "|": Next i
End Function`;
const aboveAverage = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(loadedConditionBytes), aboveAverageSource, 'Probe', [], 0, 'conditions.xlsx', {});
if (aboveAverage.error || aboveAverage.result?.value !== '0|1|16777215|16777215|255|255|255|255|') {
    failures.push({case: 'above_average', phase: 'runtime', result: aboveAverage.result, error: aboveAverage.error});
}
const conditionDirtySource = `Public Function Probe() As String
ThisWorkbook.Saved=True
Range("Z1").FormatConditions.Delete
Probe=CStr(ThisWorkbook.Saved) & "|"
Dim c As Object
Set c=Range("Z1:Z3").FormatConditions.Add(1,5,2)
Probe=Probe & CStr(ThisWorkbook.Saved) & "|"
ThisWorkbook.Saved=True: c.StopIfTrue=c.StopIfTrue
Probe=Probe & CStr(ThisWorkbook.Saved) & "|"
ThisWorkbook.Saved=True: c.Interior.Color=255
Probe=Probe & CStr(ThisWorkbook.Saved) & "|"
ThisWorkbook.Saved=True: c.Delete
Probe=Probe & CStr(ThisWorkbook.Saved)
End Function`;
const conditionDirty = wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(loadedConditionBytes), conditionDirtySource, 'Probe', [], 0, 'conditions.xlsx', {});
if (conditionDirty.error || conditionDirty.result?.value !== 'True|False|False|False|False' || !conditionDirty.workbook_changed) {
    failures.push({case: 'conditional_format_dirty', phase: 'runtime', result: conditionDirty.result, error: conditionDirty.error, changed: conditionDirty.workbook_changed});
}
assert.deepEqual(failures, [], 'Excel saves validation and inherited row/column formatting');
console.log(JSON.stringify({passed: true, native_structural_save_cases: 4, native_existing_validation_cases: 1, new_sheet_save_cases: 1, save_copy_band_cases: 1, native_band_intersections: 6, native_clear_intersections: 7, native_sheet_format_clears: 3, native_sheet_format_cross: 1, native_relative_validation_cases: 3, native_validation_shift_cases: 10, native_validation_two_bounds: 4, native_validation_value_cases: 1, native_validation_cross_sheet: 1, multi_validation_cases: 15, multi_validation_shift_cases: 5, workbook_json_roundtrip: true}));
