// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { Worker as ThreadWorker } from 'node:worker_threads';
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const pkg = path.resolve(process.argv[2] || path.join(root, 'crates/oxidocs-wasm/pkg'));
const bindingUrl = 'data:text/javascript;base64,' + Buffer.from(await readFile(path.join(pkg, 'oxidocs_wasm.js'))).toString('base64');
const wasm = await import(bindingUrl);
await wasm.default({ module_or_path: await readFile(path.join(pkg, 'oxidocs_wasm_bg.wasm')) });
const cases = JSON.parse(await readFile(path.join(root, 'tests/fixtures/vba/storage_types.json'), 'utf8'));
const emptyBook = () => ({ sheets: [{ name: 'Sheet1', rows: [], col_count: 0, col_widths: [], default_col_width: 0, default_row_height: 15, merge_cells: [] }] });
async function sequence(testCase,invoke,worker=false) {
  const source=await readFile(path.join(root,'tests/fixtures/vba',testCase.source),'utf8');
  let book=emptyBook();
  for (const step of testCase.outcome==='compile'?[null,null]:testCase.steps) {
    if (testCase.outcome==='compile' && worker) {
      await assert.rejects(invoke(book,source,testCase.procedure),error=>error.message.includes('コンパイル エラー: '+testCase.message),testCase.name);
      continue;
    }
    const result=await invoke(book,source,testCase.procedure);
    if (testCase.outcome==='compile') {
      assert.equal(result.error.kind,'compile',testCase.name);
      assert.ok(result.error.message.endsWith('コンパイル エラー: '+testCase.message),testCase.name);
    } else {
      assert.equal(result.error,undefined,testCase.name);
      assert.deepEqual(result.result,{type:'string',value:step.value},testCase.name);
    }
    book=result.workbook;
  }
}
for (const item of cases) {
  const session=new wasm.SpreadsheetVbaSession();
  try {await sequence(item,(book,source,procedure)=>session.run(book,source,procedure,[],0,'Compile.xlsm',{errorLanguage:'ja-JP'}));}
  finally {session.free();}
}
const shim = `import init, { SpreadsheetVbaSession, spreadsheet_vba_project_name, spreadsheet_vba_source } from ${JSON.stringify(bindingUrl)};
import { readFile } from 'node:fs/promises';
await init({ module_or_path: await readFile(${JSON.stringify(path.join(pkg, 'oxidocs_wasm_bg.wasm'))}) });
export { SpreadsheetVbaSession, spreadsheet_vba_project_name, spreadsheet_vba_source };
export default function ready() { return Promise.resolve(); }`;
const workerSource = (await readFile(path.join(root, 'web/vba-worker.js'), 'utf8')).replace("'./oxidocs_wasm.js'", JSON.stringify('data:text/javascript;base64,' + Buffer.from(shim).toString('base64')));
const bridge = `import { parentPort, workerData } from 'node:worker_threads';
const queued = []; let listener;
globalThis.self = { postMessage: value => parentPort.postMessage(value), addEventListener: (type, handler) => { listener = handler; for (const data of queued) listener({ data }); } };
parentPort.on('message', data => listener ? listener({ data }) : queued.push(data));
await import('data:text/javascript;base64,' + Buffer.from(workerData).toString('base64'));`;
class BrowserWorker {
  constructor() { this.thread = new ThreadWorker(new URL('data:text/javascript;base64,' + Buffer.from(bridge).toString('base64')), { workerData: workerSource }); }
  addEventListener(type, handler) { this.thread.on(type, value => handler(type === 'message' ? { data: value } : value)); }
  postMessage(value) { this.thread.postMessage(value); }
  terminate() { void this.thread.terminate(); }
}
const runner = await import(pathToFileURL(path.join(root, 'web/vba-runner.js')));
const priorWorker = globalThis.Worker;
try {
  globalThis.Worker = BrowserWorker;
  for (const item of cases) {
    runner.closeVbaProject();
    await sequence(item,(book,source,procedure)=>runner.runVbaInWorker(book,source,procedure,[],0,'Compile.xlsm',{applicationState:{},errorLanguage:'ja-JP'}),true);
  }
} finally { runner.closeVbaProject(); globalThis.Worker = priorWorker; }

const resultBoundarySource = "Public Type BoundaryRecord\n SafeValue As LongLong\n LargeValue As LongLong\nEnd Type\nPrivate retainedBoundaryObject As Collection\nPublic Function BoundarySmall() As LongLong\n BoundarySmall = CLngLng(42)\nEnd Function\nPublic Function BoundarySafe() As LongLong\n BoundarySafe = CLngLng(4503599627370496#) * CLngLng(2) - CLngLng(1)\nEnd Function\nPublic Function BoundaryUnsafe() As LongLong\n BoundaryUnsafe = CLngLng(4503599627370496#) * CLngLng(2)\nEnd Function\nPublic Function BoundaryAdjacent() As LongLong\n BoundaryAdjacent = BoundaryUnsafe() + CLngLng(1)\nEnd Function\nPublic Function BoundaryNegativeSafe() As LongLong\n BoundaryNegativeSafe = -BoundarySafe()\nEnd Function\nPublic Function BoundaryNegativeUnsafe() As LongLong\n BoundaryNegativeUnsafe = -BoundaryUnsafe()\nEnd Function\nPublic Function BoundaryMinimum() As LongLong\n BoundaryMinimum = -CLngLng(4611686018427387904#) - CLngLng(4611686018427387904#)\nEnd Function\nPublic Function BoundaryMaximum() As LongLong\n BoundaryMaximum = CLngLng(4611686018427387904#) - CLngLng(1) + CLngLng(4611686018427387904#)\nEnd Function\nPublic Function BoundaryDouble() As Double\n BoundaryDouble = 9007199254740992#\nEnd Function\nPublic Function BoundaryWorkbook() As Object\n Set BoundaryWorkbook = ThisWorkbook\nEnd Function\nPublic Function BoundaryCollection() As Object\n Set BoundaryCollection = New Collection\nEnd Function\nPublic Function BoundaryRetained() As Object\n If retainedBoundaryObject Is Nothing Then Set retainedBoundaryObject = New Collection\n Set BoundaryRetained = retainedBoundaryObject\nEnd Function\nPublic Function BoundaryArray() As Variant\n Dim values(0 To 3) As Variant\n values(0) = BoundaryUnsafe()\n Set values(1) = New Collection\n Set values(2) = New Collection\n values(3) = CLngLng(42)\n BoundaryArray = values\nEnd Function\n\nPublic Function BoundaryRecordResult() As BoundaryRecord\n Dim entry As BoundaryRecord\n entry.SafeValue = CLngLng(42)\n entry.LargeValue = BoundaryUnsafe()\n BoundaryRecordResult = entry\nEnd Function\nPublic Function DescribeBoundaryRecord() As String\n Dim entry As BoundaryRecord\n entry = BoundaryRecordResult()\n DescribeBoundaryRecord = CStr(entry.SafeValue) & \"|\" & CStr(entry.LargeValue)\nEnd Function\n";
async function verifyResultBoundaries(invoke) {
  let book = emptyBook();
  async function run(procedure) {
    const execution = await invoke(book, resultBoundarySource, procedure);
    assert.equal(execution.error, undefined, procedure);
    book = execution.workbook;
    assert.deepEqual(JSON.parse(JSON.stringify(execution.result)), execution.result,
      procedure + ': result survives JSON transport without precision loss');
    return execution.result;
  }
  const integerCases = [
    ['BoundarySmall', 42],
    ['BoundarySafe', 9007199254740991],
    ['BoundaryUnsafe', '9007199254740992'],
    ['BoundaryAdjacent', '9007199254740993'],
    ['BoundaryNegativeSafe', -9007199254740991],
    ['BoundaryNegativeUnsafe', '-9007199254740992'],
    ['BoundaryMinimum', '-9223372036854775808'],
    ['BoundaryMaximum', '9223372036854775807'],
  ];
  for (const [procedure, value] of integerCases) {
    assert.deepEqual(await run(procedure), {type:'integer', value}, procedure);
  }
  assert.deepEqual(await run('BoundaryDouble'), {type:'double', value:9007199254740992});
  const workbook = await run('BoundaryWorkbook');
  assert.equal(workbook.type, 'object');
  assert.equal(workbook.value.kind, 'Workbook');
  assert.equal(typeof workbook.value.handle, 'number');
  function collection(value) {
    assert.equal(value.type, 'object');
    assert.equal(value.value.kind, 'Collection');
    assert.equal(typeof value.value.handle, 'string');
    assert.ok(BigInt(value.value.handle) >= (1n << 63n));
    return value.value.handle;
  }
  collection(await run('BoundaryCollection'));
  const retained = collection(await run('BoundaryRetained'));
  assert.equal(collection(await run('BoundaryRetained')), retained,
    'the retained VBA object keeps its identity between invocations');
  const array = await run('BoundaryArray');
  assert.equal(array.type, 'array');
  assert.equal(array.value.lower_bound, 0);
  assert.deepEqual(array.value.values[0], {type:'integer',value:'9007199254740992'});
  assert.notEqual(collection(array.value.values[1]), collection(array.value.values[2]),
    'distinct runtime objects keep distinct 64-bit handles');
  assert.deepEqual(array.value.values[3], {type:'integer',value:42});
  const record = await run('BoundaryRecordResult');
  assert.equal(record.type, 'record');
  assert.match(record.value.type_name, /(^|\.)BoundaryRecord$/i);
  const fields = new Map(record.value.fields.map(([name, value]) => [name.toLowerCase(), value]));
  assert.deepEqual(fields.get('safevalue'), {type:'integer',value:42});
  assert.deepEqual(fields.get('largevalue'), {type:'integer',value:'9007199254740992'});
  assert.deepEqual(await run('DescribeBoundaryRecord'), {type:'string',value:'42|9007199254740992'});
}
{
  const session = new wasm.SpreadsheetVbaSession();
  try {
    await verifyResultBoundaries((book, source, procedure) =>
      session.run(book, source, procedure, [], 0, 'ResultBoundary.xlsm', {errorLanguage:'ja-JP'}));
  } finally { session.free(); }
}
try {
  globalThis.Worker = BrowserWorker;
  runner.closeVbaProject();
  await verifyResultBoundaries((book, source, procedure) =>
    runner.runVbaInWorker(book, source, procedure, [], 0, 'ResultBoundary.xlsm',
      {applicationState:{},errorLanguage:'ja-JP'}));
} finally { runner.closeVbaProject(); globalThis.Worker = priorWorker; }


const exactNumericSource = "Public Type ExactNumericRecord\n CurrencyValue As Currency\n DecimalValue As Variant\nEnd Type\nPublic Function CurrencySmall() As Currency\n CurrencySmall = CCur(\"12.3456\")\nEnd Function\nPublic Function CurrencyMaximum() As Currency\n CurrencyMaximum = CCur(\"922337203685477.5807\")\nEnd Function\nPublic Function CurrencyMinimum() As Currency\n CurrencyMinimum = CCur(\"-922337203685477.5808\")\nEnd Function\nPublic Function DecimalTenth() As Variant\n DecimalTenth = CDec(\"0.1\")\nEnd Function\nPublic Function DecimalMaximum() As Variant\n DecimalMaximum = CDec(\"79228162514264337593543950335\")\nEnd Function\nPublic Function DecimalScaledMaximum() As Variant\n DecimalScaledMaximum = CDec(\"7922816251426433759354395033.5\")\nEnd Function\nPublic Function DescribeNumericTypes() As String\n DescribeNumericTypes = TypeName(CurrencyMaximum()) & \":\" & CStr(CurrencyMaximum()) & \"|\" & TypeName(DecimalMaximum()) & \":\" & CStr(DecimalMaximum()) & \"|\" & CStr(DecimalScaledMaximum())\nEnd Function\nPublic Function ExactNumericArray() As Variant\n ExactNumericArray = Array(CurrencyMaximum(), DecimalScaledMaximum(), CLngLng(42), 9007199254740992#)\nEnd Function\nPublic Function ExactNumericRecordResult() As ExactNumericRecord\n Dim entry As ExactNumericRecord\n entry.CurrencyValue = CurrencyMinimum()\n entry.DecimalValue = DecimalMaximum()\n ExactNumericRecordResult = entry\nEnd Function\nPublic Function DescribeExactNumericComposite() As String\n Dim entry As ExactNumericRecord\n entry = ExactNumericRecordResult()\n Dim values As Variant\n values = ExactNumericArray()\n DescribeExactNumericComposite = CStr(entry.CurrencyValue) & \"|\" & TypeName(entry.CurrencyValue) & \"|\" & CStr(entry.DecimalValue) & \"|\" & TypeName(entry.DecimalValue) & \"|\" & CStr(values(0)) & \"|\" & TypeName(values(0)) & \"|\" & CStr(values(1)) & \"|\" & TypeName(values(1))\nEnd Function\n";
const exactNumericCases = [["CurrencySmall", {"type": "currency", "value": "12.3456"}], ["CurrencyMaximum", {"type": "currency", "value": "922337203685477.5807"}], ["CurrencyMinimum", {"type": "currency", "value": "-922337203685477.5808"}], ["DecimalTenth", {"type": "decimal", "value": "0.1"}], ["DecimalMaximum", {"type": "decimal", "value": "79228162514264337593543950335"}], ["DecimalScaledMaximum", {"type": "decimal", "value": "7922816251426433759354395033.5"}]];
const exactNumericDescriptor = "Currency:922337203685477.5807|Decimal:79228162514264337593543950335|7922816251426433759354395033.5";
const exactCompositeDescriptor = "-922337203685477.5808|Currency|79228162514264337593543950335|Decimal|922337203685477.5807|Currency|7922816251426433759354395033.5|Decimal";

async function verifyExactNumericResults(invoke) {
  let book=emptyBook();
  async function run(procedure) {
    const execution=await invoke(book,exactNumericSource,procedure);
    assert.equal(execution.error,undefined,procedure);
    book=execution.workbook;
    assert.deepEqual(JSON.parse(JSON.stringify(execution.result)),execution.result,
      procedure+': exact decimal result survives JSON');
    return execution.result;
  }
  for(const [procedure,expected] of exactNumericCases) assert.deepEqual(await run(procedure),expected,procedure);
  assert.deepEqual(await run('DescribeNumericTypes'),{type:'string',value:exactNumericDescriptor});
  const array=await run('ExactNumericArray');
  assert.equal(array.type,'array');
  assert.deepEqual(array.value.values,[
    {type:'currency',value:'922337203685477.5807'},
    {type:'decimal',value:'7922816251426433759354395033.5'},
    {type:'integer',value:42},
    {type:'double',value:9007199254740992},
  ]);
  const record=await run('ExactNumericRecordResult');
  assert.equal(record.type,'record');
  assert.match(record.value.type_name,/(^|\.)ExactNumericRecord$/i);
  const fields=new Map(record.value.fields.map(([name,value])=>[name.toLowerCase(),value]));
  assert.deepEqual(fields.get('currencyvalue'),{type:'currency',value:'-922337203685477.5808'});
  assert.deepEqual(fields.get('decimalvalue'),{type:'decimal',value:'79228162514264337593543950335'});
  assert.deepEqual(await run('DescribeExactNumericComposite'),{type:'string',value:exactCompositeDescriptor});
}
{
  const session=new wasm.SpreadsheetVbaSession();
  try {
    await verifyExactNumericResults((book,source,procedure)=>
      session.run(book,source,procedure,[],0,'ExactNumeric.xlsm',{errorLanguage:'ja-JP'}));
  } finally {session.free();}
}
try {
  globalThis.Worker=BrowserWorker;
  runner.closeVbaProject();
  await verifyExactNumericResults((book,source,procedure)=>
    runner.runVbaInWorker(book,source,procedure,[],0,'ExactNumeric.xlsm',
      {applicationState:{},errorLanguage:'ja-JP'}));
} finally {runner.closeVbaProject();globalThis.Worker=priorWorker;}


const nativeRawInputSource = "Public Function EchoArgument(ByVal value As Variant) As Variant\n EchoArgument = value\nEnd Function\nPublic Function DescribeArgument(ByVal value As Variant) As String\n DescribeArgument = TypeName(value) & \"|\" & CStr(VarType(value)) & \"|\" & CStr(value)\nEnd Function\nPublic Function NarrowArgument(ByVal value As Variant) As String\n On Error Resume Next\n Dim narrowed As LongLong\n narrowed = CLngLng(value)\n NarrowArgument = CStr(Err.Number) & \"|\" & CStr(narrowed)\nEnd Function\n";
const nativeRawInputCases = [{"name": "small", "number": 42.0, "native_vt": 3, "describe": "Long|3|42", "narrow": "0|42"}, {"name": "safe", "number": 9007199254740991.0, "native_vt": 5, "describe": "Double|5|9.00719925474099E+15", "narrow": "0|9007199254740991"}, {"name": "unsafe_exact", "number": 9007199254740992.0, "native_vt": 5, "describe": "Double|5|9.00719925474099E+15", "narrow": "0|9007199254740992"}, {"name": "below_positive_limit", "number": 9.223372036854775e+18, "native_vt": 5, "describe": "Double|5|9.22337203685477E+18", "narrow": "0|9223372036854774784"}, {"name": "positive_limit", "number": 9.223372036854776e+18, "native_vt": 5, "describe": "Double|5|9.22337203685478E+18", "narrow": "6|0"}, {"name": "above_positive_limit", "number": 9.223372036854778e+18, "native_vt": 5, "describe": "Double|5|9.22337203685478E+18", "narrow": "6|0"}, {"name": "negative_limit", "number": -9.223372036854776e+18, "native_vt": 5, "describe": "Double|5|-9.22337203685478E+18", "narrow": "0|-9223372036854775808"}, {"name": "below_negative_limit", "number": -9.223372036854778e+18, "native_vt": 5, "describe": "Double|5|-9.22337203685478E+18", "narrow": "6|0"}, {"name": "i32_minimum", "number": -2147483648.0, "native_vt": 3, "describe": "Long|3|-2147483648", "narrow": "0|-2147483648"}, {"name": "i32_maximum", "number": 2147483647.0, "native_vt": 3, "describe": "Long|3|2147483647", "narrow": "0|2147483647"}, {"name": "above_i32", "number": 2147483648.0, "native_vt": 5, "describe": "Double|5|2147483648", "narrow": "0|2147483648"}, {"name": "below_i32", "number": -2147483649.0, "native_vt": 5, "describe": "Double|5|-2147483649", "narrow": "0|-2147483649"}, {"name": "fraction", "number": 0.5, "native_vt": 5, "describe": "Double|5|0.5", "narrow": "0|0"}, {"name": "zero", "number": 0.0, "native_vt": 3, "describe": "Long|3|0", "narrow": "0|0"}];
const nativeTypedInputSource = "Public Function TypedScalarDriver(ByVal mode As Long) As String\n Dim value As Variant\n Select Case mode\n  Case 0\n   TypedScalarDriver = DescribeTypedArgument()\n   Exit Function\n  Case 1\n  Case 2: value = Null\n  Case 3: Set value = Nothing\n  Case 4: value = True\n  Case 5: value = False\n  Case 6: value = CByte(255)\n  Case 7: value = CInt(-32768)\n  Case 8: value = CInt(32767)\n  Case 9: value = CLng(\"-2147483648\")\n  Case 10: value = CLng(2147483647)\n  Case 11: value = -CLngLng(4611686018427387904#) - CLngLng(4611686018427387904#)\n  Case 12: value = CLngLng(4611686018427387904#) - CLngLng(1) + CLngLng(4611686018427387904#)\n  Case 13: value = CSng(0.1)\n  Case 14: value = 0.1#\n  Case 15: value = CCur(\"922337203685477.5807\")\n  Case 16: value = CCur(\"-922337203685477.5808\")\n  Case 17: value = CDec(\"79228162514264337593543950335\")\n  Case 18: value = CDec(\"7922816251426433759354395033.5\")\n  Case 19: value = CDate(45291.25)\n  Case 20: value = CVErr(2007)\n  Case 21: value = \"A\" & ChrW$(28450)\n  Case 22: value = vbNullString\n  Case 23: value = \"\"\n  Case 24: value = ChrW$(-10240) & \"A\" & ChrW$(-9216)\n  Case 25: value = \"A\" & Chr$(0) & \"B\"\n End Select\n DescribeResult:\n TypedScalarDriver = DescribeTypedArgument(value)\nEnd Function\n\nPublic Function DescribeTypedArgument(Optional ByVal value As Variant) As String\n On Error GoTo failed\n Dim text As String\n text = TypeName(value) & \"|\" & CStr(VarType(value))\n If IsMissing(value) Then\n  DescribeTypedArgument = text & \"|missing\"\n ElseIf IsEmpty(value) Then\n  DescribeTypedArgument = text & \"|empty\"\n ElseIf IsNull(value) Then\n  DescribeTypedArgument = text & \"|null\"\n ElseIf IsObject(value) Then\n  DescribeTypedArgument = text & \"|nothing=\" & CStr(value Is Nothing)\n ElseIf VarType(value) = vbString Then\n  Dim index As Long\n  text = text & \"|null_bstr=\" & CStr(StrPtr(value) = 0) & \"|units=\"\n  For index = 1 To Len(value)\n   If index > 1 Then text = text & \",\"\n   text = text & CStr(AscW(Mid$(value, index, 1)) And 65535)\n  Next index\n  DescribeTypedArgument = text\n ElseIf VarType(value) = vbDate Then\n  DescribeTypedArgument = text & \"|\" & CStr(CDbl(value))\n Else\n  DescribeTypedArgument = text & \"|\" & CStr(value)\n End If\n Exit Function\nfailed:\n DescribeTypedArgument = text & \"|failure=\" & CStr(Err.Number)\nEnd Function\n";
const nativeTypedInputCases = [{"mode": 0, "input": {"type": "missing"}, "expected": "Error|10|missing"}, {"mode": 1, "input": {"type": "empty"}, "expected": "Empty|0|empty"}, {"mode": 2, "input": {"type": "null"}, "expected": "Null|1|null"}, {"mode": 3, "input": {"type": "nothing"}, "expected": "Nothing|9|nothing=True"}, {"mode": 4, "input": {"type": "boolean", "value": true}, "expected": "Boolean|11|True"}, {"mode": 5, "input": {"type": "boolean", "value": false}, "expected": "Boolean|11|False"}, {"mode": 6, "input": {"type": "byte", "value": 255}, "expected": "Byte|17|255"}, {"mode": 7, "input": {"type": "integer", "value": -32768}, "expected": "Integer|2|-32768"}, {"mode": 8, "input": {"type": "integer", "value": 32767}, "expected": "Integer|2|32767"}, {"mode": 9, "input": {"type": "long", "value": -2147483648}, "expected": "Long|3|-2147483648"}, {"mode": 10, "input": {"type": "long", "value": 2147483647}, "expected": "Long|3|2147483647"}, {"mode": 11, "input": {"type": "long_long", "value": "-9223372036854775808"}, "expected": "LongLong|20|-9223372036854775808"}, {"mode": 12, "input": {"type": "long_long", "value": "9223372036854775807"}, "expected": "LongLong|20|9223372036854775807"}, {"mode": 13, "input": {"type": "single", "value": 0.1}, "expected": "Single|4|0.1"}, {"mode": 14, "input": {"type": "double", "value": 0.1}, "expected": "Double|5|0.1"}, {"mode": 15, "input": {"type": "currency", "value": "922337203685477.5807"}, "expected": "Currency|6|922337203685477.5807"}, {"mode": 16, "input": {"type": "currency", "value": "-922337203685477.5808"}, "expected": "Currency|6|-922337203685477.5808"}, {"mode": 17, "input": {"type": "decimal", "value": "79228162514264337593543950335"}, "expected": "Decimal|14|79228162514264337593543950335"}, {"mode": 18, "input": {"type": "decimal", "value": "7922816251426433759354395033.5"}, "expected": "Decimal|14|7922816251426433759354395033.5"}, {"mode": 19, "input": {"type": "date", "value": 45291.25}, "expected": "Date|7|45291.25"}, {"mode": 20, "input": {"type": "error", "value": 2007}, "expected": "Error|10|\u30a8\u30e9\u30fc 2007"}, {"mode": 21, "input": {"type": "string", "value": "A\u6f22"}, "expected": "String|8|null_bstr=False|units=65,28450"}, {"mode": 22, "input": {"type": "null_string"}, "expected": "String|8|null_bstr=True|units="}, {"mode": 23, "input": {"type": "string", "value": ""}, "expected": "String|8|null_bstr=False|units="}, {"mode": 24, "input": {"type": "utf16", "value": [55296, 65, 56320]}, "expected": "String|8|null_bstr=False|units=55296,65,56320"}, {"mode": 25, "input": {"type": "string", "value": "A\u0000B"}, "expected": "String|8|null_bstr=False|units=65,0,66"}];
const nativeRawDateTransfers = [{"mode": 0, "input": {"type": "date", "value": 2000036527.0}, "expected": "Date|7|2000036527"}, {"mode": 2, "input": {"type": "date", "value": 2958466.0}, "expected": "Date|7|2958466"}, {"mode": 4, "input": {"type": "date", "value": -1999963473.0}, "expected": "Date|7|-1999963473"}, {"mode": 6, "input": {"type": "date", "value": -657433.0}, "expected": "Date|7|-657433"}, {"mode": 7, "input": {"type": "date", "value": -657435.0}, "expected": "Date|7|-657435"}, {"mode": 9, "input": {"type": "date", "value": 45291.25}, "expected": "Date|7|45291.25"}, {"mode": 10, "input": {"type": "single", "value": -0.0}, "expected": "Single|4|-0"}, {"mode": 11, "input": {"type": "double", "value": -0.0}, "expected": "Double|5|-0"}, {"mode": 12, "input": {"type": "date", "value": -0.0}, "expected": "Date|7|-0"}];
const nativeDateInputGroups = [{"name": "typed_argument_date_boundary_native_controls", "source": "Public Function DateBoundaryProbe(ByVal serial As Double) As String\n On Error GoTo failed\n Dim stage As String\n Dim moment As Date\n Dim yearValue As Long\n stage = \"convert\"\n moment = CDate(serial)\n stage = \"year\"\n yearValue = Year(moment)\n DateBoundaryProbe = \"ok|\" & TypeName(moment) & \"|\" & CStr(VarType(moment)) & \"|\" & CStr(CDbl(moment)) & \"|year=\" & CStr(yearValue) & \"|same=\" & CStr(CDbl(moment) = serial)\n Exit Function\nfailed:\n DateBoundaryProbe = \"error|\" & stage & \"|\" & CStr(Err.Number)\nEnd Function\n", "cases": [{"name": "zero", "procedure": "DateBoundaryProbe", "args": [{"type": "double", "value": 0.0}], "expected": "ok|Date|7|0|year=1899|same=True"}, {"name": "negative_zero", "procedure": "DateBoundaryProbe", "args": [{"type": "double", "value": -0.0}], "expected": "ok|Date|7|-0|year=1899|same=True"}, {"name": "positive_fraction", "procedure": "DateBoundaryProbe", "args": [{"type": "double", "value": 0.75}], "expected": "ok|Date|7|0.75|year=1899|same=True"}, {"name": "negative_fraction", "procedure": "DateBoundaryProbe", "args": [{"type": "double", "value": -0.75}], "expected": "ok|Date|7|-0.75|year=1899|same=True"}, {"name": "lower_day", "procedure": "DateBoundaryProbe", "args": [{"type": "double", "value": -657434.0}], "expected": "ok|Date|7|-657434|year=100|same=True"}, {"name": "lower_fraction", "procedure": "DateBoundaryProbe", "args": [{"type": "double", "value": -657434.75}], "expected": "ok|Date|7|-657434.75|year=100|same=True"}, {"name": "just_above_lower", "procedure": "DateBoundaryProbe", "args": [{"type": "double", "value": -657434.9999999999}], "expected": "ok|Date|7|-657435|year=100|same=True"}, {"name": "lower_boundary", "procedure": "DateBoundaryProbe", "args": [{"type": "double", "value": -657435.0}], "expected": "error|convert|6"}, {"name": "just_below_lower", "procedure": "DateBoundaryProbe", "args": [{"type": "double", "value": -657435.0000000001}], "expected": "error|convert|6"}, {"name": "upper_day", "procedure": "DateBoundaryProbe", "args": [{"type": "double", "value": 2958465.0}], "expected": "ok|Date|7|2958465|year=9999|same=True"}, {"name": "upper_fraction", "procedure": "DateBoundaryProbe", "args": [{"type": "double", "value": 2958465.75}], "expected": "ok|Date|7|2958465.75|year=9999|same=True"}, {"name": "just_below_upper", "procedure": "DateBoundaryProbe", "args": [{"type": "double", "value": 2958465.9999999995}], "expected": "ok|Date|7|2958466|year=9999|same=True"}, {"name": "upper_boundary", "procedure": "DateBoundaryProbe", "args": [{"type": "double", "value": 2958466.0}], "expected": "error|convert|6"}, {"name": "just_above_upper", "procedure": "DateBoundaryProbe", "args": [{"type": "double", "value": 2958466.0000000005}], "expected": "error|convert|6"}, {"name": "typical", "procedure": "DateBoundaryProbe", "args": [{"type": "double", "value": 45291.25}], "expected": "ok|Date|7|45291.25|year=2023|same=True"}, {"name": "far_finite", "procedure": "DateBoundaryProbe", "args": [{"type": "double", "value": 1e+308}], "expected": "error|convert|6"}]}, {"name": "date_components_rounding_native_controls", "source": "Public Function CalendarParts(ByVal serial As Double) As String\n On Error GoTo failed\n Dim moment As Date\n moment = CDate(serial)\n CalendarParts = CStr(Year(moment)) & \"/\" & CStr(Month(moment)) & \"/\" & CStr(Day(moment)) & \" \" & CStr(Hour(moment)) & \":\" & CStr(Minute(moment)) & \":\" & CStr(Second(moment))\n CalendarParts = CalendarParts & \"|part=\" & CStr(DatePart(\"yyyy\", moment)) & \"/\" & CStr(DatePart(\"m\", moment)) & \"/\" & CStr(DatePart(\"d\", moment)) & \" \" & CStr(DatePart(\"h\", moment)) & \":\" & CStr(DatePart(\"n\", moment)) & \":\" & CStr(DatePart(\"s\", moment))\n Exit Function\nfailed:\n CalendarParts = \"error|\" & CStr(Err.Number)\nEnd Function\n\nPublic Function WrittenParts(ByVal serial As Double) As String\n On Error GoTo failed\n Dim moment As Date\n moment = CDate(serial)\n WrittenParts = \"date=\" & CStr(CDbl(DateValue(moment))) & \"|time=\" & CStr(CDbl(TimeValue(moment))) & \"|str=\" & CStr(moment) & \"|fmt=\" & Format$(moment, \"yyyy-mm-dd hh:nn:ss\")\n Exit Function\nfailed:\n WrittenParts = \"error|\" & CStr(Err.Number)\nEnd Function\n", "cases": [{"name": "base_day", "procedure": "CalendarParts", "args": [{"type": "double", "value": 45291.0}], "expected": "2023/12/31 0:0:0|part=2023/12/31 0:0:0"}, {"name": "base_day", "procedure": "WrittenParts", "args": [{"type": "double", "value": 45291.0}], "expected": "date=45291|time=0|str=12/31/2023|fmt=2023-12-31 00:00:00"}, {"name": "after_0499", "procedure": "CalendarParts", "args": [{"type": "double", "value": 45291.000005775466}], "expected": "2023/12/31 0:0:0|part=2023/12/31 0:0:0"}, {"name": "after_0499", "procedure": "WrittenParts", "args": [{"type": "double", "value": 45291.000005775466}], "expected": "date=45291|time=0|str=12/31/2023|fmt=2023-12-31 00:00:00"}, {"name": "after_0500", "procedure": "CalendarParts", "args": [{"type": "double", "value": 45291.000005787035}], "expected": "2023/12/31 0:0:0|part=2023/12/31 0:0:0"}, {"name": "after_0500", "procedure": "WrittenParts", "args": [{"type": "double", "value": 45291.000005787035}], "expected": "date=45291|time=0|str=12/31/2023|fmt=2023-12-31 00:00:00"}, {"name": "after_0501", "procedure": "CalendarParts", "args": [{"type": "double", "value": 45291.00000579861}], "expected": "2023/12/31 0:0:1|part=2023/12/31 0:0:1"}, {"name": "after_0501", "procedure": "WrittenParts", "args": [{"type": "double", "value": 45291.00000579861}], "expected": "date=45291|time=1.15740740740741E-05|str=12/31/2023 12:00:01 AM|fmt=2023-12-31 00:00:01"}, {"name": "before_next_0501", "procedure": "CalendarParts", "args": [{"type": "double", "value": 45291.99999420139}], "expected": "2023/12/31 23:59:59|part=2023/12/31 23:59:59"}, {"name": "before_next_0501", "procedure": "WrittenParts", "args": [{"type": "double", "value": 45291.99999420139}], "expected": "date=45291|time=0.999988425925926|str=12/31/2023 11:59:59 PM|fmt=2023-12-31 23:59:59"}, {"name": "before_next_0500", "procedure": "CalendarParts", "args": [{"type": "double", "value": 45291.999994212965}], "expected": "2024/1/1 0:0:0|part=2024/1/1 0:0:0"}, {"name": "before_next_0500", "procedure": "WrittenParts", "args": [{"type": "double", "value": 45291.999994212965}], "expected": "date=45292|time=0|str=1/1/2024|fmt=2024-01-01 00:00:00"}, {"name": "before_next_0499", "procedure": "CalendarParts", "args": [{"type": "double", "value": 45291.999994224534}], "expected": "2024/1/1 0:0:0|part=2024/1/1 0:0:0"}, {"name": "before_next_0499", "procedure": "WrittenParts", "args": [{"type": "double", "value": 45291.999994224534}], "expected": "date=45292|time=0|str=1/1/2024|fmt=2024-01-01 00:00:00"}, {"name": "max_near_upper", "procedure": "CalendarParts", "args": [{"type": "double", "value": 2958465.9999999995}], "expected": "9999/12/31 23:59:59|part=9999/12/31 23:59:59"}, {"name": "max_near_upper", "procedure": "WrittenParts", "args": [{"type": "double", "value": 2958465.9999999995}], "expected": "date=2958465|time=0.999988425925926|str=12/31/9999 11:59:59 PM|fmt=9999-12-31 23:59:59"}, {"name": "min_near_lower", "procedure": "CalendarParts", "args": [{"type": "double", "value": -657434.9999999999}], "expected": "100/1/2 0:0:0|part=100/1/2 0:0:0"}, {"name": "min_near_lower", "procedure": "WrittenParts", "args": [{"type": "double", "value": -657434.9999999999}], "expected": "date=-657433|time=0|str=1/2/100|fmt=0100-01-02 00:00:00"}, {"name": "negative_fraction", "procedure": "CalendarParts", "args": [{"type": "double", "value": -0.75}], "expected": "1899/12/30 18:0:0|part=1899/12/30 18:0:0"}, {"name": "negative_fraction", "procedure": "WrittenParts", "args": [{"type": "double", "value": -0.75}], "expected": "date=0|time=0.75|str=6:00:00 PM|fmt=1899-12-30 18:00:00"}]}, {"name": "date_components_upper_rounding_native_controls", "source": "Public Function CalendarParts(ByVal serial As Double) As String\n On Error GoTo failed\n Dim moment As Date\n moment = CDate(serial)\n CalendarParts = CStr(Year(moment)) & \"/\" & CStr(Month(moment)) & \"/\" & CStr(Day(moment)) & \" \" & CStr(Hour(moment)) & \":\" & CStr(Minute(moment)) & \":\" & CStr(Second(moment))\n CalendarParts = CalendarParts & \"|part=\" & CStr(DatePart(\"yyyy\", moment)) & \"/\" & CStr(DatePart(\"m\", moment)) & \"/\" & CStr(DatePart(\"d\", moment)) & \" \" & CStr(DatePart(\"h\", moment)) & \":\" & CStr(DatePart(\"n\", moment)) & \":\" & CStr(DatePart(\"s\", moment))\n Exit Function\nfailed:\n CalendarParts = \"error|\" & CStr(Err.Number)\nEnd Function\n\nPublic Function WrittenParts(ByVal serial As Double) As String\n On Error GoTo failed\n Dim moment As Date\n moment = CDate(serial)\n WrittenParts = \"date=\" & CStr(CDbl(DateValue(moment))) & \"|time=\" & CStr(CDbl(TimeValue(moment))) & \"|str=\" & CStr(moment) & \"|fmt=\" & Format$(moment, \"yyyy-mm-dd hh:nn:ss\")\n Exit Function\nfailed:\n WrittenParts = \"error|\" & CStr(Err.Number)\nEnd Function\n", "cases": [{"name": "upper_before_0501", "procedure": "CalendarParts", "args": [{"type": "double", "value": 2958465.9999942016}], "expected": "9999/12/31 23:59:59|part=9999/12/31 23:59:59"}, {"name": "upper_before_0501", "procedure": "WrittenParts", "args": [{"type": "double", "value": 2958465.9999942016}], "expected": "date=2958465|time=0.999988425925926|str=12/31/9999 11:59:59 PM|fmt=9999-12-31 23:59:59"}, {"name": "upper_before_0500", "procedure": "CalendarParts", "args": [{"type": "double", "value": 2958465.9999942128}], "expected": "9999/12/31 23:59:59|part=9999/12/31 23:59:59"}, {"name": "upper_before_0500", "procedure": "WrittenParts", "args": [{"type": "double", "value": 2958465.9999942128}], "expected": "date=2958465|time=0.999988425925926|str=12/31/9999 11:59:59 PM|fmt=9999-12-31 23:59:59"}, {"name": "upper_before_0499", "procedure": "CalendarParts", "args": [{"type": "double", "value": 2958465.9999942244}], "expected": "9999/12/31 23:59:59|part=9999/12/31 23:59:59"}, {"name": "upper_before_0499", "procedure": "WrittenParts", "args": [{"type": "double", "value": 2958465.9999942244}], "expected": "date=2958465|time=0.999988425925926|str=12/31/9999 11:59:59 PM|fmt=9999-12-31 23:59:59"}, {"name": "ordinary_before_0499", "procedure": "CalendarParts", "args": [{"type": "double", "value": 45291.999994224534}], "expected": "2024/1/1 0:0:0|part=2024/1/1 0:0:0"}, {"name": "ordinary_before_0499", "procedure": "WrittenParts", "args": [{"type": "double", "value": 45291.999994224534}], "expected": "date=45292|time=0|str=1/1/2024|fmt=2024-01-01 00:00:00"}]}, {"name": "date_general_epoch_rounding_native_controls", "source": "Public Function CalendarParts(ByVal serial As Double) As String\n On Error GoTo failed\n Dim moment As Date\n moment = CDate(serial)\n CalendarParts = CStr(Year(moment)) & \"/\" & CStr(Month(moment)) & \"/\" & CStr(Day(moment)) & \" \" & CStr(Hour(moment)) & \":\" & CStr(Minute(moment)) & \":\" & CStr(Second(moment))\n CalendarParts = CalendarParts & \"|part=\" & CStr(DatePart(\"yyyy\", moment)) & \"/\" & CStr(DatePart(\"m\", moment)) & \"/\" & CStr(DatePart(\"d\", moment)) & \" \" & CStr(DatePart(\"h\", moment)) & \":\" & CStr(DatePart(\"n\", moment)) & \":\" & CStr(DatePart(\"s\", moment))\n Exit Function\nfailed:\n CalendarParts = \"error|\" & CStr(Err.Number)\nEnd Function\n\nPublic Function WrittenParts(ByVal serial As Double) As String\n On Error GoTo failed\n Dim moment As Date\n moment = CDate(serial)\n WrittenParts = \"date=\" & CStr(CDbl(DateValue(moment))) & \"|time=\" & CStr(CDbl(TimeValue(moment))) & \"|str=\" & CStr(moment) & \"|fmt=\" & Format$(moment, \"yyyy-mm-dd hh:nn:ss\")\n Exit Function\nfailed:\n WrittenParts = \"error|\" & CStr(Err.Number)\nEnd Function\n", "cases": [{"name": "epoch_before_midnight", "procedure": "CalendarParts", "args": [{"type": "double", "value": 0.9999999999999999}], "expected": "1899/12/31 0:0:0|part=1899/12/31 0:0:0"}, {"name": "epoch_before_midnight", "procedure": "WrittenParts", "args": [{"type": "double", "value": 0.9999999999999999}], "expected": "date=1|time=0|str=12/31/1899|fmt=1899-12-31 00:00:00"}, {"name": "epoch_after_0501", "procedure": "CalendarParts", "args": [{"type": "double", "value": 5.798611111111111e-06}], "expected": "1899/12/30 0:0:1|part=1899/12/30 0:0:1"}, {"name": "epoch_after_0501", "procedure": "WrittenParts", "args": [{"type": "double", "value": 5.798611111111111e-06}], "expected": "date=0|time=1.15740740740741E-05|str=12:00:01 AM|fmt=1899-12-30 00:00:01"}, {"name": "negative_into_epoch", "procedure": "CalendarParts", "args": [{"type": "double", "value": -1.99999999}], "expected": "1899/12/30 0:0:0|part=1899/12/30 0:0:0"}, {"name": "negative_into_epoch", "procedure": "WrittenParts", "args": [{"type": "double", "value": -1.99999999}], "expected": "date=0|time=0|str=12:00:00 AM|fmt=1899-12-30 00:00:00"}, {"name": "negative_epoch_rollover", "procedure": "CalendarParts", "args": [{"type": "double", "value": -0.99999999}], "expected": "1899/12/31 0:0:0|part=1899/12/31 0:0:0"}, {"name": "negative_epoch_rollover", "procedure": "WrittenParts", "args": [{"type": "double", "value": -0.99999999}], "expected": "date=1|time=0|str=12/31/1899|fmt=1899-12-31 00:00:00"}, {"name": "negative_one", "procedure": "CalendarParts", "args": [{"type": "double", "value": -1.0}], "expected": "1899/12/29 0:0:0|part=1899/12/29 0:0:0"}, {"name": "negative_one", "procedure": "WrittenParts", "args": [{"type": "double", "value": -1.0}], "expected": "date=-1|time=0|str=12/29/1899|fmt=1899-12-29 00:00:00"}, {"name": "epoch_zero", "procedure": "CalendarParts", "args": [{"type": "double", "value": 0.0}], "expected": "1899/12/30 0:0:0|part=1899/12/30 0:0:0"}, {"name": "epoch_zero", "procedure": "WrittenParts", "args": [{"type": "double", "value": 0.0}], "expected": "date=0|time=0|str=12:00:00 AM|fmt=1899-12-30 00:00:00"}]}, {"name": "date_variant_state_native_controls", "source": "Public Function DateStateDriver(ByVal mode As Long) As String\n On Error GoTo failed\n Dim value As Variant\n Dim stage As String\n Dim text As String\n stage = \"construct\"\n Select Case mode\n  Case 0: value = CLng(2000000000) + #1/2/2000#\n  Case 1: value = #1/2/2000# + CLng(2000000000)\n  Case 2: value = 1# + CDate(2958465#)\n  Case 3: value = CDate(2958465#) + 1#\n  Case 4: value = CLng(-2000000000) + #1/2/2000#\n  Case 5: value = #1/2/2000# + CLng(-2000000000)\n  Case 6: value = 1# + CDate(-657434#)\n  Case 7: value = -1# + CDate(-657434#)\n  Case 8: value = CDate(-657434#) - 1#\n  Case 9: value = CDate(45291.25)\n End Select\n stage = \"type\"\n text = TypeName(value) & \"|\" & CStr(VarType(value))\n stage = \"number\"\n text = text & \"|number=\" & CStr(CDbl(value))\n stage = \"isdate\"\n text = text & \"|isdate=\" & CStr(IsDate(value))\n stage = \"text\"\n text = text & \"|text=\" & CStr(value)\n DateStateDriver = text\n Exit Function\nfailed:\n DateStateDriver = text & \"|failure=\" & stage & \":\" & CStr(Err.Number)\nEnd Function\n", "cases": [{"name": "0", "procedure": "DateStateDriver", "args": [0], "expected": "Date|7|number=2000036527|isdate=True|failure=text:5"}, {"name": "1", "procedure": "DateStateDriver", "args": [1], "expected": "|failure=construct:6"}, {"name": "2", "procedure": "DateStateDriver", "args": [2], "expected": "Date|7|number=2958466|isdate=True|failure=text:5"}, {"name": "3", "procedure": "DateStateDriver", "args": [3], "expected": "|failure=construct:6"}, {"name": "4", "procedure": "DateStateDriver", "args": [4], "expected": "Date|7|number=-1999963473|isdate=True|failure=text:5"}, {"name": "5", "procedure": "DateStateDriver", "args": [5], "expected": "|failure=construct:6"}, {"name": "6", "procedure": "DateStateDriver", "args": [6], "expected": "Date|7|number=-657433|isdate=True|text=1/2/100"}, {"name": "7", "procedure": "DateStateDriver", "args": [7], "expected": "Date|7|number=-657435|isdate=True|failure=text:5"}, {"name": "8", "procedure": "DateStateDriver", "args": [8], "expected": "|failure=construct:6"}, {"name": "9", "procedure": "DateStateDriver", "args": [9], "expected": "Date|7|number=45291.25|isdate=True|text=12/31/2023 6:00:00 AM"}]}];
const invalidNativeTypedInputs = [{"type": "byte", "value": 256}, {"type": "byte", "value": -1}, {"type": "integer", "value": 32768}, {"type": "long", "value": 2147483648}, {"type": "long_long", "value": "9223372036854775808"}, {"type": "long_long", "value": 42}, {"type": "currency", "value": "922337203685477.5808"}, {"type": "currency", "value": "-922337203685477.5809"}, {"type": "currency", "value": "0.00001"}, {"type": "decimal", "value": "79228162514264337593543950336"}, {"type": "decimal", "value": "0.00000000000000000000000000001"}, {"type": "decimal", "value": "1e100"}, {"type": "decimal", "value": "1."}, {"type": "single", "value": 1e+40}, {"type": "double", "value": "NaN"}, {"type": "utf16", "value": [65536]}, {"type": "utf16", "value": [1.5]}, {"type": "error", "value": 2147483648}, {"type": "empty", "value": 0}, {"type": "null", "extra": true}, {"type": "made_up", "value": 1}];

function parseNativeInputArgs(args) {
  return args.map(value => {
    if(value && typeof value==='object' && ['single','double','date'].includes(value.type)
        && Object.is(value.value,-0)) {
      const parsed=runner.parseVbaArguments('[{"type":'+JSON.stringify(value.type)+',"value":-0}]')[0];
      assert.ok(Object.is(parsed.value,-0),'JSON text parser preserves explicit negative zero');
      return parsed;
    }
    return runner.parseVbaArguments(JSON.stringify([value]))[0];
  });
}
async function verifyNativeInputAndDateCases(invoke) {
  let book=emptyBook();
  async function run(source,procedure,args) {
    const execution=await invoke(book,source,procedure,parseNativeInputArgs(args));
    assert.equal(execution.error,undefined,procedure);
    book=execution.workbook;
    return execution.result;
  }
  for(const entry of nativeRawInputCases) {
    const value=await run(nativeRawInputSource,'EchoArgument',[entry.number]);
    assert.equal(value.type,entry.native_vt===3?'integer':'double',entry.name+': Native argument subtype');
    assert.ok(Object.is(value.value,entry.number),entry.name+': exact Native numeric payload');
    assert.deepEqual(await run(nativeRawInputSource,'DescribeArgument',[entry.number]),
      {type:'string',value:entry.describe},entry.name+': Native TypeName/VarType/CStr');
    assert.deepEqual(await run(nativeRawInputSource,'NarrowArgument',[entry.number]),
      {type:'string',value:entry.narrow},entry.name+': Native CLngLng result');
  }
  for(const entry of nativeTypedInputCases) {
    assert.deepEqual(await run(nativeTypedInputSource,'DescribeTypedArgument',[entry.input]),
      {type:'string',value:entry.expected},'typed scalar '+entry.mode);
  }
  const unpaired={type:'string',value:String.fromCharCode(55296,65,56320)};
  assert.deepEqual(await run(nativeTypedInputSource,'DescribeTypedArgument',[unpaired]),
    {type:'string',value:nativeTypedInputCases.find(row=>row.mode===24).expected},
    'tagged JS String preserves Native UTF16 units');
  for(const entry of nativeRawDateTransfers) {
    assert.deepEqual(await run(nativeTypedInputSource,'DescribeTypedArgument',[entry.input]),
      {type:'string',value:entry.expected},'raw Date/negative-zero transfer '+entry.mode);
  }
  for(const input of invalidNativeTypedInputs) {
    await assert.rejects(async()=>invoke(book,nativeTypedInputSource,'DescribeTypedArgument',[input]),
      /invalid VBA argument/,'invalid typed transport is rejected');
  }
  assert.deepEqual(await run(nativeTypedInputSource,'DescribeTypedArgument',[{type:'boolean',value:true}]),
    {type:'string',value:nativeTypedInputCases.find(row=>row.mode===4).expected},
    'valid project invocation still works after rejected inputs');
  for(const group of nativeDateInputGroups) {
    for(const entry of group.cases) {
      assert.deepEqual(await run(group.source,entry.procedure,entry.args),
        {type:'string',value:entry.expected},group.name+' '+entry.name+' '+entry.procedure);
    }
  }
}
{
  const session=new wasm.SpreadsheetVbaSession();
  try {
    await verifyNativeInputAndDateCases((book,source,procedure,args)=>
      session.run(book,source,procedure,args,0,'NativeTypedInput.xlsm',{errorLanguage:'ja-JP'}));
  } finally {session.free();}
}
try {
  globalThis.Worker=BrowserWorker;
  runner.closeVbaProject();
  await verifyNativeInputAndDateCases((book,source,procedure,args)=>
    runner.runVbaInWorker(book,source,procedure,args,0,'NativeTypedInput.xlsm',
      {applicationState:{},errorLanguage:'ja-JP'}));
} finally {runner.closeVbaProject();globalThis.Worker=priorWorker;}
console.log(JSON.stringify({boundary_calls:342,typed_argument_rejections:42,excel_cases:cases.length,value_sequences:cases.filter(item=>item.outcome==='values').length,compile_cases:cases.filter(item=>item.outcome==='compile').length,value_steps:cases.filter(item=>item.outcome==='values').reduce((sum,item)=>sum+item.steps.length,0),direct_and_worker:true,passed:true}));
