// SPDX-License-Identifier: MIT OR Apache-2.0
import assert from 'node:assert/strict';
import {readFile} from 'node:fs/promises';
import path from 'node:path';
import {fileURLToPath,pathToFileURL} from 'node:url';
import {Worker as ThreadWorker} from 'node:worker_threads';
const root=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
const pkg=path.resolve(process.argv[2]||path.join(root,'crates/oxidocs-wasm/pkg'));
const bindingUrl='data:text/javascript;base64,'+Buffer.from(await readFile(path.join(pkg,'oxidocs_wasm.js'))).toString('base64');
const wasm=await import(bindingUrl);await wasm.default({module_or_path:await readFile(path.join(pkg,'oxidocs_wasm_bg.wasm'))});
const groups=[];
for(const name of ['open_file_operations','stream_sharing','copy_to_open_destination','fso_copy_move','fso_copy_move_edges','fso_arguments','file_argument_order']){
    const fixture=JSON.parse(await readFile(path.join(root,'tests/fixtures/vba',name+'.json'),'utf8'));
    groups.push({fixture,source:await readFile(path.join(root,'tests/fixtures/vba',fixture.source),'utf8'),
        bytes:new Uint8Array(await readFile(path.join(root,'tests/fixtures/vba',fixture.package)))});
}
const options=(group,item,state)=>({errorLanguage:'ja-JP',hostAnsiCodePage:932,applicationState:state,
    documentModules:wasm.spreadsheet_vba_document_modules(group.bytes),fileContext:{packageBytes:group.bytes,fullName:item.full_name,readOnly:true,folders:item.folders,files:item.files}});
const sorted=files=>files.map(file=>({path:file.path,bytes:Array.from(file.bytes)})).sort((a,b)=>a.path<b.path?-1:a.path>b.path?1:0);
function check(execution,item){
    const label='Native case '+item.kind;
    assert.equal(execution.error,undefined,label);
    assert.deepEqual(execution.result,{type:'string',value:item.expected},label);
    const base=item.arguments[1].toLowerCase(),prefix=base+'\\',state=execution.file_system_state;
    assert.deepEqual(sorted(state.files.filter(file=>file.path.toLowerCase().startsWith(prefix))),sorted(item.expected_files),label+' file bytes/paths');
    assert.deepEqual(state.folders.map(path=>path.toLowerCase()).filter(path=>path===base||path.startsWith(prefix)).sort(),item.expected_folder_keys,label+' folders');
    const kept=new Set(item.expected_files.map(file=>file.path.toLowerCase()));
    assert.deepEqual([...(execution.deleted_files||[])].sort(),item.files.filter(file=>!kept.has(file.path.toLowerCase())).map(file=>file.path).sort(),label+' deletion outputs');
    const overrides=new Map((state.attributes||[]).map(entry=>[entry.path.toLowerCase(),entry.value]));
    for(const entry of item.expected_attributes){
        const folder=state.folders.some(path=>path.toLowerCase()===entry.path);
        assert.equal(overrides.get(entry.path) ?? (folder?16:32),entry.value,label+' attributes '+entry.path);
    }
    for(const entry of state.attributes||[]){
        const key=entry.path.toLowerCase();
        if(key===base||key.startsWith(prefix))assert.ok(item.expected_attributes.some(expected=>expected.path===key),label+' stale attributes');
    }
}
let direct=0;
for(const group of groups){
    const state={displayAlerts:false,directoryState:group.fixture.initial_directory_state};
    for(const item of group.fixture.cases){
        const execution=wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(group.bytes),group.source,item.procedure,item.arguments,0,'input.xlsm',options(group,item,state));
        check(execution,item);direct++;
    }
}
const shim=`import init, {SpreadsheetVbaSession,spreadsheet_vba_source,spreadsheet_vba_project_name} from ${JSON.stringify(bindingUrl)};
import {readFile} from 'node:fs/promises';await init({module_or_path:await readFile(${JSON.stringify(path.join(pkg,'oxidocs_wasm_bg.wasm'))})});
export {SpreadsheetVbaSession,spreadsheet_vba_source,spreadsheet_vba_project_name};export default function ready(){return Promise.resolve();}`;
const workerSource=(await readFile(path.join(root,'web/vba-worker.js'),'utf8')).replace("'./oxidocs_wasm.js'",JSON.stringify('data:text/javascript;base64,'+Buffer.from(shim).toString('base64')));
const bridge=`import {parentPort,workerData} from 'node:worker_threads';const queued=[];let listener;
globalThis.self={postMessage:value=>parentPort.postMessage(value),addEventListener:(type,handler)=>{listener=handler;for(const data of queued)listener({data});}};
parentPort.on('message',data=>listener?listener({data}):queued.push(data));await import('data:text/javascript;base64,'+Buffer.from(workerData).toString('base64'));`;
class BrowserWorker {
    constructor(){this.thread=new ThreadWorker(new URL('data:text/javascript;base64,'+Buffer.from(bridge).toString('base64')),{workerData:workerSource});}
    addEventListener(type,handler){this.thread.on(type,value=>handler(type==='message'?{data:value}:value));}
    postMessage(value){this.thread.postMessage(value);}
    terminate(){void this.thread.terminate();}
}
const originalWorker=globalThis.Worker;globalThis.Worker=BrowserWorker;
const runner=await import(pathToFileURL(path.join(root,'web/vba-runner.js')));let worker=0;
try {
    for(const group of groups){
        for(const item of group.fixture.cases){
            runner.closeVbaProject();
            const state=worker===0?{displayAlerts:false,directoryState:group.fixture.initial_directory_state}:undefined;
            const execution=await runner.runVbaInWorker(wasm.parse_spreadsheet(group.bytes),group.source,item.procedure,item.arguments,0,'input.xlsm',options(group,item,state),group.bytes);
            check(execution,item);worker++;
        }
    }
} finally {runner.closeVbaProject();globalThis.Worker=originalWorker;}
assert.equal(direct,163);assert.equal(worker,163);

// Native TextStream controls: numbered Close/Reset and intrinsic lifetimes.
const textStreamSource="Attribute VB_Name = \"IOOwner\"\nOption Explicit\nPrivate HeldFS As Object\nPrivate HeldStream As Object\nPublic Function FileState(ByVal mode As Long, ByVal path As String, ByVal channel As Long, ByVal payload As String) As String\n Dim line As String\n On Error GoTo failed\n Select Case mode\n Case 0: FileState = \"free|\" & CStr(FreeFile(0)) & \"|\" & CStr(FreeFile(1))\n Case 1\n  Open path For Input As #channel\n  FileState = \"opened\"\n Case 2\n  Line Input #channel, line\n  FileState = \"line|\" & line & \"|loc|\" & CStr(Loc(channel)) & \"|eof|\" & CStr(EOF(channel))\n Case 3: FileState = \"state|\" & CStr(LOF(channel)) & \"|\" & CStr(Seek(channel)) & \"|\" & CStr(EOF(channel))\n Case 4\n  Close #channel\n  FileState = \"closed\"\n Case 5\n  Open path For Append As #channel\n  Print #channel, payload\n  FileState = \"append\"\n Case 6\n  Print #channel, payload\n  FileState = \"wrote\"\n Case 7\n  Close\n  FileState = \"closed-all\"\n End Select\n Exit Function\nfailed:\n FileState = \"error|\" & CStr(Err.Number) & \"|\" & Err.Description & \"|\" & Err.Source & \"|\" & Err.HelpFile & \"|\" & CStr(Err.HelpContext)\nEnd Function\nPublic Function DirectoryProbe(ByVal pattern As String) As String\n On Error GoTo failed\n If Len(pattern) = 0 Then\n  DirectoryProbe = \"dir|\" & Dir()\n Else\n  DirectoryProbe = \"dir|\" & Dir(pattern)\n End If\n Exit Function\nfailed:\n DirectoryProbe = \"error|\" & CStr(Err.Number) & \"|\" & Err.Description & \"|\" & Err.Source & \"|\" & Err.HelpFile & \"|\" & CStr(Err.HelpContext)\nEnd Function\nPublic Function CreateFileSystem() As Object\n Set CreateFileSystem = CreateObject(\"Scripting.FileSystemObject\")\nEnd Function\nPublic Function CreateTextStream(ByVal path As String) As Object\n Dim fs As Object\n Set fs = CreateObject(\"Scripting.FileSystemObject\")\n Set CreateTextStream = fs.OpenTextFile(path, 1)\nEnd Function\nPublic Function HoldFS(ByVal candidate As Object) As String\n Set HeldFS = candidate\n HoldFS = TypeName(HeldFS)\nEnd Function\nPublic Function HoldStream(ByVal candidate As Object) As String\n Set HeldStream = candidate\n HoldStream = TypeName(HeldStream)\nEnd Function\nPublic Function InspectFS(ByVal path As String) As String\n On Error GoTo failed\n InspectFS = \"fs|\" & CStr(HeldFS.FileExists(path)) & \"|\" & CStr(HeldFS.GetFile(path).Size)\n Exit Function\nfailed:\n InspectFS = \"error|\" & CStr(Err.Number) & \"|\" & Err.Description\nEnd Function\nPublic Function ReadStream() As String\n On Error GoTo failed\n ReadStream = \"stream|\" & HeldStream.ReadLine\n Exit Function\nfailed:\n ReadStream = \"error|\" & CStr(Err.Number) & \"|\" & Err.Description\nEnd Function\nPublic Function ReleaseIO() As String\n On Error Resume Next\n HeldStream.Close\n Set HeldStream = Nothing\n Set HeldFS = Nothing\n Close\n ReleaseIO = \"released\"\nEnd Function\nPublic Function StreamCloseProbe(ByVal path As String, ByVal mode As Long) As String\n Dim fs As Object, stream As Object\n On Error GoTo failed\n Set fs = CreateObject(\"Scripting.FileSystemObject\")\n Set stream = fs.OpenTextFile(path, 1)\n If mode = 0 Then\n  Open path For Input As #1\n  Close\n ElseIf mode = 1 Then\n  Close\n ElseIf mode = 3 Then\n  Open path For Input As #1\n  Reset\n Else\n  stream.Close\n End If\n StreamCloseProbe = CStr(FreeFile(0)) & \"|\" & stream.ReadLine & \"|\" & CStr(stream.AtEndOfStream)\n stream.Close\n Exit Function\nfailed:\n StreamCloseProbe = \"error|\" & CStr(Err.Number) & \"|\" & Err.Description & \"|\" & Err.Source & \"|\" & CStr(Err.HelpContext)\n On Error Resume Next\n stream.Close\n Close #1\nEnd Function\nPublic Function ClosedStreamProbe(ByVal path As String, ByVal mode As Long, ByVal ioMode As Long, ByVal closed As Boolean) As String\n Dim fs As Object, stream As Object, answer As Variant\n On Error GoTo failed\n Set fs = CreateObject(\"Scripting.FileSystemObject\")\n Set stream = fs.OpenTextFile(path, ioMode, True)\n If closed Then stream.Close\n Select Case mode\n Case 0: answer = stream.ReadLine\n Case 1: answer = stream.Read(1)\n Case 2: answer = stream.ReadAll\n Case 3: stream.SkipLine\n Case 4: stream.Skip 1\n Case 5: answer = stream.AtEndOfStream\n Case 6: answer = stream.AtEndOfLine\n Case 7: answer = stream.Line\n Case 8: answer = stream.Column\n Case 9: stream.Close\n Case 10: stream.Write \"x\"\n Case 11: stream.WriteLine \"x\"\n Case 12: stream.WriteBlankLines 1\n Case 13: answer = stream.NoSuchMember()\n Case 14: answer = stream.Read()\n Case 15: answer = stream.Read(-1)\n Case 16: answer = stream.Read(\"bad\")\n Case 17: answer = stream.Read(Null)\n Case 18: answer = CallByName(stream, \"Read\", vbGet, 1)\n Case 19: answer = CallByName(stream, \"ReadLine\", vbMethod)\n Case 20: answer = CallByName(stream, \"ReadLine\", vbGet)\n Case 21: answer = TypeName(stream)\n Case 22: answer = (stream Is Nothing)\n Case 23: answer = stream.Read(0)\n Case 24: stream.Close 1\n Case 25: answer = stream.AtEndOfStream(1)\n Case 26: answer = CallByName(stream, \"AtEndOfStream\", vbMethod)\n Case 27: answer = CallByName(stream, \"Column\", vbGet)\n Case 28: answer = CallByName(stream, \"Line\", 3)\n Case 29: answer = CallByName(stream, \"ReadLine\", 3)\n Case 30: answer = CallByName(stream, \"AtEndOfStream\", vbGet, 1)\n Case 31: answer = CallByName(stream, \"Line\", vbGet, 1)\n Case 32: answer = CallByName(stream, \"Column\", vbGet, 1)\n Case 33: answer = CallByName(stream, \"AtEndOfLine\", vbGet, 1)\n Case 34: answer = stream.Column(1)\n Case 35: answer = stream.Read(100)\n Case 36\n  answer = stream.ReadAll\n  answer = stream.ReadAll\n Case 37: answer = stream.Read(0.5)\n Case 38: answer = stream.Read(2.5)\n Case 39: stream.Skip 100\n Case 40: stream.WriteBlankLines 0\n Case 41: stream.WriteBlankLines -1\n Case 42: stream.WriteLine\n Case 43: stream.Write\n Case 44: stream.Write Null\n Case 45: stream.WriteLine Null\n Case 46: stream.WriteBlankLines Null\n Case 47: answer = stream.Read(2147483648#)\n Case 48: stream.Skip\n Case 49: stream.Skip -1\n Case 50: answer = stream.Read(Empty)\n Case 51: answer = stream.Read(CVErr(7))\n Case 52: stream.Write \"x\", \"y\"\n Case 53: stream.SkipLine 1\n Case 54: answer = stream.ReadLine(1)\n Case 55: answer = stream.ReadAll(1)\n Case 56: answer = CallByName(stream, \"Close\", vbGet)\n Case 57: stream.Skip 0\n Case 58: stream.WriteBlankLines \"bad\"\n Case 59: answer = CallByName(stream, \"AtEndOfStream\", vbGet)\n End Select\n ClosedStreamProbe = \"value|\" & TypeName(answer) & \"|\" & CStr(answer)\n Exit Function\nfailed:\n ClosedStreamProbe = \"error|\" & CStr(Err.Number) & \"|\" & Err.Description & \"|\" & Err.Source & \"|\" & Err.HelpFile & \"|\" & CStr(Err.HelpContext)\nEnd Function\nPublic Function CreateWriter(ByVal path As String) As Object\n Dim fs As Object\n Set fs = CreateObject(\"Scripting.FileSystemObject\")\n Set CreateWriter = fs.CreateTextFile(path, True)\nEnd Function\nPublic Sub CloseNumbered()\n Close\nEnd Sub\n";
const textStreamNumbers=[[91, 91, 91, 91, 91, 91, 91, 91, 91, 0, 91, 54, 91, 438, 450, 5, 13, 13, 438, 91, 438, 0, 0, 0, 450, 91, 438, 91, 91, 91, 91, 91, 91, 91, 91, 91, 91, 0, 91, 91, 91, 91, 54, 450, 13, 13, 13, 6, 450, 5, 0, 13, 450, 450, 91, 91, 438, 0, 13, 91], [54, 54, 54, 54, 54, 91, 54, 91, 91, 0, 91, 91, 91, 438, 450, 54, 13, 13, 438, 54, 438, 0, 0, 54, 450, 91, 438, 91, 91, 54, 91, 91, 91, 54, 91, 54, 54, 54, 54, 54, 91, 91, 91, 450, 13, 13, 13, 6, 450, 54, 54, 13, 450, 450, 54, 54, 438, 54, 13, 91], [54, 54, 54, 54, 54, 91, 54, 91, 91, 0, 91, 91, 91, 438, 450, 54, 13, 13, 438, 54, 438, 0, 0, 54, 450, 91, 438, 91, 91, 54, 91, 91, 91, 54, 91, 54, 54, 54, 54, 54, 91, 91, 91, 450, 13, 13, 13, 6, 450, 54, 54, 13, 450, 450, 54, 54, 438, 54, 13, 91], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 54, 54, 54, 438, 450, 5, 13, 13, 438, 0, 438, 0, 0, 0, 450, 451, 438, 0, 0, 0, 451, 451, 451, 451, 451, 0, 62, 0, 0, 0, 54, 54, 54, 450, 13, 13, 13, 6, 450, 5, 0, 13, 450, 450, 451, 440, 438, 0, 13, 0], [54, 54, 54, 54, 54, 54, 54, 0, 0, 0, 0, 0, 0, 438, 450, 54, 13, 13, 438, 54, 438, 0, 0, 54, 450, 54, 438, 0, 0, 54, 54, 451, 451, 54, 451, 54, 54, 54, 54, 54, 0, 5, 0, 450, 13, 13, 13, 6, 450, 54, 54, 13, 450, 450, 54, 54, 438, 54, 13, 54], [54, 54, 54, 54, 54, 54, 54, 0, 0, 0, 0, 0, 0, 438, 450, 54, 13, 13, 438, 54, 438, 0, 0, 54, 450, 54, 438, 0, 0, 54, 54, 451, 451, 54, 451, 54, 54, 54, 54, 54, 0, 5, 0, 450, 13, 13, 13, 6, 450, 54, 54, 13, 450, 450, 54, 54, 438, 54, 13, 54]];
const textStreamDescriptions={"91": "\u30aa\u30d6\u30b8\u30a7\u30af\u30c8\u5909\u6570\u307e\u305f\u306f With \u30d6\u30ed\u30c3\u30af\u5909\u6570\u304c\u8a2d\u5b9a\u3055\u308c\u3066\u3044\u307e\u305b\u3093\u3002", "54": "\u30d5\u30a1\u30a4\u30eb \u30e2\u30fc\u30c9\u304c\u4e0d\u6b63\u3067\u3059\u3002", "438": "\u30aa\u30d6\u30b8\u30a7\u30af\u30c8\u306f\u3001\u3053\u306e\u30d7\u30ed\u30d1\u30c6\u30a3\u307e\u305f\u306f\u30e1\u30bd\u30c3\u30c9\u3092\u30b5\u30dd\u30fc\u30c8\u3057\u3066\u3044\u307e\u305b\u3093\u3002", "450": "\u5f15\u6570\u306e\u6570\u304c\u4e00\u81f4\u3057\u3066\u3044\u307e\u305b\u3093\u3002\u307e\u305f\u306f\u4e0d\u6b63\u306a\u30d7\u30ed\u30d1\u30c6\u30a3\u3092\u6307\u5b9a\u3057\u3066\u3044\u307e\u3059\u3002", "5": "\u30d7\u30ed\u30b7\u30fc\u30b8\u30e3\u306e\u547c\u3073\u51fa\u3057\u3001\u307e\u305f\u306f\u5f15\u6570\u304c\u4e0d\u6b63\u3067\u3059\u3002", "13": "\u578b\u304c\u4e00\u81f4\u3057\u307e\u305b\u3093\u3002", "6": "\u30aa\u30fc\u30d0\u30fc\u30d5\u30ed\u30fc\u3057\u307e\u3057\u305f\u3002", "451": "Property Let \u30d7\u30ed\u30b7\u30fc\u30b8\u30e3\u304c\u5b9a\u7fa9\u3055\u308c\u3066\u304a\u3089\u305a\u3001Property Get \u30d7\u30ed\u30b7\u30fc\u30b8\u30e3\u304b\u3089\u30aa\u30d6\u30b8\u30a7\u30af\u30c8\u304c\u8fd4\u3055\u308c\u307e\u305b\u3093\u3067\u3057\u305f\u3002", "62": "\u30d5\u30a1\u30a4\u30eb\u306b\u3053\u308c\u4ee5\u4e0a\u30c7\u30fc\u30bf\u304c\u3042\u308a\u307e\u305b\u3093\u3002", "440": "\u30aa\u30fc\u30c8\u30e1\u30fc\u30b7\u30e7\u30f3 \u30a8\u30e9\u30fc\u3067\u3059\u3002"};
const textStreamValues={"true:1:9": "value|Empty|", "true:1:21": "value|String|TextStream", "true:1:22": "value|Boolean|False", "true:1:23": "value|String|", "true:1:37": "value|String|", "true:1:50": "value|String|", "true:1:57": "value|Empty|", "true:2:9": "value|Empty|", "true:2:21": "value|String|TextStream", "true:2:22": "value|Boolean|False", "true:8:9": "value|Empty|", "true:8:21": "value|String|TextStream", "true:8:22": "value|Boolean|False", "false:1:0": "value|String|first", "false:1:1": "value|String|f", "false:1:2": "value|String|first\r\nsecond\r\nthird\r\n", "false:1:3": "value|Empty|", "false:1:4": "value|Empty|", "false:1:5": "value|Boolean|False", "false:1:6": "value|Boolean|False", "false:1:7": "value|Long|1", "false:1:8": "value|Long|1", "false:1:9": "value|Empty|", "false:1:19": "value|String|first", "false:1:21": "value|String|TextStream", "false:1:22": "value|Boolean|False", "false:1:23": "value|String|", "false:1:27": "value|Long|1", "false:1:28": "value|Long|1", "false:1:29": "value|String|first", "false:1:35": "value|String|first\r\nsecond\r\nthird\r\n", "false:1:37": "value|String|", "false:1:38": "value|String|fi", "false:1:39": "value|Empty|", "false:1:50": "value|String|", "false:1:57": "value|Empty|", "false:1:59": "value|Boolean|False", "false:2:7": "value|Long|1", "false:2:8": "value|Long|1", "false:2:9": "value|Empty|", "false:2:10": "value|Empty|", "false:2:11": "value|Empty|", "false:2:12": "value|Empty|", "false:2:21": "value|String|TextStream", "false:2:22": "value|Boolean|False", "false:2:27": "value|Long|1", "false:2:28": "value|Long|1", "false:2:40": "value|Empty|", "false:2:42": "value|Empty|", "false:8:7": "value|Long|1", "false:8:8": "value|Long|1", "false:8:9": "value|Empty|", "false:8:10": "value|Empty|", "false:8:11": "value|Empty|", "false:8:12": "value|Empty|", "false:8:21": "value|String|TextStream", "false:8:22": "value|Boolean|False", "false:8:27": "value|Long|1", "false:8:28": "value|Long|3", "false:8:40": "value|Empty|", "false:8:42": "value|Empty|"};
const textStreamPackage=wasm.save_spreadsheet_vba_source(groups[0].bytes,textStreamSource,undefined);
const textStreamRoot='C:\\stream-tests';
const textStreamInput=textStreamRoot+'\\input.txt';
const textStreamOptions={errorLanguage:'ja-JP',projectName:'VBAProject',hostAnsiCodePage:932,
 documentModules:wasm.spreadsheet_vba_document_modules(textStreamPackage),
 fileContext:{packageBytes:textStreamPackage,fullName:textStreamRoot+'\\stream.xlsm',readOnly:true,
 folders:[textStreamRoot],files:[{path:textStreamInput,bytes:Array.from(Buffer.from('first\r\nsecond\r\nthird\r\n'))}]}};
async function verifyTextStreams(api){
 let book=wasm.parse_spreadsheet(textStreamPackage),calls=0;
 async function run(procedure,args,options=textStreamOptions){
  const execution=await api.run(book,procedure,args,options);calls++;
  assert.equal(execution.error,undefined,procedure);book=execution.workbook;return execution.result;
 }
 for(let mode=0;mode<4;mode++){
  const expected=mode===2?'error|91|'+textStreamDescriptions[91]+'|VBAProject|1000091':'1|first|False';
  assert.deepEqual(await run('IOOwner.StreamCloseProbe',[textStreamInput,mode]),{type:'string',value:expected});
 }
 let group=0;
 for(const closed of [true,false])for(const io of [1,2,8]){
  for(let mode=0;mode<60;mode++){
   const file=textStreamRoot+(io===1?'\\input.txt':io===2?'\\output.txt':'\\append.txt');
   const number=textStreamNumbers[group][mode];
   const expected=number?'error|'+number+'|'+textStreamDescriptions[number]+'|VBAProject||'+(1000000+number):textStreamValues[String(closed)+':'+io+':'+mode];
   assert.deepEqual(await run('IOOwner.ClosedStreamProbe',[file,mode,io,closed]),{type:'string',value:expected},'closed='+closed+' io='+io+' mode='+mode);
  }
  group++;
 }
 // Caller-held streams survive VBA Close. Closing an old alias cannot close
 // a later stream, and releasing the final lease must release the file lock.
 const first=await run('IOOwner.CreateTextStream',[textStreamInput]);
 assert.equal(first.type,'object');assert.equal(first.value.kind,'TextStream');
 const firstReference=first.value.reference;
 await run('IOOwner.CloseNumbered',[],{errorLanguage:'ja-JP'});
 async function invoke(reference,member){
  let execution;
  try{execution=await api.invoke(book,reference,member);}
  catch(error){if(!error?.execution)throw error;execution=error.execution;}
  calls++;book=execution.workbook;return execution;
 }
 let execution=await invoke(firstReference,'ReadLine');
 assert.equal(execution.error,undefined);assert.deepEqual(execution.result,{type:'string',value:'first'});
 execution=await invoke(firstReference,'Close');assert.equal(execution.error,undefined);
 const second=await run('IOOwner.CreateTextStream',[textStreamInput],{errorLanguage:'ja-JP'});
 execution=await invoke(firstReference,'ReadLine');assert.ok(execution.error,'closed alias must stay closed');
 execution=await api.release(book,[firstReference]);calls++;assert.equal(execution.error,undefined);book=execution.workbook;
 execution=await invoke(second.value.reference,'ReadLine');
 assert.equal(execution.error,undefined);assert.deepEqual(execution.result,{type:'string',value:'first'});
 execution=await api.release(book,[second.value.reference]);calls++;assert.equal(execution.error,undefined);book=execution.workbook;
 assert.deepEqual(execution.reference_state.tokens,[]);
 const writer=await run('IOOwner.CreateWriter',[textStreamInput],{errorLanguage:'ja-JP'});
 assert.equal(writer.type,'object','last stream lease must release its read lock');
 execution=await api.release(book,[writer.value.reference]);calls++;assert.equal(execution.error,undefined);
 assert.deepEqual(execution.reference_state.tokens,[]);
 return calls;
}
const textStreamSession=new wasm.SpreadsheetVbaSession();
let textStreamDirectCalls,textStreamWorkerCalls;
try{
 textStreamDirectCalls=await verifyTextStreams({
  run:(book,procedure,args,options)=>textStreamSession.run(book,textStreamSource,procedure,args,0,'stream.xlsm',options),
  invoke:(book,reference,member)=>textStreamSession.invoke_member(book,textStreamSource,{reference,member,kind:1,args:[]},0,'stream.xlsm',{}),
  release:(book,references)=>textStreamSession.release_objects(book,references,0,'stream.xlsm',{})
 });
}finally{textStreamSession.free();}
try{
 globalThis.Worker=BrowserWorker;runner.closeVbaProject();
 textStreamWorkerCalls=await verifyTextStreams({
  run:(book,procedure,args,options)=>runner.runVbaInWorker(book,textStreamSource,procedure,args,0,'stream.xlsm',options,textStreamPackage),
  invoke:(book,reference,member)=>runner.invokeVbaMemberInWorker(book,textStreamSource,reference,member,1,[],0,'stream.xlsm',{},textStreamPackage),
  release:(book,references)=>runner.releaseVbaObjectsInWorker(book,textStreamSource,references,0,'stream.xlsm',{},textStreamPackage)
 });
}finally{runner.closeVbaProject();globalThis.Worker=originalWorker;}
assert.equal(textStreamDirectCalls,textStreamWorkerCalls);
assert.equal(textStreamDirectCalls,375);

console.log(JSON.stringify({native_cases:163,direct,worker,text_stream_cases:360,text_stream_close_reset_cases:4,text_stream_direct_calls:textStreamDirectCalls,text_stream_worker_calls:textStreamWorkerCalls,complete_results:true,post_file_bytes_and_folders:true,post_attributes:true,deletion_outputs:true,passed:true}));
