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


const endGraphSource="Attribute VB_Name = \"IOOwner\"\nOption Explicit\nPrivate OwnCounter As Long\nPrivate HeldClass As Object\nPrivate FreshClass As Object\nPrivate HeldCollection As Object\nPrivate HeldDictionary As Object\nPrivate HeldFS As Object\nPrivate HeldStream As Object\nPublic Function FileState(ByVal mode As Long, ByVal path As String, ByVal channel As Long, ByVal payload As String) As String\n Dim line As String\n On Error GoTo failed\n Select Case mode\n Case 0: FileState = \"free|\" & CStr(FreeFile(0)) & \"|\" & CStr(FreeFile(1))\n Case 1\n  Open path For Input As #channel\n  FileState = \"opened\"\n Case 2\n  Line Input #channel, line\n  FileState = \"line|\" & line & \"|loc|\" & CStr(Loc(channel)) & \"|eof|\" & CStr(EOF(channel))\n Case 3: FileState = \"state|\" & CStr(LOF(channel)) & \"|\" & CStr(Seek(channel)) & \"|\" & CStr(EOF(channel))\n Case 4\n  Close #channel\n  FileState = \"closed\"\n Case 5\n  Open path For Append As #channel\n  Print #channel, payload\n  FileState = \"append\"\n Case 6\n  Print #channel, payload\n  FileState = \"wrote\"\n Case 7\n  Close\n  FileState = \"closed-all\"\n End Select\n Exit Function\nfailed:\n FileState = \"error|\" & CStr(Err.Number) & \"|\" & Err.Description & \"|\" & Err.Source & \"|\" & Err.HelpFile & \"|\" & CStr(Err.HelpContext)\nEnd Function\nPublic Function DirectoryProbe(ByVal pattern As String) As String\n On Error GoTo failed\n If Len(pattern) = 0 Then\n  DirectoryProbe = \"dir|\" & Dir()\n Else\n  DirectoryProbe = \"dir|\" & Dir(pattern)\n End If\n Exit Function\nfailed:\n DirectoryProbe = \"error|\" & CStr(Err.Number) & \"|\" & Err.Description & \"|\" & Err.Source & \"|\" & Err.HelpFile & \"|\" & CStr(Err.HelpContext)\nEnd Function\nPublic Function CreateFileSystem() As Object\n Set CreateFileSystem = CreateObject(\"Scripting.FileSystemObject\")\nEnd Function\nPublic Function CreateTextStream(ByVal path As String) As Object\n Dim fs As Object\n Set fs = CreateObject(\"Scripting.FileSystemObject\")\n Set CreateTextStream = fs.OpenTextFile(path, 1)\nEnd Function\nPublic Function HoldFS(ByVal candidate As Object) As String\n Set HeldFS = candidate\n HoldFS = TypeName(HeldFS)\nEnd Function\nPublic Function HoldStream(ByVal candidate As Object) As String\n Set HeldStream = candidate\n HoldStream = TypeName(HeldStream)\nEnd Function\nPublic Function InspectFS(ByVal path As String) As String\n On Error GoTo failed\n InspectFS = \"fs|\" & CStr(HeldFS.FileExists(path)) & \"|\" & CStr(HeldFS.GetFile(path).Size)\n Exit Function\nfailed:\n InspectFS = \"error|\" & CStr(Err.Number) & \"|\" & Err.Description\nEnd Function\nPublic Function ReadStream() As String\n On Error GoTo failed\n ReadStream = \"stream|\" & HeldStream.ReadLine\n Exit Function\nfailed:\n ReadStream = \"error|\" & CStr(Err.Number) & \"|\" & Err.Description\nEnd Function\nPublic Function ReleaseIO() As String\n On Error Resume Next\n HeldStream.Close\n Set HeldStream = Nothing\n Set HeldFS = Nothing\n Set HeldClass = Nothing\n Set FreshClass = Nothing\n Set HeldCollection = Nothing\n Set HeldDictionary = Nothing\n Close\n ReleaseIO = \"released\"\nEnd Function\nPublic Function Counter(ByVal value As Long) As Long\n If value >= 0 Then OwnCounter = value\n Counter = OwnCounter\nEnd Function\nPublic Sub TriggerStop(ByVal mode As Long)\n Select Case mode\n Case 1: End\n Case 2: Err.Raise 5\n End Select\nEnd Sub\nPublic Function CreateClass() As Object\n Dim item As New RetainedClass\n item.Seed 77\n Set CreateClass = item\nEnd Function\nPublic Function CreateCollection(ByVal item As Object, ByVal stream As Object) As Object\n Dim result As New Collection\n result.Add \"kept\"\n result.Add item\n result.Add stream\n Set CreateCollection = result\nEnd Function\nPublic Function CreateDictionary(ByVal item As Object, ByVal stream As Object) As Object\n Dim result As Object\n Set result = CreateObject(\"Scripting.Dictionary\")\n result.Add \"kept\", \"kept\"\n result.Add \"item\", item\n result.Add \"stream\", stream\n Set CreateDictionary = result\nEnd Function\nPublic Function HoldComponents(ByVal item As Object, ByVal collection As Object, ByVal dictionary As Object) As String\n Set HeldClass = item\n Set HeldCollection = collection\n Set HeldDictionary = dictionary\n HoldComponents = \"held\"\nEnd Function\nPublic Function ProbeClass(ByVal route As Long) As String\n Dim answer As Variant\n On Error GoTo failed\n If route = 0 Then\n  answer = HeldClass.Report\n Else\n  answer = CallByName(HeldClass, \"Report\", vbMethod)\n End If\n ProbeClass = \"value|\" & CStr(answer)\n Exit Function\nfailed:\n ProbeClass = \"error|\" & CStr(Err.Number) & \"|\" & Err.Description & \"|\" & Err.Source & \"|\" & Err.HelpFile & \"|\" & CStr(Err.HelpContext)\nEnd Function\nPublic Function ProbeComponents() As String\n On Error GoTo failed\n ProbeComponents = CStr(HeldCollection.Count) & \"|\" & HeldCollection(1) & \"|\" & CStr(HeldCollection(2) Is HeldClass) & \"|\" & CStr(HeldCollection(3) Is HeldStream) & \"|\" & CStr(HeldDictionary.Count) & \"|\" & HeldDictionary(\"kept\") & \"|\" & CStr(HeldDictionary(\"item\") Is HeldClass) & \"|\" & CStr(HeldDictionary(\"stream\") Is HeldStream)\n Exit Function\nfailed:\n ProbeComponents = \"error|\" & CStr(Err.Number) & \"|\" & Err.Description & \"|\" & Err.Source & \"|\" & CStr(Err.HelpContext)\nEnd Function\nPublic Function HoldFresh(ByVal item As Object) As String\n Set FreshClass = item\n HoldFresh = \"held\"\nEnd Function\nPublic Function ProbeFresh() As String\n On Error GoTo failed\n ProbeFresh = \"value|\" & FreshClass.Report & \"|\" & CStr(HeldClass Is FreshClass)\n Exit Function\nfailed:\n ProbeFresh = \"error|\" & CStr(Err.Number) & \"|\" & Err.Description\nEnd Function\nPublic Function ProbeDead(ByVal route As Long) As String\n Dim answer As Variant, item As Variant, target As Object, picked As Object\n On Error GoTo failed\n Set target = HeldClass\n Select Case route\n Case 0: answer = TypeName(target)\n Case 1: answer = VarType(target)\n Case 2: answer = CStr(target Is Nothing)\n Case 3: answer = CStr(TypeOf target Is RetainedClass)\n Case 4: answer = target.Missing\n Case 5: target.Missing = 42\n Case 6: Set target.Missing = Nothing\n Case 7: answer = target\n Case 8\n  For Each item In target\n   answer = \"enumerated\"\n   Exit For\n  Next\n Case 9: answer = CallByName(target, \"Missing\", vbGet)\n Case 10: answer = CallByName(target, \"Missing\", vbLet, 42)\n Case 11: target.Seed 42\n Case 12: answer = target.Report\n Case 13: answer = CStr(target)\n Case 14: answer = Len(target)\n Case 15: answer = LenB(target)\n Case 16: answer = IsNumeric(target)\n Case 17: answer = IsArray(target)\n Case 18: answer = IsEmpty(target)\n Case 19: answer = IsNull(target)\n Case 20: answer = IsError(target)\n Case 21: answer = IsObject(target)\n Case 22: answer = IsDate(target)\n Case 23: answer = CBool(target)\n Case 24: answer = CLng(target)\n Case 25: answer = CVar(target)\n Case 26: answer = CVErr(target)\n Case 27: answer = TypeName(CVar(target))\n Case 28: answer = VarType(CVar(target))\n Case 29: answer = ObjPtr(target)\n Case 30: answer = IIf(True, target, Nothing)\n Case 31\n  Set picked = IIf(True, target, Nothing)\n  answer = CStr(picked Is target)\n Case 32\n  Set picked = target\n  answer = TypeName(picked)\n Case 33: target.Missing(1) = 42\n Case 34: Set target.Missing(1) = Nothing\n Case 35: answer = target.Missing(1)\n Case 36: answer = CallByName(target, \"Missing\", vbSet, Nothing)\n Case 37: answer = CallByName(target, \"Report\", vbGet)\n End Select\n ProbeDead = \"value|\" & TypeName(answer) & \"|\" & CStr(answer)\n Exit Function\nfailed:\n ProbeDead = \"error|\" & CStr(Err.Number) & \"|\" & Err.Description & \"|\" & Err.Source & \"|\" & Err.HelpFile & \"|\" & CStr(Err.HelpContext)\nEnd Function\nVERSION 1.0 CLASS\nAttribute VB_Name = \"RetainedClass\"\nAttribute VB_Exposed = True\nOption Explicit\nPrivate stored As Long\nPublic Sub Seed(ByVal value As Long)\n stored = value\nEnd Sub\nPublic Function Report() As String\n stored = stored + 1\n Report = CStr(stored) & \"|\" & CStr(IOOwner.Counter(-1)) & \"|\" & ThisWorkbook.Name\nEnd Function\n";
const endLocalSource="Attribute VB_Name = \"OwnerDriver\"\nOption Explicit\nPublic Counter As Long\nPublic TraceLog As String\nPublic Function MakeOwner() As Object\n Set MakeOwner = New SharedName\nEnd Function\nPublic Function ReadTrace() As String\n ReadTrace = TraceLog\nEnd Function\nPublic Sub EndOwner()\n End\nEnd Sub\nPublic Function EndLocalProbe(ByVal mode As Long) As String\n On Error Resume Next\n TraceLog = \"\"\n EndLocalInner mode\n EndLocalProbe = CStr(Err.Number) & \"|\" & CStr(Application.StatusBar)\nEnd Function\nPrivate Sub EndLocalInner(ByVal mode As Long)\n Dim held As Object\n Set held = MakeOwner()\n Application.StatusBar = \"before-local-exit\"\n Select Case mode\n Case 1: End\n Case 2: Err.Raise 5\n End Select\nEnd Sub\nVERSION 1.0 CLASS\nAttribute VB_Name = \"SharedName\"\nAttribute VB_Exposed = True\nPrivate Peer As Object\nPrivate Serial As Long\nPrivate TextValue As String\nPrivate Sub Class_Initialize()\n OwnerDriver.Counter = OwnerDriver.Counter + 1\n Serial = OwnerDriver.Counter\n TextValue = \"A-init\"\n OwnerDriver.TraceLog = OwnerDriver.TraceLog & \"I\" & CStr(Serial) & \";\"\nEnd Sub\nPublic Function InspectOwner() As String\n InspectOwner = \"A|\" & CStr(OwnerDriver.Counter) & \"|\" & CStr(Serial) & \"|\" & TextValue & \"|\" & ThisWorkbook.Name & \"|\" & ActiveWorkbook.Name\nEnd Function\nPublic Property Get Label() As String\n Label = TextValue\nEnd Property\nPublic Property Let Label(ByVal value As String)\n TextValue = value\n OwnerDriver.TraceLog = OwnerDriver.TraceLog & \"S\" & CStr(Serial) & \";\"\nEnd Property\nPrivate Sub Class_Terminate()\n Application.StatusBar = \"term-A-entered\"\n OwnerDriver.TraceLog = OwnerDriver.TraceLog & \"T\" & CStr(Serial) & \";\"\n Application.StatusBar = \"term-A-done\"\nEnd Sub\nPublic Function Touch(ByRef value As Long) As String\n value = value + 5\n Touch = \"A|\" & TypeName(value) & \"|\" & CStr(value)\nEnd Function\nPublic Function Ping() As String\n OwnerDriver.Counter = OwnerDriver.Counter + 10\n Ping = \"A:\" & CStr(OwnerDriver.Counter)\nEnd Function\nPublic Function Reentrant(ByVal other As Object) As String\n Reentrant = \"A(\" & other.Ping() & \")|\" & CStr(OwnerDriver.Counter)\nEnd Function\nPublic Function HoldPeer(ByVal other As Object) As String\n Set Peer = other\n HoldPeer = TypeName(Peer)\nEnd Function\nPublic Function ErrRoundTrip(ByVal mode As Long) As String\n Select Case mode\n Case 0\n  On Error Resume Next\n  Err.Raise 5\n  ErrRoundTrip = \"raised|\" & CStr(Err.Number)\n Case 1\n  On Error GoTo handled\n  Err.Raise 5\n Case 2\n  Err.Raise -2147220803, \"A.Owner\", \"A.Fail\", \"A.Help\", 71\n Case 3\n  ErrRoundTrip = \"incoming|\" & CStr(Err.Number) & \"|\" & Err.Description & \"|\" & Err.Source\n Case 4\n  ErrRoundTrip = \"clearfrom|\" & CStr(Err.Number)\n  Err.Clear\n End Select\n Exit Function\nhandled:\n ErrRoundTrip = \"handled|\" & CStr(Err.Number)\n Exit Function\nEnd Function\n";
const endDispatchGold={"before":[{"route":0,"value":"value|String|RetainedClass"},{"route":1,"value":"value|Long|9"},{"route":2,"value":"value|String|False"},{"route":4,"number":438,"context":1000438},{"route":5,"number":438,"context":1000438},{"route":6,"number":438,"context":1000438},{"route":7,"number":438,"context":1000438},{"route":8,"number":438,"context":1000438},{"route":9,"number":438,"context":1000438},{"route":10,"number":438,"context":1000438},{"route":11,"value":"value|Empty|"},{"route":12,"value":"value|String|43|10|A.xlsm"},{"route":13,"number":438,"context":1000438},{"route":14,"number":438,"context":1000438},{"route":15,"number":438,"context":1000438},{"route":16,"value":"value|Boolean|False"},{"route":17,"value":"value|Boolean|False"},{"route":18,"value":"value|Boolean|False"},{"route":19,"value":"value|Boolean|False"},{"route":20,"value":"value|Boolean|False"},{"route":21,"value":"value|Boolean|True"},{"route":22,"value":"value|Boolean|False"},{"route":23,"number":438,"context":1000438},{"route":24,"number":438,"context":1000438},{"route":25,"number":438,"context":1000438},{"route":26,"number":438,"context":1000438},{"route":27,"value":"value|String|RetainedClass"},{"route":28,"value":"value|Long|9"},{"route":30,"number":438,"context":1000438},{"route":31,"value":"value|String|True"},{"route":32,"value":"value|String|RetainedClass"},{"route":33,"number":438,"context":1000438},{"route":34,"number":438,"context":1000438},{"route":35,"number":438,"context":1000438},{"route":36,"number":438,"context":1000438},{"route":37,"number":450,"context":1000450}],"after":[{"route":0,"value":"value|String|Object"},{"route":1,"value":"value|Long|9"},{"route":2,"value":"value|String|False"},{"route":4,"number":-2147418105,"context":1000440},{"route":5,"number":-2147418105,"context":1000440},{"route":6,"number":-2147418105,"context":1000440},{"route":7,"number":-2147418105,"context":1000440},{"route":8,"number":-2147418105,"context":1000440},{"route":9,"number":440,"context":1000440},{"route":10,"number":440,"context":1000440},{"route":11,"number":-2147418105,"context":1000440},{"route":12,"number":-2147418105,"context":1000440},{"route":13,"number":-2147418105,"context":1000440},{"route":14,"number":-2147418105,"context":1000440},{"route":15,"number":-2147418105,"context":1000440},{"route":16,"value":"value|Boolean|False"},{"route":17,"value":"value|Boolean|False"},{"route":18,"value":"value|Boolean|False"},{"route":19,"value":"value|Boolean|False"},{"route":20,"value":"value|Boolean|False"},{"route":21,"value":"value|Boolean|True"},{"route":22,"value":"value|Boolean|False"},{"route":23,"number":-2147418105,"context":1000440},{"route":24,"number":-2147418105,"context":1000440},{"route":25,"number":-2147418105,"context":1000440},{"route":26,"number":-2147418105,"context":1000440},{"route":27,"value":"value|String|Object"},{"route":28,"value":"value|Long|9"},{"route":30,"number":-2147418105,"context":1000440},{"route":31,"value":"value|String|True"},{"route":32,"value":"value|String|Object"},{"route":33,"number":-2147418105,"context":1000440},{"route":34,"number":-2147418105,"context":1000440},{"route":35,"number":-2147418105,"context":1000440},{"route":36,"number":440,"context":1000440},{"route":37,"number":440,"context":1000440}]};
const endRpcBody="\u547c\u3073\u51fa\u3057\u5148 (\u30b5\u30fc\u30d0\u30fc \u30a2\u30d7\u30ea\u30b1\u30fc\u30b7\u30e7\u30f3\u3067\u306f\u306a\u304f\u3001\u30b5\u30fc\u30d0\u30fc) \u306f\u6d88\u5931\u3057\u3001\u5229\u7528\u3067\u304d\u307e\u305b\u3093\u3002\u3059\u3079\u3066\u306e\u63a5\u7d9a\u306f\u7121\u52b9\u3067\u3059\u3002\u547c\u3073\u51fa\u3057\u304c\u5b9f\u884c\u3055\u308c\u3066\u3044\u308b\u53ef\u80fd\u6027\u304c\u3042\u308a\u307e\u3059\u3002";
async function verifyEndOwnership(route){
 let calls=0;
 for(const local of [false,true]){
  const source=local?endLocalSource:endGraphSource;
  const packet=wasm.save_spreadsheet_vba_source(groups[0].bytes,source,undefined);
  let book=wasm.parse_spreadsheet(packet);
  const session=route==='direct'?new wasm.SpreadsheetVbaSession():undefined;
  if(!session)runner.closeVbaProject();
  const input='C:\\end-owner\\input.txt';
  const options={projectName:'VBAProject',errorLanguage:'ja-JP',vbaHelpFile:'vba-help.chm',hostAnsiCodePage:932,
   documentModules:wasm.spreadsheet_vba_document_modules(packet),
   fileContext:{packageBytes:packet,fullName:'C:\\end-owner\\A.xlsm',readOnly:true,
    folders:['C:\\end-owner'],files:[{path:input,bytes:Array.from(Buffer.from('first\r\nsecond\r\nthird\r\n'))}]}};
  async function execute(action,procedure,args=[],kind=1){
   let execution;
   try{
    if(action==='run')execution=await(session?session.run(book,source,procedure,args,0,'A.xlsm',options)
     :runner.runVbaInWorker(book,source,procedure,args,0,'A.xlsm',options,packet));
    else if(action==='invoke'){
     const[reference,member]=procedure;
     execution=await(session?session.invoke_member(book,source,{reference,member,kind,args},0,'A.xlsm',options)
      :runner.invokeVbaMemberInWorker(book,source,reference,member,kind,args,0,'A.xlsm',options,packet));
    }else{
     execution=await(session?session.release_all_objects(book,0,'A.xlsm',{})
      :runner.releaseAllVbaObjectsInWorker(book,0,'A.xlsm',{},packet));
    }
   }catch(error){if(!error?.execution)throw error;execution=error.execution;}
   calls++;book=execution.workbook;return execution;
  }
  async function run(procedure,args=[]){
   const execution=await execute('run',procedure,args);
   assert.equal(execution.error,undefined,procedure);return execution.result;
  }
  async function invoke(object,member,kind=1,args=[]){
   const execution=await execute('invoke',[object.value.reference,member],args,kind);
   assert.equal(execution.error,undefined,member);return execution.result;
  }

  async function verifyDeadOperations(stage){
   // Foreign-project TypeOf identity and native addresses need separate host state.
   for(const sample of endDispatchGold[stage]){
    let expected=sample.value;
    if(sample.number){
     const description=sample.number===-2147418105?textStreamDescriptions[440]+'\n'+endRpcBody+' ':textStreamDescriptions[sample.number];
     expected='error|'+sample.number+'|'+description+'|VBAProject|vba-help.chm|'+sample.context;
    }
    assert.deepEqual(await run('IOOwner.ProbeDead',[sample.route]),{type:'string',value:expected},stage+' dead-object operation '+sample.route);
   }
  }
  try{
   if(local){
    // StatusBar exposes termination even though End also clears TraceLog.
    for(const mode of [0,1,2]){
     const execution=await execute('run','OwnerDriver.EndLocalProbe',[mode]);
     if(mode===1){
      assert.equal(execution.error.kind,'terminated');
      assert.equal(execution.status_bar,'before-local-exit');
     }else{
      assert.equal(execution.error,undefined);
      assert.deepEqual(execution.result,{type:'string',value:mode===0?'0|term-A-done':'5|term-A-done'});
      assert.equal(execution.status_bar,'term-A-done');
     }
     assert.deepEqual(await run('OwnerDriver.ReadTrace'),{type:'string',value:mode===1?'':'I1;T1;'});
    }
   }else{
    await run('IOOwner.Counter',[10]);
    await run('IOOwner.FileState',[1,input,1,'']);
    assert.deepEqual(await run('IOOwner.DirectoryProbe',['C:\\end-owner\\*.txt']),{type:'string',value:'dir|input.txt'});
    const old=await run('IOOwner.CreateClass');
    const stream=await run('IOOwner.CreateTextStream',[input]);
    const collection=await run('IOOwner.CreateCollection',[old,stream]);
    const dictionary=await run('IOOwner.CreateDictionary',[old,stream]);
    await run('IOOwner.HoldComponents',[old,collection,dictionary]);
    await run('IOOwner.HoldStream',[stream]);
    assert.deepEqual(await invoke(stream,'ReadLine'),{type:'string',value:'first'});
    await verifyDeadOperations('before');
    const ended=await execute('run','IOOwner.TriggerStop',[1]);
    assert.equal(ended.error.kind,'terminated');
    for(const object of [old,stream,collection,dictionary])assert.ok(ended.reference_state.tokens.includes(object.value.reference));
    assert.deepEqual(await run('IOOwner.Counter',[-1]),{type:'integer',value:0});
    assert.deepEqual(await run('IOOwner.FileState',[0,input,1,'']),{type:'string',value:'free|1|256'});
    assert.deepEqual(await run('IOOwner.DirectoryProbe',['']),{type:'string',value:'error|5|'+textStreamDescriptions[5]+'|VBAProject|vba-help.chm|1000005'});
    assert.deepEqual(await invoke(stream,'ReadLine'),{type:'string',value:'second'});
    assert.deepEqual(await invoke(collection,'Count'),{type:'integer',value:3});
    assert.deepEqual(await invoke(dictionary,'Count',2),{type:'integer',value:3});
    for(const[container,index]of[[collection,2],[dictionary,'item']]){
     const alias=await invoke(container,'Item',1,[index]);
     assert.equal(alias.value.handle,old.value.handle);
    }
    for(const[container,index]of[[collection,3],[dictionary,'stream']]){
     const alias=await invoke(container,'Item',1,[index]);
     assert.equal(alias.value.handle,stream.value.handle);
    }
    const disconnected=await execute('invoke',[old.value.reference,'Report']);
    assert.equal(disconnected.error.kind,'runtime');
    await run('IOOwner.HoldComponents',[old,collection,dictionary]);
    assert.deepEqual(await run('IOOwner.ProbeClass',[0]),{type:'string',value:'error|-2147418105|'+textStreamDescriptions[440]+'\n'+endRpcBody+' |VBAProject|vba-help.chm|1000440'});
    assert.deepEqual(await run('IOOwner.ProbeClass',[1]),{type:'string',value:'error|440|'+textStreamDescriptions[440]+'|VBAProject|vba-help.chm|1000440'});
    await verifyDeadOperations('after');
    const fresh=await run('IOOwner.CreateClass');
    assert.notEqual(fresh.value.handle,old.value.handle);
    assert.deepEqual(await invoke(fresh,'Report'),{type:'string',value:'78|0|A.xlsm'});
    assert.deepEqual(await run('IOOwner.ProbeClass',[1]),{type:'string',value:'error|440|'+textStreamDescriptions[440]+'|VBAProject|vba-help.chm|1000440'});
    const newStream=await run('IOOwner.CreateTextStream',[input]);
    assert.notEqual(newStream.value.handle,stream.value.handle);
    assert.deepEqual(await invoke(stream,'ReadLine'),{type:'string',value:'third'});
    assert.deepEqual(await invoke(newStream,'ReadLine'),{type:'string',value:'first'});
    await invoke(stream,'Close');
    assert.deepEqual(await invoke(newStream,'ReadLine'),{type:'string',value:'second'});
   }
   const released=await execute('release');
   assert.equal(released.error,undefined);assert.deepEqual(released.reference_state.tokens,[]);
  }finally{session?.free();if(!session)runner.closeVbaProject();}
 }
 return calls;
}
let endDirectCalls,endWorkerCalls;
try{
 endDirectCalls=await verifyEndOwnership('direct');
 globalThis.Worker=BrowserWorker;
 endWorkerCalls=await verifyEndOwnership('worker');
}finally{runner.closeVbaProject();globalThis.Worker=originalWorker;}
assert.equal(endDirectCalls,endWorkerCalls);
assert.equal(endDirectCalls,113);


// Switch editor views through one actual application, including returning to
// owner zero and invoking an object retained before leaving that workbook.
const selectionSource="Attribute VB_Name = \"Driver\"\nPrivate count As Long\nPrivate kept As Object\nPrivate heldBook As Workbook\nPublic Function Seed() As Object\n SeedLocal 10\n Dim peer As Workbook\n Set peer = Workbooks.Open(\"C:\\B.xlsm\", , True)\n Application.Run \"'B.xlsm'!Driver.SeedLocal\", 20\n Application.StatusBar = \"shared\"\n Kill \"C:\\a1.txt\"\n Set Seed = kept\nEnd Function\nPublic Sub SeedLocal(ByVal start As Long)\n count = start\n Set kept = New Counter\n kept.Data = start * 10\n Set heldBook = ThisWorkbook\n ThisWorkbook.Worksheets(1).Range(\"A1\").Value = start\n Open \"C:\\\" & ThisWorkbook.Name & \".dat\" For Binary Access Read Write Shared As #1\n Put #1, , CByte(start)\n Dim first As String\n first = Dir(\"C:\\\" & LCase(Left(ThisWorkbook.Name, 1)) & \"*.txt\")\nEnd Sub\nPublic Function Read() As String\n Read = CStr(count) & \":\" & CStr(kept.Data)\nEnd Function\nPublic Function Probe(ByVal peer As String) As String\n count = count + 1\n kept.Data = kept.Data + 1\n Put #1, , CByte(count)\n ThisWorkbook.Worksheets(1).Range(\"A1\").Value = count\n Probe = ThisWorkbook.Name & \"|\" & CStr(count) & \"|\" & CStr(kept.Data) & \"|\" & CStr(Loc(1)) & \"|\" & Dir() & \"|\" & CStr(heldBook Is ThisWorkbook) & \"|\" & ActiveWorkbook.Name & \"|\" & Application.Run(\"'\" & peer & \"'!Driver.Read\") & \"|\" & CStr(Application.StatusBar)\nEnd Function\nVERSION 1.0 CLASS\nAttribute VB_Name = \"Counter\"\nPublic Data As Long\n";
const selectionSeed=new Uint8Array(await readFile(path.join(root,'tests/fixtures/vba/array_argument_rank.xlsm')));
const selectionPackage=wasm.save_spreadsheet_vba_source(selectionSeed,selectionSource);
async function verifyWorkbookSelection(mode){
 const session=mode==='direct'?new wasm.SpreadsheetVbaSession():null;
 let book=wasm.parse_spreadsheet(selectionSeed),source=selectionSource,fileName='A.xlsm',projectName='RootProject',calls=0;
 const files=[{path:'C:\\B.xlsm',bytes:selectionPackage},
  ...['a1','a2','b1','b2'].map(name=>({path:'C:\\'+name+'.txt',bytes:[111,114,105,103,105,110,97,108]}))];
 async function run(procedure,args=[]){
  const options={projectName,fileContext:{files}};
  const execution=session?session.run(book,source,procedure,args,0,fileName,options)
   :await runner.runVbaInWorker(book,source,procedure,args,0,fileName,options);
  assert.equal(execution.error,undefined,mode+' '+procedure);
  calls++;book=execution.workbook;return execution;
 }
 async function select(identity){
  const view=session?session.select_workbook(identity):await runner.selectVbaWorkbookInWorker(identity);
  calls++;book=view.selected.workbook;source=view.selected.source;fileName=view.selected.name;projectName=view.selected.project_name;
  assert.equal(view.selected.identity,identity);
  assert.ok(!view.file_system_state.files.some(file=>file.path.toLowerCase()==='c:\\a1.txt'));
  return view;
 }
 try{
  const seeded=await run('Driver.Seed');
  const reference=seeded.result.value.reference;
  const invalid=()=>session?session.select_workbook('999'):runner.selectVbaWorkbookInWorker('999');
  await assert.rejects(async()=>invalid(),/workbook is not open/);
  for(const [identity,name,peer,expected] of [
   ['1','B.xlsm','A.xlsm','B.xlsm|21|201|2|b2.txt|True|B.xlsm|10:100|shared'],
   ['0','A.xlsm','B.xlsm','A.xlsm|11|101|2|a2.txt|True|A.xlsm|21:201|shared'],
   ['1','B.xlsm','A.xlsm','B.xlsm|22|202|3||True|B.xlsm|11:101|shared'],
  ]){
   const view=await select(identity);
   assert.equal(view.selected.name,name);
   assert.equal(view.selected.read_only,identity==='1');
   assert.equal(view.other_workbooks.length,1);
   if(!session)assert.ok(runner.getVbaObjectReferences().includes(reference));
   const outcome=await run('Driver.Probe',[peer]);
   assert.deepEqual(outcome.result,{type:'string',value:expected},mode+' '+name);
   assert.equal(outcome.active_workbook,identity);
   assert.ok(!outcome.file_system_state.files.some(file=>file.path.toLowerCase()==='c:\\a1.txt'));
  }
  const held=session?session.invoke_member(book,source,{reference,member:'Data',kind:2,args:[]},0,fileName,{projectName})
   :await runner.invokeVbaMemberInWorker(book,source,reference,'Data',2,[],0,fileName,{projectName});
  calls++;book=held.workbook;
  assert.equal(held.error,undefined);
  assert.deepEqual(held.result,{type:'integer',value:101});
  if(!session){
   assert.ok(runner.getVbaObjectReferences().includes(reference));
   const edited=source+"\n' pending edit in B\n";
   assert.equal(runner.recordVbaSourceChange(source,edited),true);
   await select('0');
   const restored=await select('1');
   assert.equal(restored.selected.source,edited,'B pending edits survive A selection');
   assert.equal(runner.getVbaSourceForSave(),edited);
   assert.deepEqual((await run('Driver.Read')).result,{type:'string',value:'22:202'});
  }
  const rootView=await select('0');
  assert.equal(rootView.selected.project_name,'RootProject');
  assert.deepEqual((await run('Driver.Read')).result,{type:'string',value:'11:101'});
 }finally{session?.free();if(!session)runner.closeVbaProject();}
 return calls;
}
let workbookSelectionDirectCalls,workbookSelectionWorkerCalls;
try{
 workbookSelectionDirectCalls=await verifyWorkbookSelection('direct');
 globalThis.Worker=BrowserWorker;
 workbookSelectionWorkerCalls=await verifyWorkbookSelection('worker');
}finally{runner.closeVbaProject();globalThis.Worker=originalWorker;}
assert.equal(workbookSelectionDirectCalls,10);
assert.equal(workbookSelectionWorkerCalls,13);

console.log(JSON.stringify({native_cases:163,direct,worker,text_stream_cases:360,text_stream_close_reset_cases:4,text_stream_direct_calls:textStreamDirectCalls,text_stream_worker_calls:textStreamWorkerCalls,end_direct_calls:endDirectCalls,end_worker_calls:endWorkerCalls,workbook_selection_direct_calls:workbookSelectionDirectCalls,workbook_selection_worker_calls:workbookSelectionWorkerCalls,complete_results:true,post_file_bytes_and_folders:true,post_attributes:true,deletion_outputs:true,passed:true}));
