// SPDX-License-Identifier: MPL-2.0
// Replay the complete Native OLE/note source, copy and saved/reopened vectors.
// The CC BY-SA 4.0 corpus shares the repository's two original input fixtures.
// Usage: node web/vba-ole-copy-state.test.mjs pkg [capture.json] [corpus.json] [root] [expectedConditions]
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {createHash} from 'node:crypto';
import {isDeepStrictEqual} from 'node:util';

const here=path.dirname(fileURLToPath(import.meta.url));
const root=path.resolve(process.argv[5]??path.join(here,'..'));
const pkg=path.resolve(process.argv[2]??path.join(root,'crates/oxidocs-wasm/pkg'));
const output=process.argv[3]?path.resolve(process.argv[3]):null;
const corpusPath=path.resolve(process.argv[4]??path.join(root,'tests/fixtures/vba/ole_copy_native.json'));
if(output)assert(!fs.existsSync(output),'Capture must be a new file');
const hash=bytes=>createHash('sha256').update(bytes).digest('hex');
const corpus=JSON.parse(fs.readFileSync(corpusPath,'utf8'));
assert.equal(corpus.license,'CC-BY-SA-4.0');
const conditions=process.argv[6]===undefined?69:Number(process.argv[6]);
assert([69,18,33].includes(conditions));assert.equal(corpus.conditions,conditions);
const inputPath=path.join(root,corpus.input_fixture),callerPath=path.join(root,corpus.caller_fixture);
const input=new Uint8Array(fs.readFileSync(inputPath)),caller=new Uint8Array(fs.readFileSync(callerPath));
assert.equal(hash(input),corpus.input_sha256);assert.equal(hash(caller),corpus.caller_sha256);
const jsPath=path.join(pkg,'oxidocs_wasm.js'),wasmPath=path.join(pkg,'oxidocs_wasm_bg.wasm');
const js=fs.readFileSync(jsPath),binary=fs.readFileSync(wasmPath);
const wasm=await import('data:text/javascript;base64,'+js.toString('base64'));
await wasm.default({module_or_path:binary});
const guards={};
for(const p of [fileURLToPath(import.meta.url),corpusPath,inputPath,callerPath,jsPath,wasmPath])guards[p]=hash(fs.readFileSync(p));
const report={completed:false,passed:false,conditions,expected_observations:conditions*2,
    scope:corpus.scope,source_guards:guards,cases:[]};
const save=()=>{if(output)fs.writeFileSync(output,JSON.stringify(report)+'\n');};
save();
const seen=new Set();
for(const group of corpus.groups){
    assert.equal(group.reference.independent_orders.length,2);
    for(const test of group.cases){
        assert(!seen.has(group.id+'/'+test.mode));seen.add(group.id+'/'+test.mode);
        for(const route of ['direct','session']){
            const row={group:group.id,mode:test.mode,route,expected:test.expected,passed:false};
            let session;
            try{
                const bytes=group.separate_caller?caller:input;
                const options={errorLanguage:'ja-JP',hostAnsiCodePage:932,
                    fileContext:{packageBytes:bytes,fullName:'C:\\work\\original.xlsx',readOnly:true,
                        editRevision:0,folders:['C:\\work'],
                        files:group.separate_caller?[{path:'C:\\work\\provider.xlsx',bytes:input},...(group.sample_text!==undefined?[{path:'C:\\work\\sample.txt',bytes:new TextEncoder().encode(group.sample_text)}]:[])]:[]}};
                let result;
                if(route==='session'){
                    session=new wasm.SpreadsheetVbaSession();
                    result=session.run(wasm.parse_spreadsheet(bytes),group.program,'Probe',[test.mode],0,'C:\\work\\original.xlsx',options);
                }else result=wasm.run_spreadsheet_vba_with_outcome(wasm.parse_spreadsheet(bytes),group.program,'Probe',[test.mode],0,'C:\\work\\original.xlsx',options);
                row.actual=result.result;row.error=result.error??null;row.caller_changed=result.workbook_changed;
                row.files=(result.files??[]).map(file=>({path:file.path,bytes_base64:Buffer.from(file.bytes).toString('base64')}));
                row.expected_paths=['C:\\work\\mixed-copy.xlsx','C:\\work\\mixed-source.xlsx'];
                row.actual_paths=row.files.map(file=>file.path).sort();
                row.full_vector_equal=isDeepStrictEqual(row.actual,row.expected);
                // Independent callers must retain their original workbook state.
                // The self-modifying source cases record the caller flag without
                // claiming a separate Native caller-state oracle.
                row.passed=row.full_vector_equal&&row.error===null
                    &&isDeepStrictEqual(row.actual_paths,row.expected_paths)
                    &&(!group.separate_caller||row.caller_changed===false);
            }catch(error){row.exception={name:error.name,message:error.message};}
            finally{session?.free();}
            report.cases.push(row);save();
        }
    }
}
assert.equal(seen.size,conditions);assert.equal(report.cases.length,conditions*2);
report.completed=true;report.guards_unchanged=Object.entries(guards).every(([p,h])=>hash(fs.readFileSync(p))===h);
report.passed=report.guards_unchanged&&report.cases.every(row=>row.passed);save();
console.log(JSON.stringify({completed:true,passed:report.passed,conditions,observations:conditions*2,
    matched:report.cases.filter(row=>row.passed).length,different:report.cases.filter(row=>!row.passed).length,
    exceptions:report.cases.filter(row=>row.exception).length}));
process.exitCode=report.passed?0:1;
