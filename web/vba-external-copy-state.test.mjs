// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.
import assert from 'node:assert/strict';
import { readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { inflateRawSync } from 'node:zlib';
import { createHash } from 'node:crypto';

const root=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
const pkg=path.resolve(process.argv[2] || path.join(root,'crates/oxidocs-wasm/pkg'));
const capture=process.argv[3];
const overrides=process.argv[4] ? JSON.parse(await readFile(process.argv[4],'utf8')) : {};
const hash=bytes=>createHash('sha256').update(bytes).digest('hex');
const specs=[
  {
    "label": "linked",
    "fixture": "copy_linked_source_native8.zip",
    "sha256": "04eb73ca0b40cc6bb913db5529d91a56a9c1ee79d66f5090b13cb9a22738df0b"
  },
  {
    "label": "link_enumeration2",
    "fixture": "external_link_save_order_native8.zip",
    "sha256": "60a259d0b85de6732e781ebf5b9e105ec263f19298abdebfe98b0d68f33387a0"
  },
  {
    "label": "link_enumeration3",
    "fixture": "external_link_load_order3_native8.zip",
    "sha256": "a5bd6e5ded21b4f214d6881136afbb14f51218901cd4d7b5c6ad23b9e4f2d402"
  },
  {
    "label": "external_formula_setters",
    "fixture": "external_formula_setters_native8.zip",
    "sha256": "ef44e905b1b4929baa52de73b83269519a4f9a778ac6dd3c15a254b12eac1bde"
  },
  {
    "label": "external_formula_other_setters",
    "fixture": "external_formula_other_setters_native8.zip",
    "sha256": "8b051afd0645b5c1740ee011c976360229078267d730ccde9e5f4cb909f2fe1a"
  }
];
const js=await readFile(path.join(pkg,'oxidocs_wasm.js'));
const binary=await readFile(path.join(pkg,'oxidocs_wasm_bg.wasm'));
const wasm=await import('data:text/javascript;base64,'+js.toString('base64'));
await wasm.default({module_or_path:binary});

// These reviewed archives are small, non-ZIP64 packages. No files are extracted.
function zipPart(bytes,target) {
    const zip=Buffer.from(bytes),end=zip.lastIndexOf(Buffer.from([80,75,5,6]));
    assert(end>=0);
    let at=zip.readUInt32LE(end+16);
    for(let i=0;i<zip.readUInt16LE(end+10);i++) {
        assert.equal(zip.readUInt32LE(at),0x02014b50);
        const length=zip.readUInt16LE(at+28);
        if(zip.subarray(at+46,at+46+length).toString('utf8')===target) {
            const local=zip.readUInt32LE(at+42);
            assert.equal(zip.readUInt32LE(local),0x04034b50);
            const start=local+30+zip.readUInt16LE(local+26)+zip.readUInt16LE(local+28);
            const data=zip.subarray(start,start+zip.readUInt32LE(at+20));
            const method=zip.readUInt16LE(at+10);
            assert(method===0 || method===8);
            const decoded=method===8 ? inflateRawSync(data) : data;
            assert.equal(decoded.length,zip.readUInt32LE(at+24));
            return decoded;
        }
        at+=46+length+zip.readUInt16LE(at+30)+zip.readUInt16LE(at+32);
    }
    assert.fail('Missing frozen Native entry: '+target);
}
function externalIdentities(bytes,fullName) {
    return (wasm.parse_spreadsheet(bytes).external_books ?? []).map(book=>{
        let target=book.target ?? book.absolute_target;
        assert.equal(typeof target,'string','Saved external book identity');
        target=decodeURIComponent(target);
        if(target.startsWith('file:///')) target=target.slice(8);
        else if(target.startsWith('file://')) target='\\\\'+target.slice(7);
        target=target.replaceAll('/','\\');
        return path.win32.normalize(path.win32.isAbsolute(target) ? target : path.win32.join(path.win32.dirname(fullName),target));
    });
}
const stringify=value=>JSON.stringify(value,(_key,v)=>typeof v==='bigint'?{$bigint:v.toString()}:v);
const report={completed:false,passed:false,whole_goal_complete:false,
    artifact_hashes:{'oxidocs_wasm.js':hash(js),'oxidocs_wasm_bg.wasm':hash(binary)},
    expected_conditions:40,expected_observations:80,cases:[]};
const persist=async()=>{if(capture) await writeFile(capture,stringify(report)+'\n','utf8');};
await persist();
for(const spec of specs) {
    const archive=await readFile(overrides[spec.label] || path.join(root,'tests/fixtures/vba',spec.fixture));
    assert.equal(hash(archive),spec.sha256,'Frozen Native corpus '+spec.label);
    const controls=JSON.parse(zipPart(archive,'controls.json'));
    const bytes=new Uint8Array(zipPart(archive,'program'));
    const source=wasm.spreadsheet_vba_source(bytes,932);
    const documents=wasm.spreadsheet_vba_document_modules(bytes);
    assert.equal(typeof source,'string');
    assert.deepEqual([...controls.cases.map(c=>c.mode)].sort(),[0,1,2,3,4,5,6,7]);
    for(const route of ['direct','session']) {
        for(const test of controls.cases) {
            let session;
            const record={label:spec.label,mode:test.mode,route,passed:false,
                expected:{type:'string',value:test.result},
                expected_paths:Object.keys(test).filter(k=>k.endsWith('_path')).map(k=>test[k]).sort()};
            try {
                const options={errorLanguage:'ja-JP',hostAnsiCodePage:932,projectName:'OxiNamedProject',
                    vbaHelpFile:'C:\\Program Files\\Common Files\\Microsoft Shared\\VBA\\VBA7.1\\1041\\VbLR6.chm',
                    documentModules:documents,
                    applicationState:{standardFontStartup:{profileUtf16:Array.from({length:controls.startup.length},(_,i)=>controls.startup.charCodeAt(i))}},
                    fileContext:{packageBytes:bytes,fullName:controls.source_path,readOnly:true,editRevision:0}};
                const book=wasm.parse_spreadsheet(bytes);
                let result;
                if(route==='session') {
                    session=new wasm.SpreadsheetVbaSession();
                    result=session.run(book,source,'DrawingIdentityProbe.Measure',[test.mode],0,controls.source_path,options);
                } else {
                    result=wasm.run_spreadsheet_vba_with_outcome(book,source,'DrawingIdentityProbe.Measure',[test.mode],0,controls.source_path,options);
                }
                record.actual=result.result;
                record.error=result.error ?? null;
                assert.equal(typeof result.workbook_changed,'boolean');
                record.caller_changed=result.workbook_changed;
                const files=result.files ?? [];
                record.actual_paths=[...new Set(files.map(file=>file.path))].sort();
                record.files=files.map(file=>({path:file.path,sha256:hash(Buffer.from(file.bytes)),
                    ...(capture ? {bytes_base64:Buffer.from(file.bytes).toString('base64')} : {})}));
                record.full_vector_equal=stringify(record.actual)===stringify(record.expected);
                record.passed=record.full_vector_equal && record.error===null && record.caller_changed===false
                    && stringify(record.actual_paths)===stringify(record.expected_paths);
                if(spec.label.startsWith('link_enumeration')) {
                    record.saved_external_identities=[];
                    for(const key of Object.keys(test).filter(k=>k.endsWith('_path'))) {
                        const role=key.slice(0,-5),file=files.find(f=>f.path===test[key]);
                        const expected=externalIdentities(zipPart(archive,`goldens/forward/${test.mode}/${role}.xlsx`),test[key]);
                        const actual=file ? externalIdentities(file.bytes,test[key]) : null;
                        const equal=stringify(actual)===stringify(expected);
                        record.saved_external_identities.push({role,actual,expected,equal});
                        record.passed &&= equal;
                    }
                }
            } catch(error) {
                record.exception={name:error.name,message:error.message};
                record.passed=false;
            } finally {session?.free();}
            report.cases.push(record);
            await persist();
        }
    }
}
assert.equal(report.cases.length,80);
assert.equal(new Set(report.cases.map(c=>`${c.label}/${c.mode}/${c.route}`)).size,80);
report.completed=true;
report.passed=report.cases.every(c=>c.passed);
await persist();
const failures=report.cases.filter(c=>!c.passed).map(c=>({label:c.label,mode:c.mode,route:c.route,
    exception:c.exception,actual_sha256:hash(Buffer.from(stringify(c.actual) ?? 'undefined')),
    expected_sha256:hash(Buffer.from(stringify(c.expected)))}));
console.log(JSON.stringify({conditions:40,observations:80,matched:80-failures.length,different:failures.length}));
assert.equal(failures.length,0,JSON.stringify(failures.slice(0,12)));
