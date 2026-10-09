# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

"""Add OS/2 average advances to an existing font catalog, without font programs."""
import argparse,gzip,hashlib,json,os
from pathlib import Path
from fontTools.ttLib import TTFont

def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('index',type=Path)
    parser.add_argument('--font-root',action='append',type=Path)
    parser.add_argument('--report',type=Path,required=True)
    args=parser.parse_args()
    roots=args.font_root
    if not roots:
        local=Path(os.environ.get('LOCALAPPDATA',str(Path.home()/'AppData/Local')))
        roots=[Path(os.environ.get('SystemRoot','C:/Windows'))/'Fonts',
               local/'Microsoft/Windows/Fonts',local/'Microsoft/FontCache/4/CloudFonts']
    raw=args.index.read_bytes()
    catalog=json.loads(raw)
    wanted={f['key'].split(':')[0] for f in catalog['faces']}
    matches={}
    sources=[]
    for root in roots:
        if not root.exists():continue
        for path in sorted(root.rglob('*')):
            if not path.is_file() or path.suffix.lower() not in ['.ttf','.otf','.ttc']:continue
            digest=hashlib.sha256(path.read_bytes()).hexdigest()
            sources.append((path,digest))
            if digest in wanted:matches.setdefault(digest,path)
    results=[]
    for face in catalog['faces']:
        parts=face['key'].split(':')
        assert len(parts) in [2,4] and (len(parts)==2 or parts[2]=='instance'),face['key']
        digest,number=parts[:2]
        path=matches.get(digest)
        row={'key':face['key']}
        if path is None:
            row['error']='Matching source font hash not found'
        else:
            with TTFont(path,fontNumber=int(number),lazy=True) as font:
                average=font['OS/2'].xAvgCharWidth if 'OS/2' in font else None
                upm=font['head'].unitsPerEm
                row.update(source_path=str(path),source_sha256=digest,face_index=int(number),average=average,upm=upm)
                if len(parts)==4:
                    instance=font['fvar'].instances[int(parts[3])]
                    row.update(named_instance=int(parts[3]),coordinates=instance.coordinates,
                               average_source='OS/2 header of the exact source face')
                if average is not None and average>0 and upm>0:
                    face['average_width_em']=average/upm
                    row['average_width_em']=face['average_width_em']
                else:row['error']='No positive OS/2 average advance'
        results.append(row)
    # A system update can change a font file's names/signature without changing
    # its stored metrics. Recover those faces only after checking all catalog
    # advances and vertical metrics against a same-name, same-style source face.
    pending={r['key']:r for r in results if 'error' in r}
    by_key={f['key']:f for f in catalog['faces']}
    data=args.index.with_name('font_catalog_metrics.gz').read_bytes()
    for path,digest in sources:
        if not pending:break
        with path.open('rb') as stream:header=stream.read(12)
        count=int.from_bytes(header[8:12],'big') if header[:4]==b'ttcf' else 1
        for number in range(count):
            with TTFont(path,fontNumber=number,lazy=True) as font:
                names=set()
                for record in font['name'].names:
                    if record.nameID in [1,4,16]:
                        try:names.add(record.toUnicode().strip().casefold())
                        except UnicodeError:pass
                for key,row in list(pending.items()):
                    face=by_key[key]
                    if not names.intersection(face['families']+face['full_names']):continue
                    if 'OS/2' not in font:continue
                    os2=font['OS/2']
                    if os2.usWeightClass!=face['weight'] or os2.usWidthClass!=face['width_class']:continue
                    if bool(os2.fsSelection & 1)!=face['italic']:continue
                    raw_face=json.loads(gzip.decompress(data[face['offset']:face['offset']+face['length']]))
                    upm=font['head'].unitsPerEm
                    if upm!=raw_face['units_per_em']:continue
                    hhea=font['hhea']
                    if (hhea.ascent,hhea.descent,hhea.lineGap)!=(raw_face['ascender'],raw_face['descender'],raw_face['line_gap']):continue
                    cmap=font.getBestCmap()
                    hmtx=font['hmtx'].metrics
                    changed=[int(cp) for cp,advance in raw_face['widths'].items()
                             if cmap.get(int(cp)) not in hmtx or hmtx[cmap[int(cp)]][0]!=advance]
                    latin={int(cp):advance for cp,advance in raw_face['widths'].items() if 32<=int(cp)<127}
                    # This attribute is used for preceding ASCII text. A newer
                    # font with changed non-Latin outlines is still an exact
                    # source for that domain when all 95 ASCII advances match.
                    if changed and (len(latin)!=95 or any(cp in latin for cp in changed)):continue
                    average=os2.xAvgCharWidth
                    if average<=0:continue
                    face['average_width_em']=average/upm
                    row.pop('error')
                    row.update(source_path=str(path),source_sha256=digest,face_index=number,average=average,upm=upm,
                               average_width_em=face['average_width_em'],
                               resolution='matching_latin_metrics' if changed else 'matching_catalog_metrics',
                               matched_advances=len(latin) if changed else len(raw_face['widths']),
                               changed_other_advances=len(changed),catalog_source_sha256=key.split(':')[0])
                    pending.pop(key)
    output=json.dumps(catalog,ensure_ascii=True,separators=(',',':'),sort_keys=True).encode('utf8')
    args.index.write_bytes(output)
    report=dict(index_before_sha256=hashlib.sha256(raw).hexdigest(),index_after_sha256=hashlib.sha256(output).hexdigest(),
                faces=len(results),updated=sum('average_width_em' in r for r in results),
                missing=[r for r in results if 'error' in r],results=results)
    args.report.write_text(json.dumps(report,ensure_ascii=True,indent=2),encoding='utf8')
    print(json.dumps({k:v for k,v in report.items() if k not in ['results','missing']}),flush=True)

if __name__=='__main__':main()
