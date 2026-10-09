# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.
"""Inventory installed SFNT faces and export complete numeric metrics.

Scan system, per-user, and Office cloud fonts, including every TTC face.
Output contains name mappings, provenance hashes, and cmap/hmtx/vertical
metrics, never font programs. Ambiguous names and variable fonts are explicit;
this inventory does not certify Word layout or choose substitute fonts.

Requires fontTools. All output is deterministic for identical input files.
"""

import argparse
import copy
import gzip
import hashlib
import io
import json
import os
import shutil
from pathlib import Path
import zipfile
import xml.etree.ElementTree as ET

from fontTools.ttLib import TTCollection, TTFont
from fontTools.varLib.instancer import instantiateVariableFont

from font_geometry import face_geometry

FONT_EXTENSIONS = {'.ttf', '.otf', '.ttc', '.otc'}
NAME_IDS = {1, 2, 4, 6, 16, 17}
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'


def default_roots():
    local = os.environ.get('LOCALAPPDATA')
    roots = []
    if local:
        roots += [Path(local)/'Microsoft/FontCache/4/CloudFonts',
                  Path(local)/'Microsoft/Windows/Fonts']
    roots.append(Path(os.environ.get('SystemRoot', 'C:/Windows'))/'Fonts')
    return roots


def normalize_instance_style(meta):
    """Instancing hmtx does not consistently rewrite OS/2 style bits."""
    coordinates=meta.get('instance_coordinates',{})
    if 'wght' in coordinates:
        meta['weight']=round(coordinates['wght'])
        meta['bold']=coordinates['wght']>=700
    if 'ital' in coordinates:
        meta['italic']=coordinates['ital']>=0.5
    if 'slnt' in coordinates:
        meta['italic']=meta['italic'] or coordinates['slnt']!=0
    return meta


def instance_metadata(font, instance, meta):
    """Retain fvar's names even when STAT omits a named optical size."""
    meta['instance_coordinates']=instance.coordinates
    meta['instance_name']=font['name'].getDebugName(instance.subfamilyNameID)
    meta['names']=[n for n in meta['names'] if n['id'] in {1,2,16,17}]
    family=font['name'].getDebugName(16) or font['name'].getDebugName(1)
    full=family+' '+meta['instance_name']
    meta['full_name']=full
    meta['names'].append(dict(id=4,platform=3,language=1033,value=full))
    if instance.postscriptNameID!=65535:
        ps=font['name'].getDebugName(instance.postscriptNameID)
        if ps:meta['names'].append(dict(id=6,platform=3,language=1033,value=ps))
    # The legacy family groups Regular/Bold/Italic together. Weight and
    # optical-size qualifiers remain part of that group name.
    group=full.split()
    while group and group[-1] in {'Regular','Bold','Italic','Oblique'}:group.pop()
    if group:meta['names'].append(dict(id=1,platform=3,language=1033,value=' '.join(group)))
    normalize_instance_style(meta)
    return meta


def face_data(face):
    names = []
    for record in face['name'].names:
        if record.nameID not in NAME_IDS:
            continue
        value = record.toUnicode().strip()
        if value:
            names.append((record.nameID, record.platformID, record.langID, value))
    names = sorted(set(names))
    cmap = face.getBestCmap() or {}
    # Symbol faces use a Windows symbol cmap rather than a Unicode cmap.
    # Preserve its actual codepoints; do not invent an ASCII mapping.
    if not cmap:
        for table in face['cmap'].tables:
            if table.platformID == 3 and table.platEncID == 0:
                cmap.update(table.cmap)
    h = face['hhea']
    o = face.get('OS/2')
    metrics = dict(
        family=face['name'].getDebugName(1),
        units_per_em=face['head'].unitsPerEm,
        ascender=h.ascent, descender=h.descent, line_gap=h.lineGap,
        win_ascent=o.usWinAscent if o else max(0,h.ascent),
        win_descent=o.usWinDescent if o else max(0,-h.descent),
        typo_ascender=o.sTypoAscender if o else h.ascent,
        typo_descender=o.sTypoDescender if o else h.descent,
        typo_line_gap=o.sTypoLineGap if o else h.lineGap,
        use_typo_metrics=bool(o and o.fsSelection & 128),
        average_width=getattr(o, 'xAvgCharWidth', None),
        widths={str(cp):face['hmtx'].metrics[glyph][0] for cp,glyph in sorted(cmap.items())},
    )
    meta = dict(
        family=metrics['family'], full_name=face['name'].getDebugName(4),
        postscript_name=face['name'].getDebugName(6),
        names=[dict(id=i,platform=p,language=l,value=v) for i,p,l,v in names],
        bold=bool(o.fsSelection & 32) if o else bool(face['head'].macStyle & 1),
        italic=bool(o.fsSelection & 1) if o else bool(face['head'].macStyle & 2),
        weight=o.usWeightClass if o else None,
        width_class=o.usWidthClass if o else None,
        codepage_range1=getattr(o, 'ulCodePageRange1', None),
        average_width_em=(metrics['average_width']/metrics['units_per_em']
                          if metrics['average_width'] and metrics['average_width']>0 else None),
        cmap_count=len(cmap), has_kana=0x3042 in cmap,
        has_ideographs=0x4E00 in cmap,
        symbol_cmap=any(t.platformID==3 and t.platEncID==0 for t in face['cmap'].tables),
        has_gpos='GPOS' in face, has_gsub='GSUB' in face,
        variable_axes=[], variable_instances=[],
    )
    if 'fvar' in face:
        meta['variable_axes']=[dict(tag=x.axisTag,min=x.minValue,default=x.defaultValue,max=x.maxValue) for x in face['fvar'].axes]
        meta['variable_instances']=[dict(name=face['name'].getDebugName(x.subfamilyNameID),coordinates=x.coordinates) for x in face['fvar'].instances]
        meta['variation_scope']='Default coordinates only; named instances require separate instantiation'
    return meta, metrics


def export(roots, output, named_instances=True):
    output.mkdir(parents=True,exist_ok=True)
    faces=[]; metrics={}; errors=[]; seen={}; sources=[]
    for priority,root in enumerate(roots):
        paths=([root] if root.is_file() and root.suffix.lower() in FONT_EXTENSIONS else
               sorted(p for p in root.rglob('*') if p.is_file() and p.suffix.lower() in FONT_EXTENSIONS)) if root.exists() else []
        sources.append(dict(priority=priority,path=str(root),exists=root.exists(),files=len(paths)))
        for path in paths:
            font_bytes=path.read_bytes()
            digest=hashlib.sha256(font_bytes).hexdigest()
            if digest in seen:
                for f in seen[digest]: f['sources'].append(str(path))
                continue
            loaded=None
            stream=io.BytesIO(font_bytes)
            try:
                if path.suffix.lower() in {'.ttc','.otc'}:
                    loaded=TTCollection(stream,lazy=True); fonts=loaded.fonts
                else:
                    loaded=TTFont(stream,lazy=True); fonts=[loaded]
                seen[digest]=[]
                for index,font in enumerate(fonts):
                    try:
                        meta,data=face_data(font)
                        key=digest+':'+str(index)
                        data['glyph_geometry']=face_geometry(font,key)
                        meta.update(key=key,source_sha256=digest,face_index=index,sources=[str(path)],priority=priority)
                        faces.append(meta); metrics[key]=data; seen[digest].append(meta)
                        if named_instances and 'fvar' in font:
                            metric_font=copy.deepcopy(font)
                            # GPOS/GDEF/GSUB do not supply the advances or
                            # vertical metrics exported here. Some system
                            # fonts carry OTL variation records which the
                            # general-purpose outline instancer cannot merge.
                            # Keep HVAR, MVAR and gvar; they DO affect metrics.
                            for tag in ['GPOS','GDEF','GSUB']:
                                if tag in metric_font:del metric_font[tag]
                            for instance_index,instance in enumerate(font['fvar'].instances):
                                try:
                                    instantiated=instantiateVariableFont(metric_font,instance.coordinates,inplace=False,
                                                                       optimize=False,updateFontNames=False)
                                    try:
                                        imeta,idata=face_data(instantiated)
                                        ikey=key+':instance:'+str(instance_index)
                                        # The instancer changes glyf, while the reader still
                                        # points to the original file. Read instantiated bounds
                                        # and preserve the original ssty glyph-name mappings.
                                        idata['glyph_geometry']=face_geometry(
                                            instantiated,ikey,ssty_source=font,stored_glyf=False)
                                        imeta.update(key=ikey,source_sha256=digest,face_index=index,sources=[str(path)],
                                                     priority=priority,instance_coordinates=instance.coordinates,
                                                     instance_name=font['name'].getDebugName(instance.subfamilyNameID))
                                        instance_metadata(font,instance,imeta)
                                        imeta['has_gpos']=meta['has_gpos'];imeta['has_gsub']=meta['has_gsub']
                                        faces.append(imeta);metrics[ikey]=idata;seen[digest].append(imeta)
                                    finally: instantiated.close()
                                except Exception as exc:
                                    errors.append(dict(path=str(path),face_index=index,instance=instance_index,error=str(exc)))
                            metric_font.close()
                    except Exception as exc:
                        errors.append(dict(path=str(path),face_index=index,error=str(exc)))
            except Exception as exc:
                errors.append(dict(path=str(path),error=str(exc)))
            finally:
                if loaded is not None: loaded.close()
                stream.close()
    faces.sort(key=lambda f:(f['priority'],f['sources'][0].casefold(),f['face_index']))
    aliases={}
    for f in faces:
        for name in sorted({n['value'] for n in f['names'] if n['id'] in {1,4,6,16}}):
            aliases.setdefault(name.casefold(),[]).append(f['key'])
    raw=json.dumps(metrics,ensure_ascii=False,sort_keys=True,separators=(',',':')).encode('utf-8')
    payload=gzip.compress(raw,mtime=0)
    (output/'metrics.json.gz').write_bytes(payload)
    catalog=dict(schema=1,sources=sources,faces=faces,aliases=aliases,errors=errors,
                 metrics_sha256=hashlib.sha256(payload).hexdigest(),
                 layout_validation='Inventory only; not Word conformance',
                 uncompressed_bytes=len(raw),compressed_bytes=len(payload))
    (output/'catalog.json').write_text(json.dumps(catalog,ensure_ascii=False,indent=2),encoding='utf-8')
    return catalog


def corpus_audit(paths,catalog,output):
    supported=resolution_names(catalog)
    names={}; errors=[]; documents=0; theme_refs={}; skipped=[]
    for root in paths:
        files=[root] if root.is_file() else sorted(root.rglob('*.docx'))
        for path in files:
            if path.name.startswith('~$'):
                skipped.append(dict(path=str(path),reason='Office owner/lock file'))
                continue
            documents+=1
            try:
                with zipfile.ZipFile(path) as package:
                    used=set()
                    for name in package.namelist():
                        if not name.startswith('word/') or not name.endswith('.xml'):continue
                        tree=ET.fromstring(package.read(name))
                        for e in tree.iter():
                            if e.tag==W+'rFonts':
                                for k in ['ascii','hAnsi','eastAsia','cs']:
                                    if e.get(W+k):used.add(e.get(W+k))
                                for k in ['asciiTheme','hAnsiTheme','eastAsiaTheme','cstheme']:
                                    if e.get(W+k):theme_refs[e.get(W+k)]=theme_refs.get(e.get(W+k),0)+1
                            elif e.tag in {A+'latin',A+'ea',A+'cs',A+'font'}:
                                if e.get('typeface'):used.add(e.get('typeface'))
                    for name in used:names.setdefault(name,[]).append(str(path))
            except Exception as exc:errors.append(dict(path=str(path),error=str(exc)))
    rows=[]
    for name,docs in sorted(names.items()):
        candidates=supported.get(name.casefold(),[])
        rows.append(dict(name=name,documents=len(docs),examples=docs[:3],faces=candidates,
                         status='installed_name' if candidates else 'unresolved_name'))
    report=dict(documents=documents,names=len(rows),rows=rows,errors=errors,skipped=skipped,theme_references=theme_refs,
                scope='Declared font names including styles and theme alternatives; not all are active runs. Unresolved means no installed name, not automatic layout failure.')
    (output/'corpus.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    return report


def resolution_names(catalog):
    """OOXML family/full names; a PostScript name alone is not a family.

    Preserve name ID 6 in the inventory, but do not let it change Word's
    substitution for unknown family spellings such as MS-Mincho.
    """
    names={}
    for face in catalog['faces']:
        for name in sorted({n['value'] for n in face['names'] if n['id'] in {1,4,16}}):
            names.setdefault(name.casefold(),[]).append(face['key'])
    return names


def pack_portable(catalog, output, destination):
    """One gzip member per face; the index permits lazy per-face decoding.

    The portable files contain no workstation paths or font programs. Source
    hashes and collection indices identify the measured font versions.
    """
    metrics=json.loads(gzip.decompress((output/'metrics.json.gz').read_bytes()))
    destination.mkdir(parents=True,exist_ok=True)
    data=bytearray();index=[]
    for face in catalog['faces']:
        face=normalize_instance_style(dict(face))
        raw=metrics[face['key']]
        payload=gzip.compress(json.dumps(raw,sort_keys=True,separators=(',',':')).encode('utf-8'),mtime=0)
        families=sorted({n['value'].lower() for n in face['names'] if n['id'] in {1,16}})
        full_names=sorted({n['value'].lower() for n in face['names'] if n['id']==4})
        index.append(dict(key=face['key'],offset=len(data),length=len(payload),families=families,full_names=full_names,
                          bold=face['bold'],italic=face['italic'],weight=face['weight'] or 400,
                          width_class=face['width_class'] or 5,priority=face['priority'],
                          codepage_range1=face.get('codepage_range1'),
                          average_width_em=face.get('average_width_em')))
        data.extend(payload)
    (destination/'font_catalog_metrics.gz').write_bytes(data)
    manifest=dict(schema=1,metrics_sha256=hashlib.sha256(data).hexdigest(),faces=index)
    (destination/'font_catalog_index.json').write_text(json.dumps(manifest,sort_keys=True,separators=(',',':')),encoding='utf-8')
    return manifest


def main():
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('--font-root',type=Path,action='append')
    p.add_argument('--corpus',type=Path,action='append',default=[])
    p.add_argument('--output',type=Path,required=True)
    p.add_argument('--portable-output',type=Path)
    p.add_argument('--skip-variable-instances',action='store_true',help='Inventory default variation coordinates only')
    p.add_argument('--reuse-catalog',type=Path,help='Audit against an existing inventory without reading installed fonts')
    p.add_argument('--require-font',action='append',default=[],help='Fail if this explicitly required font name is absent')
    args=p.parse_args()
    args.output.mkdir(parents=True,exist_ok=True)
    if args.reuse_catalog:
        catalog=json.loads(args.reuse_catalog.read_text(encoding='utf-8'))
        source=args.reuse_catalog.parent
        assert hashlib.sha256((source/'metrics.json.gz').read_bytes()).hexdigest()==catalog['metrics_sha256']
        for face in catalog['faces']:normalize_instance_style(face)
        (args.output/'catalog.json').write_text(json.dumps(catalog,ensure_ascii=False,indent=2),encoding='utf-8')
        if source.resolve()!=args.output.resolve():
            shutil.copyfile(source/'metrics.json.gz',args.output/'metrics.json.gz')
    else:
        catalog=export(args.font_root or default_roots(),args.output,not args.skip_variable_instances)
        source=args.output
    report=dict(faces=len(catalog['faces']),aliases=len(catalog['aliases']),errors=len(catalog['errors']),
                compressed_bytes=catalog['compressed_bytes'],variable_faces=sum(bool(f['variable_axes']) for f in catalog['faces']))
    if args.corpus:
        corpus=corpus_audit(args.corpus,catalog,args.output)
        report.update(documents=corpus['documents'],declared_names=corpus['names'],
                      unresolved_names=sum(r['status']=='unresolved_name' for r in corpus['rows']),corpus_errors=len(corpus['errors']))
    if args.portable_output:
        if catalog['errors']:raise RuntimeError('Refusing to pack a catalog with extraction errors')
        pack_portable(catalog,source,args.portable_output)
    supported=resolution_names(catalog)
    missing=[name for name in args.require_font if name.casefold() not in supported]
    report['missing_required_fonts']=missing
    (args.output/'summary.json').write_text(json.dumps(report,indent=2),encoding='utf-8')
    print(json.dumps(report),flush=True)
    if catalog['errors']:raise SystemExit(1)
    if missing:raise SystemExit(2)


if __name__=='__main__':main()
