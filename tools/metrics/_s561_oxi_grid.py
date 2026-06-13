# -*- coding: utf-8 -*-
import os,sys,glob,subprocess,json,tempfile
sys.stdout.reconfigure(encoding='utf-8')
R=os.path.abspath('tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe')
base=os.path.abspath('tools/golden-test/repros/s561_bottom_grid')
print('variant       Oxi_pages  MARKER_page  MARKER_y')
for f in sorted(glob.glob(base+'/*.docx')):
    with tempfile.TemporaryDirectory() as td:
        dj=os.path.join(td,'l.json')
        subprocess.run([R,f,os.path.join(td,'p'),'150','--dump-layout='+dj],capture_output=True)
        d=json.load(open(dj,encoding='utf-8'))
        mk_pg=mk_y=None
        for pgno,p in enumerate(d['pages']):
            for e in p['elements']:
                if e['type']=='text' and e['text'].strip().startswith('MARKER'):
                    if mk_pg is None: mk_pg=pgno+1; mk_y=round(e['y'],1)
                # MARKER may be split per char; first char
            # check joined line too
            from collections import defaultdict
            lines=defaultdict(list)
            for e in p['elements']:
                if e['type']=='text' and e['text'].strip(): lines[round(e['y'])].append(e)
            for y,es in lines.items():
                t=''.join(x['text'] for x in sorted(es,key=lambda x:x['x']))
                if 'MARKER' in t and mk_pg is None:
                    mk_pg=pgno+1; mk_y=round(y,1)
        print('%-22s %d         p%s        %s'%(os.path.basename(f)[:-11],len(d['pages']),mk_pg,mk_y))
