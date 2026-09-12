import importlib.util, os, sys, json, subprocess, tempfile
from pathlib import Path
spec=importlib.util.spec_from_file_location("lp","tools/metrics/_pb_line_pitch.py"); lp=importlib.util.module_from_spec(spec); spec.loader.exec_module(lp)
GDI=Path(os.environ.get("OXI_GDI_EXE") or "tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe")
OUT=Path('tests/fixtures/cell_docdefaults_after'); OUT.mkdir(parents=True,exist_ok=True)
def styles(after, tblstyle_after):
    ts=''
    if tblstyle_after is not None:
        ts=('<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/>'
            f'<w:pPr><w:spacing w:after="{tblstyle_after}" w:line="240" w:lineRule="auto"/></w:pPr>'
            '<w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr></w:style>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Cambria" w:eastAsia="ＭＳ 明朝" w:hAnsi="Cambria" w:cs="Times New Roman"/>'
            '<w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>'
            f'<w:pPrDefault><w:pPr><w:spacing w:after="{after}" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
            '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>'
            + ts + '</w:styles>')
def doc(cells, grid, tcw, tblstyle):
    tp=('<w:tblPr>'+(f'<w:tblStyle w:val="{tblstyle}"/>' if tblstyle else '')+'<w:tblW w:type="auto" w:w="0"/><w:jc w:val="center"/></w:tblPr>')
    tg='<w:tblGrid>'+''.join(f'<w:gridCol w:w="{grid}"/>' for _ in range(3))+'</w:tblGrid>'
    def row(texts):
        return '<w:tr>'+''.join(f'<w:tc><w:tcPr><w:tcW w:type="dxa" w:w="{tcw}"/></w:tcPr><w:p><w:r><w:t>{t}</w:t></w:r></w:p></w:tc>' for t in texts)+'</w:tr>'
    tbl='<w:tbl>'+tp+tg+''.join(row(r) for r in cells)+'</w:tbl>'
    p=lambda t:f'<w:p><w:r><w:t>{t}</w:t></w:r></w:p>'
    sect=('<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="720" w:footer="720" w:gutter="0"/><w:docGrid w:linePitch="360"/></w:sectPr>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
            +p('HEAD')+tbl+p('TAIL')+sect+'</w:body></w:document>')
LAT=[["-","north:38.85464","-"],["west:100.3722","-","east:100.3722"],["-","south:38.85464","-"]]
CJK=[["-","north：38.85464","-"],["west：100.3722","-","east：100.3722"],["-","south：38.85464","-"]]
MIX=[["-","north：38.85464","文本"],["west：100.3722","-","east：100.3722"],["-","south：38.85464","-"]]
arms=[("A_lat_nostyle_g1440_w2880",LAT,1440,2880,200,None,None),
      ("B_cjk_nostyle_g1440_w2880",CJK,1440,2880,200,None,None),
      ("C_mix_nostyle_g1440_w2880",MIX,1440,2880,200,None,None),
      ("D_cjk_grid_g1440_w1440",CJK,1440,1440,200,None,None),
      ("E_cjk_tblgrid_after0",CJK,1440,2880,200,"TableGrid",0),
      ("F_cjk_nostyle_after0",CJK,1440,2880,0,None,None),
      ("G_lat_grid_g1440_w1440",LAT,1440,1440,200,None,None)]
import pymupdf, win32com.client
app=win32com.client.DispatchEx("Word.Application")
try: app.Visible=False
except Exception: pass
res={}
try:
  for name,cells,grid,tcw,after,ts,tsa in arms:
    at=OUT/f"{name}.docx"
    lp.write(at, doc(cells,grid,tcw,ts), styles(after,tsa))
    pdf=str(at)[:-5]+'.pdf'
    d=app.Documents.Open(str(at.resolve()),False,True)
    try:
        d.ExportAsFixedFormat(OutputFileName=str(Path(pdf).resolve()),ExportFormat=17)
        cw=[round(d.Tables(1).Cell(1,c).Width,2) for c in (1,2,3)]
    finally: d.Close(False)
    pg=pymupdf.open(pdf)[0]
    lines=sorted((round(l['bbox'][1],2), round(l['bbox'][0],1), ''.join(s['text'] for s in l['spans'])) for b in pg.get_text('dict')['blocks'] for l in b.get('lines',[]))
    ys=sorted({y for y,_,_ in lines})
    e=dict(os.environ); e['OXI_S1363']='1'
    with tempfile.TemporaryDirectory() as t:
        dump=Path(t)/'l.json'; subprocess.run([str(GDI),str(at),str(Path(t)/'p'),"96",f"--dump-layout={dump}"],capture_output=True,env=e)
        dd=json.loads(dump.read_text(encoding='utf-8'))
    rows={}
    for el in dd['pages'][0]['elements']:
        if el.get('type')=='text' and el.get('text','').strip(): rows.setdefault(round(el['y'],2),[]).append((round(el['x'],1),el['text']))
    print(f"=== {name}  Word cell widths {cw}")
    print("  WORD y:", [f"{y}:{''.join(t for yy,_,t in lines if yy==y)[:28]}" for y in ys])
    print("  OXI  y:", [f"{y}:{''.join(t for _,t in sorted(rows[y]))[:28]}" for y in sorted(rows)])
finally: app.Quit()
