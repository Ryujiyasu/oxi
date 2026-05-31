"""Derive Word's table cell row-height formula: vary border sz, cell margins,
font size, line rule; measure Word row pitch (col1 cell tops). Isolates the
source of the +0.375pt cell-row excess over the body line (15.0pt @ Cambria 11)."""
import zipfile, os
import win32com.client as win32

OUT=r"C:\Users\ryuji\oxi-main\tools\golden-test\repros\gen2_lineheight"
VPOS=6
CT='''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>'''
RELS='''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'''
DOC_RELS='''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'''
STYLES='''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Cambria" w:hAnsi="Cambria"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>'''

def write_docx(path, doc):
    with zipfile.ZipFile(path,"w",zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml",CT); z.writestr("_rels/.rels",RELS)
        z.writestr("word/_rels/document.xml.rels",DOC_RELS); z.writestr("word/styles.xml",STYLES)
        z.writestr("word/document.xml",doc)

def table_doc(rows=5, cols=3, sz=22, border=4, tcmar=None, line=276):
    def cell(r,c):
        mar=""
        if tcmar is not None:
            mar=(f'<w:tcMar><w:top w:w="{tcmar}" w:type="dxa"/><w:bottom w:w="{tcmar}" w:type="dxa"/></w:tcMar>')
        return ('<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/>'+mar+'</w:tcPr>'
                f'<w:p><w:pPr><w:spacing w:after="0" w:line="{line}" w:lineRule="auto"/>'
                f'<w:rPr><w:rFonts w:ascii="Cambria" w:hAnsi="Cambria"/><w:sz w:val="{sz}"/></w:rPr></w:pPr>'
                f'<w:r><w:rPr><w:rFonts w:ascii="Cambria" w:hAnsi="Cambria"/><w:sz w:val="{sz}"/></w:rPr>'
                f'<w:t>Cell {r}-{c}</w:t></w:r></w:p></w:tc>')
    trs=["<w:tr>"+"".join(cell(r,c) for c in range(1,cols+1))+"</w:tr>" for r in range(1,rows+1)]
    if border>0:
        b=("".join(f'<w:{e} w:val="single" w:sz="{border}" w:space="0" w:color="auto"/>'
           for e in ["top","left","bottom","right","insideH","insideV"]))
        borders=f'<w:tblBorders>{b}</w:tblBorders>'
    else:
        borders=('<w:tblBorders>'+ "".join(f'<w:{e} w:val="none" w:sz="0" w:space="0"/>'
                 for e in ["top","left","bottom","right","insideH","insideV"])+'</w:tblBorders>')
    tbl=f'<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>{borders}</w:tblPr>'+"".join(trs)+"</w:tbl>"
    mk=lambda t:(f'<w:p><w:pPr><w:spacing w:after="0" w:line="{line}" w:lineRule="auto"/>'
        f'<w:rPr><w:rFonts w:ascii="Cambria" w:hAnsi="Cambria"/><w:sz w:val="{sz}"/></w:rPr></w:pPr>'
        f'<w:r><w:rPr><w:rFonts w:ascii="Cambria" w:hAnsi="Cambria"/><w:sz w:val="{sz}"/></w:rPr><w:t>{t}</w:t></w:r></w:p>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
        +mk("BEFORE")+tbl+mk("AFTER")
        +'<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')

VARIANTS={
 "tbl_b4_sz22":      dict(sz=22,border=4),            # baseline (=earlier: 15.375)
 "tbl_b0_sz22":      dict(sz=22,border=0),            # no borders
 "tbl_b8_sz22":      dict(sz=22,border=8),            # 1pt borders
 "tbl_b4_tcmar0_sz22":dict(sz=22,border=4,tcmar=0),   # explicit 0 cell margins
 "tbl_b4_sz28":      dict(sz=28,border=4),            # 14pt (proportionality)
 "tbl_b0_sz28":      dict(sz=28,border=0),
}
for nm,kw in VARIANTS.items():
    p=os.path.join(OUT,nm+".docx"); write_docx(p,table_doc(**kw))

# measure Word row pitch for each
word=win32.gencache.EnsureDispatch("Word.Application");word.Visible=False
print(f"{'variant':<22} {'rowpitches':<34} {'mean':>7} {'bodyline':>9}")
print("-"*78)
for nm,kw in VARIANTS.items():
    docx=os.path.join(OUT,nm+".docx")
    doc=word.Documents.Open(docx,ReadOnly=True)
    try:
        tbl=doc.Tables(1); ys=[]
        for r in range(1,tbl.Rows.Count+1):
            cell=tbl.Cell(r,1).Range; st=doc.Range(cell.Start,cell.Start)
            ys.append(st.Information(VPOS))
        # body line: BEFORE para to first row? use BEFORE->AFTER block / rows for context
        pitches=[round(ys[i+1]-ys[i],3) for i in range(len(ys)-1)]
        mean=round(sum(pitches)/len(pitches),3) if pitches else 0
    finally:
        doc.Close(False)
    sz=kw.get("sz",22)
    print(f"{nm:<22} {str(pitches):<34} {mean:>7.3f} {'sz='+str(sz):>9}")
word.Quit()
