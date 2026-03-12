#!/usr/bin/env python3
"""Generate test .xlsx and .pptx fixture files for Oxi."""
import zipfile
import os

FIXTURES = os.path.join(os.path.dirname(__file__), "fixtures")
os.makedirs(FIXTURES, exist_ok=True)

def write_zip(path, files):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, content in files.items():
            zf.writestr(name, content)
    print(f"Generated {path}")

# === basic_test.xlsx ===
write_zip(os.path.join(FIXTURES, "basic_test.xlsx"), {
    "[Content_Types].xml": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>''',
    "_rels/.rels": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>''',
    "xl/workbook.xml": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sales" sheetId="1" r:id="rId1"/></sheets>
</workbook>''',
    "xl/_rels/workbook.xml.rels": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>''',
    "xl/sharedStrings.xml": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="7" uniqueCount="7">
  <si><t>Product</t></si><si><t>Q1</t></si><si><t>Q2</t></si><si><t>Q3</t></si><si><t>Q4</t></si>
  <si><t>Widget A</t></si><si><t>Widget B</t></si>
</sst>''',
    "xl/worksheets/sheet1.xml": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c>
      <c r="C1" t="s"><v>2</v></c><c r="D1" t="s"><v>3</v></c><c r="E1" t="s"><v>4</v></c>
    </row>
    <row r="2">
      <c r="A2" t="s"><v>5</v></c><c r="B2"><v>1200</v></c>
      <c r="C2"><v>1500.5</v></c><c r="D2"><v>980</v></c><c r="E2"><v>2100</v></c>
    </row>
    <row r="3">
      <c r="A3" t="s"><v>6</v></c><c r="B3"><v>800</v></c>
      <c r="C3"><v>920</v></c><c r="D3"><v>1100</v></c><c r="E3"><v>1350</v></c>
    </row>
    <row r="5">
      <c r="A5" t="b"><v>1</v></c><c r="B5" t="e"><v>#N/A</v></c>
      <c r="C5" t="str"><v>inline text</v></c>
    </row>
  </sheetData>
</worksheet>''',
})

# === multi_sheet.xlsx ===
write_zip(os.path.join(FIXTURES, "multi_sheet.xlsx"), {
    "[Content_Types].xml": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>''',
    "_rels/.rels": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>''',
    "xl/workbook.xml": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Data" sheetId="1" r:id="rId1"/>
    <sheet name="Summary" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>''',
    "xl/_rels/workbook.xml.rels": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>''',
    "xl/sharedStrings.xml": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="5" uniqueCount="5">
  <si><t>Name</t></si><si><t>Score</t></si><si><t>Alice</t></si><si><t>Bob</t></si><si><t>Total</t></si>
</sst>''',
    "xl/worksheets/sheet1.xml": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>
    <row r="2"><c r="A2" t="s"><v>2</v></c><c r="B2"><v>95</v></c></row>
    <row r="3"><c r="A3" t="s"><v>3</v></c><c r="B3"><v>87</v></c></row>
  </sheetData>
</worksheet>''',
    "xl/worksheets/sheet2.xml": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="s"><v>4</v></c><c r="B1"><v>182</v></c></row>
  </sheetData>
</worksheet>''',
})

# === basic_test.pptx ===
write_zip(os.path.join(FIXTURES, "basic_test.pptx"), {
    "[Content_Types].xml": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
</Types>''',
    "_rels/.rels": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>''',
    "ppt/presentation.xml": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst>
  <p:sldSz cx="9144000" cy="6858000"/>
</p:presentation>''',
    "ppt/_rels/presentation.xml.rels": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>''',
    "ppt/slides/slide1.xml": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="457200" y="274638"/><a:ext cx="8229600" cy="1143000"/></a:xfrm></p:spPr>
      <p:txBody><a:bodyPr/><a:p><a:pPr algn="ctr"/>
        <a:r><a:rPr lang="en-US" sz="4400" b="1"/><a:t>Welcome to Oxi</a:t></a:r>
      </a:p></p:txBody>
    </p:sp>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="3" name="Body"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="457200" y="1600200"/><a:ext cx="8229600" cy="4525963"/></a:xfrm></p:spPr>
      <p:txBody><a:bodyPr/>
        <a:p><a:r><a:rPr lang="en-US" sz="2400"/><a:t>Open-source document processing</a:t></a:r></a:p>
        <a:p><a:r><a:rPr lang="en-US" sz="2000" i="1"><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></a:rPr><a:t>Rust + WebAssembly</a:t></a:r></a:p>
        <a:p><a:r><a:rPr lang="ja-JP" sz="2000"/><a:t>\u65e5\u672c\u8a9e\u30c6\u30ad\u30b9\u30c8\u306e\u30c6\u30b9\u30c8</a:t></a:r></a:p>
      </p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
</p:sld>''',
})

# === multi_slide.pptx ===
write_zip(os.path.join(FIXTURES, "multi_slide.pptx"), {
    "[Content_Types].xml": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  <Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  <Override PartName="/ppt/slides/slide3.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
</Types>''',
    "_rels/.rels": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>''',
    "ppt/presentation.xml": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId2"/>
    <p:sldId id="257" r:id="rId3"/>
    <p:sldId id="258" r:id="rId4"/>
  </p:sldIdLst>
  <p:sldSz cx="9144000" cy="6858000"/>
</p:presentation>''',
    "ppt/_rels/presentation.xml.rels": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide3.xml"/>
</Relationships>''',
    "ppt/slides/slide1.xml": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="685800" y="2286000"/><a:ext cx="7772400" cy="1470025"/></a:xfrm></p:spPr>
      <p:txBody><a:bodyPr/><a:p><a:pPr algn="ctr"/>
        <a:r><a:rPr sz="5400" b="1"/><a:t>Oxi Project</a:t></a:r>
      </a:p></p:txBody>
    </p:sp>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="3" name="Subtitle"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="1371600" y="3886200"/><a:ext cx="6400800" cy="1752600"/></a:xfrm></p:spPr>
      <p:txBody><a:bodyPr/><a:p><a:pPr algn="ctr"/>
        <a:r><a:rPr sz="2400" i="1"><a:solidFill><a:srgbClr val="666666"/></a:solidFill></a:rPr><a:t>Document processing for the browser</a:t></a:r>
      </a:p></p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
</p:sld>''',
    "ppt/slides/slide2.xml": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="457200" y="274638"/><a:ext cx="8229600" cy="1143000"/></a:xfrm></p:spPr>
      <p:txBody><a:bodyPr/><a:p>
        <a:r><a:rPr sz="3600" b="1"/><a:t>Supported Formats</a:t></a:r>
      </a:p></p:txBody>
    </p:sp>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="3" name="Content"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="457200" y="1600200"/><a:ext cx="8229600" cy="4525963"/></a:xfrm></p:spPr>
      <p:txBody><a:bodyPr/>
        <a:p><a:r><a:rPr sz="2400"/><a:t>.docx \u2014 Word documents</a:t></a:r></a:p>
        <a:p><a:r><a:rPr sz="2400"/><a:t>.xlsx \u2014 Excel spreadsheets</a:t></a:r></a:p>
        <a:p><a:r><a:rPr sz="2400"/><a:t>.pptx \u2014 PowerPoint presentations</a:t></a:r></a:p>
        <a:p><a:r><a:rPr sz="2000" i="1"><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></a:rPr><a:t>All processed in-browser with Rust + Wasm</a:t></a:r></a:p>
      </p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
</p:sld>''',
    "ppt/slides/slide3.xml": u'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="457200" y="274638"/><a:ext cx="8229600" cy="1143000"/></a:xfrm></p:spPr>
      <p:txBody><a:bodyPr/><a:p><a:pPr algn="ctr"/>
        <a:r><a:rPr sz="3600" b="1"><a:latin typeface="Yu Gothic"/><a:ea typeface="Yu Gothic"/></a:rPr><a:t>\u65e5\u672c\u8a9e\u7d44\u7248\u30c6\u30b9\u30c8</a:t></a:r>
      </a:p></p:txBody>
    </p:sp>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="3" name="Body"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="457200" y="1600200"/><a:ext cx="8229600" cy="4525963"/></a:xfrm></p:spPr>
      <p:txBody><a:bodyPr/>
        <a:p><a:r><a:rPr sz="2000"><a:latin typeface="Yu Mincho"/><a:ea typeface="Yu Mincho"/></a:rPr><a:t>\u543e\u8f29\u306f\u732b\u3067\u3042\u308b\u3002\u540d\u524d\u306f\u307e\u3060\u7121\u3044\u3002</a:t></a:r></a:p>
        <a:p><a:r><a:rPr sz="2000"><a:ea typeface="Yu Gothic"/></a:rPr><a:t>\u3069\u3053\u3067\u751f\u308c\u305f\u304b\u3068\u3093\u3068\u898b\u5f53\u304c\u3064\u304b\u306c\u3002</a:t></a:r></a:p>
        <a:p><a:r><a:rPr sz="1800" i="1"><a:solidFill><a:srgbClr val="C00000"/></a:solidFill></a:rPr><a:t>\u7981\u5247\u51e6\u7406\u30fb\u30eb\u30d3\u30fb\u7e26\u66f8\u304d\u306b\u5bfe\u5fdc\u4e88\u5b9a</a:t></a:r></a:p>
      </p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
</p:sld>''',
})

print("Done!")
