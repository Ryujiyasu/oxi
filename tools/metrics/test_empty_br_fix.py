"""Test: does the page_break_after fix render Variant A correctly?

Word (expected):
  P1 page=1 y= 56.50
  P2 page=1 y= 70.50
  empty-br page=1 y= 84.00   ← stub line on p1
  P3 page=2 y= 56.50
  P4 page=2 y= 70.50
"""
import os, sys, subprocess, zipfile, json
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

TMP = os.path.abspath("pipeline_data/_empty_br_test.docx")
LAYOUT = os.path.abspath("pipeline_data/_empty_br_layout.json")

CT = '<?xml version="1.0"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
RELS = '<?xml version="1.0"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'

FONT_RPR = '<w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:sz w:val="21"/>'

def text_para(t, pbb=False):
    p = '<w:pageBreakBefore/>' if pbb else ''
    return f'<w:p><w:pPr>{p}<w:rPr>{FONT_RPR}</w:rPr></w:pPr><w:r><w:rPr>{FONT_RPR}</w:rPr><w:t>{t}</w:t></w:r></w:p>'

empty_br = f'<w:p><w:pPr><w:rPr>{FONT_RPR}</w:rPr></w:pPr><w:r><w:rPr>{FONT_RPR}</w:rPr><w:br w:type="page"/></w:r></w:p>'

body = "\n".join([
    text_para("Para1 on p1"),
    text_para("Para2 on p1"),
    empty_br,
    text_para("Para3 after break should be p2"),
    text_para("Para4 on p2"),
])
xml = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>\n{body}\n<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="851" w:bottom="1134" w:left="851" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr></w:body></w:document>'

with zipfile.ZipFile(TMP, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", CT)
    z.writestr("_rels/.rels", RELS)
    z.writestr("word/document.xml", xml)

# Run Oxi renderer with --dump-layout
renderer = os.path.abspath("tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe")
r = subprocess.run([renderer, TMP, "pipeline_data/_empty_br", "150",
                    f"--dump-layout={LAYOUT}", "--no-render"],
                   capture_output=True, text=True)
print("Renderer stderr:", r.stderr.strip())

with open(LAYOUT, encoding="utf-8") as f:
    d = json.load(f)

print()
print(f"Oxi produced {len(d['pages'])} pages")
for pg in d['pages']:
    ys = sorted({e['y'] for e in pg['elements'] if e.get('text','').strip() or 'text' not in e})
    # Show unique y values and first text at each
    seen = set()
    for e in pg['elements']:
        y = e['y']
        if y in seen: continue
        seen.add(y)
        t = e.get('text','')
        print(f"  p{pg['page']} y={y:.2f} text={t[:30]!r}")

print()
print("Expected (from Word):")
print("  p1 y=56.50 'Para1'")
print("  p1 y=70.50 'Para2'")
print("  p1 y=84.00 '' (empty-br stub)")
print("  p2 y=56.50 'Para3'")
print("  p2 y=70.50 'Para4'")
