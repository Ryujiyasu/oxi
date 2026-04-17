"""Measure Word's x-positions for the EXACT text of d77a PARA 21 L1.
If Oxi breaks at 41 chars and Word at 39, the cumulative widths MUST diverge somewhere.
"""
import os, sys, time, zipfile
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

CT = """<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"""
RELS = """<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"""

RPR = '<w:rFonts w:ascii="ＭＳ ゴシック" w:eastAsia="ＭＳ ゴシック" w:hAnsi="ＭＳ ゴシック"/><w:sz w:val="24"/>'

# The full PARA 21 text (approx 280 chars) — we want to see Word's actual line breaks
TEXT = "平成26年6月19日に決定した「公共ＷｅｂＷｅｂサイト運用ガイドライン（案1.0版）」は、各省庁から示された意見を踏まえ、その利用形態を禁止したり制限したりせず、一方で、「対象として利用の様態が明確ではない利用の委託業者等の意見を聴き、平成27年度に見直しを行う予定となっている見直しが実施された後のガイドラインの活用を十分検討して作業を行われたい。"

body = f'<w:p><w:pPr><w:ind w:firstLine="240"/><w:rPr>{RPR}</w:rPr></w:pPr><w:r><w:rPr>{RPR}</w:rPr><w:t xml:space="preserve">{TEXT}</w:t></w:r></w:p>'
DOC = f'<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="851" w:bottom="1134" w:left="851" w:header="851" w:footer="992"/><w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr></w:body></w:document>'
path = os.path.abspath("pipeline_data/d77a_para21_repro.docx")
with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", CT)
    z.writestr("_rels/.rels", RELS)
    z.writestr("word/document.xml", DOC)

w = win32com.client.Dispatch("Word.Application"); w.Visible = False
doc = w.Documents.Open(path, ReadOnly=True); time.sleep(0.5)

# Measure each character
chars = doc.Paragraphs(1).Range.Characters
n = min(chars.Count, 50)  # first 50 chars covers L1 and some of L2
print(f"Para has {chars.Count} chars. Measuring first 50 with advance from prev.")
print(f"{'#':>3} {'char':^4} {'codepoint':^8} {'x':>8} {'y':>7} {'advance':>8}")
prev_x = None
for c in range(1, n + 1):
    ch = chars(c)
    x = ch.Information(7)
    y = ch.Information(6)
    txt = ch.Text
    cp = f"U+{ord(txt):04X}" if len(txt) == 1 else " ".join(f"U+{ord(t):04X}" for t in txt)
    adv = "-" if prev_x is None else f"{x - prev_x:.2f}"
    print(f"{c:>3} '{txt}' {cp:^8} {x:>8.2f} {y:>7.2f} {adv:>8}")
    prev_x = x if y == (chars(1).Information(6) if c == 1 else y) else None

doc.Close(False); w.Quit()
