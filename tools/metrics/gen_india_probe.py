# -*- coding: utf-8 -*-
"""A minimal Devanagari .docx that stresses Indic shaping, for a first read of
Oxi vs Word before real government documents arrive."""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn

LINES = [
    ("भारत सरकार", "plain word"),
    ("नमस्ते दुनिया", "greeting"),
    ("कि की कु कू के कैं कों", "matras / i-matra reordering"),
    ("क्ष त्र ज्ञ श्री द्ध क्त", "conjuncts (half-forms)"),
    ("दस्तावेज़ संख्या 2026", "nukta + Latin digits"),
    ("राष्ट्रीय सूचना विज्ञान केंद्र", "long compound"),
    ("Government of India — भारत सरकार", "mixed Latin + Devanagari"),
]

doc = Document()
doc.add_heading("India corpus probe — Devanagari shaping", level=1)
for text, note in LINES:
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.font.size = Pt(20)
    # Devanagari is a COMPLEX script: Word picks the font from w:cs, not
    # w:ascii. Set all three so both engines resolve Nirmala UI (Windows'
    # Indic UI font).
    rpr = r._element.get_or_add_rPr()
    rfonts = rpr.find(qn('w:rFonts'))
    if rfonts is None:
        rfonts = rpr.makeelement(qn('w:rFonts'), {})
        rpr.insert(0, rfonts)
    for attr in ('w:ascii', 'w:hAnsi', 'w:cs'):
        rfonts.set(qn(attr), 'Nirmala UI')
    # cs=1 marks the run as complex-script so the cs font/size apply.
    szcs = rpr.makeelement(qn('w:szCs'), {qn('w:val'): '40'})
    rpr.append(szcs)
    n = p.add_run("    (" + note + ")")
    n.font.size = Pt(9)

out = "pipeline_data/india_corpus/docx/_probe_devanagari.docx"
doc.save(out)
print("wrote", out)
