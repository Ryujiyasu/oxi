# -*- coding: utf-8 -*-
"""Adversarial probe: OFFSET-positioned narrow floating table (tblpX, no
tblpXSpec align) — does Word wrap body text beside it, and at what
geometry? (The S758b v1 scope excluded offset floats pending this truth;
ed025c's OFF-column offset floats measured no-wrap, but the IN-COLUMN
offset case is unpinned.)

Two floats: 甲 at tblpX=3200tw (160pt, mid-column) and 乙 at tblpX=5400tw
(270pt, right-of-center), both vertAnchor=text, 170pt wide.

Run: python tools/metrics/_probe_tbloffset.py
Then: python tools/metrics/measure_pagination_word.py probeqtbloffset
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _probe_gen as pg
from _probe_gen3 import rpr, conds, pg2_borders, esc, out


def cellp(txt):
    r = rpr()
    return (f'<w:p><w:pPr><w:jc w:val="left"/><w:rPr>{r}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">{esc(txt)}</w:t></w:r></w:p>')


def ftbl(tag, tblp_x_tw):
    rows = "".join(
        f'<w:tr><w:tc><w:tcPr><w:tcW w:w="3400" w:type="dxa"/></w:tcPr>{cellp(f"{tag}項目{j+1}：数値{(j+1)*7}")}</w:tc></w:tr>'
        for j in range(8))
    return ('<w:tbl><w:tblPr>'
            '<w:tblpPr w:leftFromText="142" w:rightFromText="142" w:vertAnchor="text" '
            f'w:horzAnchor="margin" w:tblpX="{tblp_x_tw}" w:tblpY="1"/>'
            '<w:tblW w:w="3400" w:type="dxa"/>' + pg2_borders() + '</w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="3400"/></w:tblGrid>' + rows + '</w:tbl>')


def build():
    body = (conds(1, 5) + ftbl("甲", 3200) + conds(6, 28) + ftbl("乙", 5400)
            + conds(29, 48) + pg.sectpr())
    o = out("probeqtbloffset_offsetfloattable.docx")
    pg.write_docx(o, pg.doc(body))
    print("wrote", o)


if __name__ == "__main__":
    build()
