# -*- coding: utf-8 -*-
"""compatLnSpc text box: does a mixed line (UD デジタル 教科書体 NP-B 10 + Arial 14)
take the union box (largest ascent over largest descent) like a body line?
ja/administrative__0ba21b0c's frame heading «～個人懇談会～Reunião de pais e
mestres» (compatLnSpc=1, atLeast 0): Word's next line is 2.7pt lower than Oxi's
(Oxi sizes the CLNSP line by its largest run, Arial 14).
Arms: a text box (compatLnSpc 1 / 0) with A / X / B, X = Arial 14 alone or
after a UD 10 run, rule auto / atLeast 0. Readout: baselines."""
import os, sys, zipfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
from _pb_shape_cmpd_shadow_gen import NS, CT, RR, DRELS, STYLES, EMU
OUT = 'tests/fixtures/clnsp_mixed_union'; os.makedirs(OUT, exist_ok=True)
UD = 'UD デジタル 教科書体 NP-B'
def doc(lead, rule, clnsp):
    sp = '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>' if rule == 'auto' else '<w:spacing w:after="0" w:line="0" w:lineRule="atLeast"/>'
    r = lambda t, rf: f'<w:r><w:rPr>{rf}</w:rPr><w:t xml:space="preserve">{t}</w:t></w:r>'
    ud = f'<w:rFonts w:ascii="{UD}" w:eastAsia="{UD}" w:hAnsi="{UD}"/><w:sz w:val="20"/>'
    ar = f'<w:rFonts w:ascii="Arial" w:eastAsia="{UD}" w:hAnsi="Arial"/><w:sz w:val="28"/>'
    p = lambda b: f'<w:p><w:pPr>{sp}</w:pPr>{b}</w:p>'
    content = p(r('ああああ', ud)) + p((r('～個人～', ud) if lead else '') + r('Reunião de pais', ar)) + p(r('いいいい', ud))
    box = (f'<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="2" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>'
           f'<wp:positionH relativeFrom="page"><wp:posOffset>{100 * EMU}</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>{100 * EMU}</wp:posOffset></wp:positionV>'
           f'<wp:extent cx="{300 * EMU}" cy="{200 * EMU}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="2" name="B"/><wp:cNvGraphicFramePr/>'
           f'<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{300 * EMU}" cy="{200 * EMU}"/></a:xfrm>'
           f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr><wps:txbx><w:txbxContent>{content}</w:txbxContent></wps:txbx>'
           f'<wps:bodyPr wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t" compatLnSpc="{clnsp}"/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>')
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body><w:p>{box}</w:p>'
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1080" w:bottom="1440" w:left="1080" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')
if __name__ == '__main__':
    import win32com.client, fitz
    app = win32com.client.DispatchEx('Word.Application'); app.Visible = False; app.DisplayAlerts = 0
    try:
        for clnsp in (1, 0):
            for rule in ('auto', 'atl'):
                for lead in (False, True):
                    name = f'c{clnsp}_{rule}_{"lead" if lead else "alone"}'
                    at = os.path.abspath(os.path.join(OUT, name + '.docx'))
                    with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
                        z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR)
                        z.writestr('word/_rels/document.xml.rels', DRELS); z.writestr('word/styles.xml', STYLES); z.writestr('word/document.xml', doc(lead, rule, clnsp))
                    d = app.Documents.Open(at, ReadOnly=True)
                    try: d.ExportAsFixedFormat(at[:-5] + '.pdf', 17)
                    finally: d.Close(False)
                    ys = sorted({round(l['spans'][0]['origin'][1], 2) for b in fitz.open(at[:-5] + '.pdf')[0].get_text('dict')['blocks'] for l in b.get('lines', [])})
                    print(f'{name:18} gaps {[round(b - a, 2) for a, b in zip(ys, ys[1:])]}', flush=True)
    finally:
        app.Quit()
