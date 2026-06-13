# -*- coding: utf-8 -*-
"""S559 — fingerprint the THREE cells that change line count under the gate:
  ⑦  (p79, 1->2, the CORRECT fix; Word reserves cellMar)
  p19 (労働基準法においては, 2->3, WRONG over-wrap; Word does NOT reserve)
  p94 (１（１）（２）..., 119->121, WRONG over-wrap; Word does NOT reserve)
Find each anchor's containing <w:tc> + <w:tbl>, dump tcW / gridCol / tblPr /
the paragraph's pPr (jc, ind, firstLine). The discriminator that includes ⑦ but
excludes p19/p94 is the gate condition that yields 55/55.
"""
import sys
import zipfile
import xml.etree.ElementTree as ET

DOCX = r'c:\tmp\3a4f9f.docx'
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
sys.stdout.reconfigure(encoding='utf-8')
ANCHORS = {
    'maru7': u'常に整理整頓',
    'p19': u'労働基準法においては、労働時間',
    'p94': u'安全衛生委員会',  # placeholder; refined below by search if missing
}


def q(t):
    return W + t


def attr(el, n):
    return el.get(W + n) if el is not None else None


def para_text(p):
    return ''.join(t.text or '' for t in p.iter(q('t')))


def dump_tbl(tbl):
    pr = tbl.find(q('tblPr'))
    def f(tag, a='val'):
        e = pr.find(q(tag)) if pr is not None else None
        return attr(e, a) if e is not None else None
    tblw_e = pr.find(q('tblW')) if pr is not None else None
    g = tbl.find(q('tblGrid'))
    gcols = [attr(c, 'w') for c in g.findall(q('gridCol'))] if g is not None else []
    cm = pr.find(q('tblCellMar')) if pr is not None else None
    cmv = None
    if cm is not None:
        cmv = {s: (attr(cm.find(q(s)), 'w')) for s in ('left', 'right') if cm.find(q(s)) is not None}
    print('    tblStyle=%s tblW=(%s,%s) layout=%s tblInd=%s nrows=%d ncols=%d gridcols=%s cellMar=%s'
          % (f('tblStyle'),
             attr(tblw_e, 'type'), attr(tblw_e, 'w'),
             f('tblLayout', 'type'), f('tblInd', 'w'),
             len(tbl.findall(q('tr'))), len(gcols), gcols, cmv))


def dump_cell(tc):
    pr = tc.find(q('tcPr'))
    tcw = pr.find(q('tcW')) if pr is not None else None
    gs = pr.find(q('gridSpan')) if pr is not None else None
    tcmar = pr.find(q('tcMar')) if pr is not None else None
    print('    tcW=(%s,%s) gridSpan=%s tcMar=%s npara=%d'
          % (attr(tcw, 'type'), attr(tcw, 'w'), attr(gs, 'val'),
             'yes' if tcmar is not None else 'no', len(tc.findall(q('p')))))


def dump_ppr(p):
    pr = p.find(q('pPr'))
    if pr is None:
        print('    pPr: NONE')
        return
    jc = pr.find(q('jc'))
    ind = pr.find(q('ind'))
    pstyle = pr.find(q('pStyle'))
    indinfo = {}
    if ind is not None:
        for a in ('left', 'leftChars', 'right', 'rightChars', 'firstLine',
                  'firstLineChars', 'hanging', 'hangingChars'):
            v = attr(ind, a)
            if v is not None:
                indinfo[a] = v
    print('    pStyle=%s jc=%s ind=%s' % (attr(pstyle, 'val'), attr(jc, 'val'), indinfo))


def main():
    with zipfile.ZipFile(DOCX) as z:
        root = ET.fromstring(z.read('word/document.xml'))

    # build parent map for tc/tbl ancestry
    parent = {c: p for p in root.iter() for c in p}

    def ancestors(el):
        cur = el
        chain = []
        while cur in parent:
            cur = parent[cur]
            chain.append(cur)
        return chain

    # locate target paras by anchor; for p94 fall back to the 119-line cell text
    targets = {}
    for p in root.iter(q('p')):
        txt = para_text(p)
        for name, anc in ANCHORS.items():
            if name not in targets and anc and anc in txt:
                targets[name] = p

    # p94 fallback: find the para whose text starts with '１（１）（２）'? It is a
    # marker concat; instead search a cell containing many short numbered paras.
    if 'p94' not in targets:
        # find a paragraph literally '（１）' repeated context near doc end
        cand = None
        for p in root.iter(q('p')):
            if para_text(p).strip() in (u'１', u'（１）', u'（２）'):
                cand = p
        if cand is not None:
            targets['p94'] = cand

    for name in ('maru7', 'p19', 'p94'):
        print('\n===== %s =====' % name)
        p = targets.get(name)
        if p is None:
            print('  NOT FOUND')
            continue
        print('  text=%r' % para_text(p)[:50])
        dump_ppr(p)
        chain = ancestors(p)
        tc = next((a for a in chain if a.tag == q('tc')), None)
        tbl = next((a for a in chain if a.tag == q('tbl')), None)
        if tc is not None:
            print('  CELL:')
            dump_cell(tc)
            # which row index, ncells in that row
            tr = next((a for a in chain if a.tag == q('tr')), None)
            if tr is not None:
                print('    row ncells=%d' % len(tr.findall(q('tc'))))
        else:
            print('  (not in a table cell)')
        if tbl is not None:
            print('  TABLE:')
            dump_tbl(tbl)


if __name__ == '__main__':
    main()
