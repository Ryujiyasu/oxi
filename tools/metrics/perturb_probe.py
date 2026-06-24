# -*- coding: utf-8 -*-
"""perturbation / ablation harness (Ra, 2026-06-24).

User's idea (2026-06-24): take a clean base that Oxi already renders well, inject
ONE "weird" feature at a time, and observe how the rendering changes — isolating
each feature's effect. Refinement learned in the S652/S653 math pass: whole-page
SSIM is too COARSE to see a single small perturbation (a few-pt shift moves SSIM
below its noise floor), so this measures the **downstream shift** via 300dpi
pixel ink-bands instead.

Method: build  [TOP][injection][BOTTOM]  with the injection between two body
markers. Measure the TOP->BOTTOM pixel distance in Word and Oxi; the signed delta
(Oxi - Word) = how wrong Oxi's reserved vertical space for that feature is. The
injection sits between markers so the delta is its OWN reservation error, not a
page-cascade artifact. Ranked by |delta|. cp932-safe.

Usage: python tools/metrics/perturb_probe.py [grid|nogrid|both]
"""
import os, sys, io, subprocess
import numpy as np
from PIL import Image

sys.path.insert(0, 'tools/metrics')
from mixedh_lineplace import build_generic, EXE, _ink_bands, _word_render_png
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='backslashreplace')

DPI = 300
FONT = 'ＭＳ 明朝'


def _rpr(sz=22, extra=''):
    return ('<w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s"/><w:sz w:val="%d"/>%s'
            % (FONT, FONT, FONT, sz, extra))


def marker(text):
    return '<w:p><w:pPr><w:rPr>%s</w:rPr></w:pPr><w:r><w:rPr>%s</w:rPr><w:t>%s</w:t></w:r></w:p>' % (
        _rpr(), _rpr(), text)


def run(text, sz=22, extra=''):
    return '<w:r><w:rPr>%s</w:rPr><w:t xml:space="preserve">%s</w:t></w:r>' % (_rpr(sz, extra), text)


# Each injection is the paragraph(s) placed between TOP and BOTTOM.
INJECTIONS = {
    'baseline':      '<w:p><w:pPr><w:rPr>%s</w:rPr></w:pPr>%s</w:p>' % (_rpr(), run('ふつうの段落')),
    'bigfont_inline': '<w:p><w:pPr><w:rPr>%s</w:rPr></w:pPr>%s%s%s</w:p>' % (
                      _rpr(), run('前 '), run('大', 96), run(' 後')),       # 48pt char in an 11pt line
    'superscript':   '<w:p><w:pPr><w:rPr>%s</w:rPr></w:pPr>%s%s</w:p>' % (
                      _rpr(), run('x'), run('2', 22, '<w:vertAlign w:val="superscript"/>')),
    'position_up':   '<w:p><w:pPr><w:rPr>%s</w:rPr></w:pPr>%s%s</w:p>' % (
                      _rpr(), run('A'), run('B', 22, '<w:position w:val="12"/>')),
    'position_dn22': '<w:p><w:pPr><w:rPr>%s</w:rPr></w:pPr>%s%s</w:p>' % (
                      _rpr(), run('A'), run('B', 22, '<w:position w:val="-22"/>')),  # corpus value (3a4f/model)  # raised 6pt
    'charspacing':   '<w:p><w:pPr><w:rPr>%s</w:rPr></w:pPr>%s</w:p>' % (
                      _rpr(), run('文字間隔', 22, '<w:spacing w:val="60"/>')),       # +3pt tracking
    'wscale50':      '<w:p><w:pPr><w:rPr>%s</w:rPr></w:pPr>%s</w:p>' % (
                      _rpr(), run('横圧縮文字', 22, '<w:w w:val="50"/>')),           # 50% width
    'exact30':       '<w:p><w:pPr><w:spacing w:line="600" w:lineRule="exact"/><w:rPr>%s</w:rPr></w:pPr>%s</w:p>' % (
                      _rpr(), run('行間固定30pt')),
    'ruby':          ('<w:p><w:pPr><w:rPr>%s</w:rPr></w:pPr><w:r><w:rPr>%s</w:rPr><w:ruby>'
                      '<w:rubyPr><w:rubyAlign w:val="distributeSpace"/><w:hps w:val="11"/>'
                      '<w:hpsRaise w:val="22"/><w:hpsBaseText w:val="22"/><w:lid w:val="ja-JP"/></w:rubyPr>'
                      '<w:rt><w:r><w:rPr>%s</w:rPr><w:t>かんじ</w:t></w:r></w:rt>'
                      '<w:rubyBase><w:r><w:rPr>%s</w:rPr><w:t>漢字</w:t></w:r></w:rubyBase>'
                      '</w:ruby></w:r></w:p>' % (_rpr(), _rpr(), _rpr(11), _rpr())),
    'dropcap':       ('<w:p><w:pPr><w:framePr w:dropCap="drop" w:lines="3" w:wrap="around"'
                      ' w:vAnchor="text" w:hAnchor="text"/><w:rPr>%s</w:rPr></w:pPr>%s</w:p>'
                      '<w:p><w:pPr><w:rPr>%s</w:rPr></w:pPr>%s</w:p>'
                      % (_rpr(66), run('D', 66), _rpr(), run('ropCap本文がまわり込む段落です。'))),
    'tab_right':     ('<w:p><w:pPr><w:tabs><w:tab w:val="right" w:pos="9000"/></w:tabs>'
                      '<w:rPr>%s</w:rPr></w:pPr>%s<w:r><w:rPr>%s</w:rPr><w:tab/></w:r>%s</w:p>'
                      % (_rpr(), run('左'), _rpr(), run('右'))),
    # batch 2 (2026-06-24): height/placement features likely to exercise line growth
    'em_dot':        '<w:p><w:pPr><w:rPr>%s</w:rPr></w:pPr>%s</w:p>' % (
                      _rpr(), run('圏点', 22, '<w:em w:val="dot"/>')),                 # 圏点 (sesame dots above)
    'atleast30':     '<w:p><w:pPr><w:spacing w:line="600" w:lineRule="atLeast"/><w:rPr>%s</w:rPr></w:pPr>%s</w:p>' % (
                      _rpr(), run('atLeast30pt')),
    'before12grid':  '<w:p><w:pPr><w:spacing w:before="240"/><w:rPr>%s</w:rPr></w:pPr>%s</w:p>' % (
                      _rpr(), run('前空き12pt')),                                       # spacing before 12pt
    'run_border':    '<w:p><w:pPr><w:rPr>%s</w:rPr></w:pPr>%s</w:p>' % (
                      _rpr(), run('囲み罫', 22, '<w:bdr w:val="single" w:sz="4" w:space="0" w:color="auto"/>')),
    'updn_mixed':    '<w:p><w:pPr><w:rPr>%s</w:rPr></w:pPr>%s%s%s</w:p>' % (
                      _rpr(), run('上', 22, '<w:position w:val="12"/>'), run('中'),
                      run('下', 22, '<w:position w:val="-12"/>')),                      # +6 and −6 on same line
    'combine':       '<w:p><w:pPr><w:rPr>%s</w:rPr></w:pPr><w:r><w:rPr>%s<w:eastAsianLayout w:combine="1"/></w:rPr><w:t>令和</w:t></w:r></w:p>' % (
                      _rpr(), _rpr()),                                                  # 割注/combine (2 chars in 1 cell)
    # batch 3 (2026-06-24): border/shading/spacing/break features
    'para_border':   ('<w:p><w:pPr><w:pBdr><w:top w:val="single" w:sz="8" w:space="4" w:color="auto"/>'
                      '<w:bottom w:val="single" w:sz="8" w:space="4" w:color="auto"/></w:pBdr>'
                      '<w:rPr>%s</w:rPr></w:pPr>%s</w:p>' % (_rpr(), run('段落罫線'))),
    'after_auto':    '<w:p><w:pPr><w:spacing w:afterAutospacing="1"/><w:rPr>%s</w:rPr></w:pPr>%s</w:p>' % (
                      _rpr(), run('後ろ自動空き')),
    'contextual':    '<w:p><w:pPr><w:contextualSpacing/><w:spacing w:after="240"/><w:rPr>%s</w:rPr></w:pPr>%s</w:p>' % (
                      _rpr(), run('文脈空き')),
    'soft_break':    '<w:p><w:pPr><w:rPr>%s</w:rPr></w:pPr>%s<w:r><w:rPr>%s</w:rPr><w:br/></w:r>%s</w:p>' % (
                      _rpr(), run('一行目'), _rpr(), run('二行目')),                     # w:br soft line break
    'large96':       '<w:p><w:pPr><w:rPr>%s</w:rPr></w:pPr>%s</w:p>' % (
                      _rpr(192), run('大', 192)),                                       # single 96pt char (extreme)
    'shd_para':      ('<w:p><w:pPr><w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/>'
                      '<w:rPr>%s</w:rPr></w:pPr>%s</w:p>' % (_rpr(), run('網掛け段落'))),
}


def t2b(png):
    b = _ink_bands(png, DPI)
    return round(b[-1][0] - b[0][0], 2) if len(b) >= 3 else None


def measure(grid_pitch):
    g = 'grid%d' % grid_pitch if grid_pitch else 'nogrid'
    rows = []
    for name, inj in INJECTIONS.items():
        body = marker('TOP') + inj + marker('BOTTOM')
        dx = build_generic('pb_%s_%s.docx' % (name, g), body, grid_pitch)
        op = os.path.join('c:/tmp/mixedh', 'pb_%s_%s' % (name, g))
        subprocess.run([EXE, os.path.abspath(dx), op, str(DPI)], capture_output=True, text=True)
        o = t2b(op + '_p1.png')
        w = t2b(_word_render_png(dx, DPI))
        d = round(o - w, 2) if (o is not None and w is not None) else None
        rows.append((name, w, o, d))
    print('=== PERTURBATION (%s): TOP->BOTTOM reserved space, Word vs Oxi ===' % g)
    print('%-16s | %-7s %-7s %-8s' % ('feature', 'Word', 'Oxi', 'd(O-W)'))
    # ranked by |delta|, None (overlap/failure) first
    def key(r):
        return (-1e9, r[0]) if r[3] is None else (-abs(r[3]), r[0])
    for name, w, o, d in sorted(rows, key=key):
        flag = '  <<< OVERLAP/FAIL' if d is None else ('  <<<' if abs(d) >= 3 else '')
        print('%-16s | %-7s %-7s %-8s%s' % (name, w, o, ('%+.2f' % d) if d is not None else 'None', flag))


if __name__ == '__main__':
    mode = sys.argv[1] if len(sys.argv) > 1 else 'both'
    if mode in ('grid', 'both'):
        measure(360)
    if mode in ('nogrid', 'both'):
        measure(None)
