"""Demo decks for the Oxi Slides listing images (fictional content)."""
import sys
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

INK = RGBColor(0x1F, 0x29, 0x37)
MUTED = RGBColor(0x6B, 0x72, 0x80)
ACCENT = RGBColor(0x25, 0x63, 0xEB)
ACCENT2 = RGBColor(0x10, 0xB9, 0x81)
ACCENT3 = RGBColor(0xF5, 0x9E, 0x0B)
LIGHT = RGBColor(0xEF, 0xF4, 0xFF)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)


def deck():
    p = Presentation()
    p.slide_width, p.slide_height = Inches(13.333), Inches(7.5)
    return p


def text(slide, x, y, w, h, s, size, color=INK, bold=False, align=PP_ALIGN.LEFT, font=None):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    lines = s if isinstance(s, list) else [s]
    for i, line in enumerate(lines):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = align
        r = para.add_run()
        r.text = line
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.color.rgb = color
        if font:
            r.font.name = font
    return tb


def box(slide, shape, x, y, w, h, fill, line=None):
    s = slide.shapes.add_shape(shape, x, y, w, h)
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
    s.shadow.inherit = False
    return s


def header(slide, title, font):
    box(slide, MSO_SHAPE.RECTANGLE, 0, 0, Inches(0.25), Inches(7.5), ACCENT)
    text(slide, Inches(0.8), Inches(0.5), Inches(11.5), Inches(0.9), title, 32, INK, True, font=font)
    box(slide, MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.35), Inches(1.2), Inches(0.08), ACCENT)


def build(lang, out):
    ja = lang == 'ja'
    font = 'Noto Sans CJK JP' if ja else 'Calibri'
    p = deck()
    blank = p.slide_layouts[6]
    T = {
        'title': '新サービス ローンチ計画' if ja else 'Q4 Launch Plan',
        'sub': '2026年度 第4四半期 / 企画部' if ja else 'Product team · Fourth quarter 2026',
        'tag': 'エージェントが作成' if ja else 'Drafted by an agent',
        'goals': '今期のゴール' if ja else 'Goals for the quarter',
        'g': ([('新規ユーザー', '12,000', '前期比 +40%'), ('継続率', '87%', '30日後'), ('問い合わせ対応', '4時間', '初回返信まで')]
              if ja else [('New users', '12,000', '+40% on Q3'), ('Retention', '87%', 'after 30 days'), ('Support reply', '4 h', 'first response')]),
        'plan': 'スケジュール' if ja else 'Timeline',
        'steps': (['ベータ募集', '価格決定', '一般公開', '振り返り'] if ja else ['Beta sign-ups', 'Pricing', 'Launch', 'Review']),
        'months': (['10月', '11月', '12月', '1月'] if ja else ['Oct', 'Nov', 'Dec', 'Jan']),
        'budget': '予算と担当' if ja else 'Budget and owners',
        'rows': ([('項目', '担当', '予算', '状態'), ('広告', 'マーケティング', '¥3,200,000', '承認済み'),
                  ('サポート体制', 'CS', '¥1,100,000', '確認中'), ('インフラ', '開発', '¥2,400,000', '承認済み'),
                  ('イベント', '広報', '¥800,000', '未着手')]
                 if ja else [('Item', 'Owner', 'Budget', 'Status'), ('Ads', 'Marketing', '$32,000', 'Approved'),
                             ('Support staffing', 'Customer care', '$11,000', 'In review'), ('Infrastructure', 'Engineering', '$24,000', 'Approved'),
                             ('Launch event', 'Communications', '$8,000', 'Not started')]),
        'risks': 'リスクと対策' if ja else 'Risks and mitigations',
        'r': ([('需要の読み違い', 'ベータの登録数で毎週見直す'), ('サポートの逼迫', '公開前にFAQと自動返信を用意'), ('決済の障害', '予備の決済手段を用意')]
              if ja else [('Demand is lower than planned', 'Review beta sign-ups every week'),
                          ('Support is overloaded', 'Ship the FAQ and auto-replies before launch'),
                          ('Payments fail', 'Keep a second payment provider ready')]),
    }

    # 1 title
    s = p.slides.add_slide(blank)
    box(s, MSO_SHAPE.RECTANGLE, 0, 0, p.slide_width, p.slide_height, LIGHT)
    box(s, MSO_SHAPE.OVAL, Inches(8.6), Inches(-1.6), Inches(6.5), Inches(6.5), ACCENT)
    box(s, MSO_SHAPE.OVAL, Inches(10.8), Inches(3.9), Inches(3.6), Inches(3.6), ACCENT2)
    box(s, MSO_SHAPE.OVAL, Inches(7.9), Inches(4.9), Inches(1.4), Inches(1.4), ACCENT3)
    text(s, Inches(0.9), Inches(2.3), Inches(8), Inches(1.4), T['title'], 50, INK, True, font=font)
    text(s, Inches(0.9), Inches(3.6), Inches(8), Inches(0.6), T['sub'], 22, MUTED, font=font)
    pill = box(s, MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.9), Inches(4.6), Inches(3.0), Inches(0.5), WHITE, ACCENT)
    tf = pill.text_frame
    tf.paragraphs[0].alignment = PP_ALIGN.CENTER
    r = tf.paragraphs[0].add_run(); r.text = T['tag']; r.font.size = Pt(14); r.font.color.rgb = ACCENT; r.font.name = font

    # 2 goals
    s = p.slides.add_slide(blank)
    header(s, T['goals'], font)
    for i, (k, v, note) in enumerate(T['g']):
        x = Inches(0.8 + i * 4.1)
        c = [ACCENT, ACCENT2, ACCENT3][i]
        card = box(s, MSO_SHAPE.ROUNDED_RECTANGLE, x, Inches(2.0), Inches(3.7), Inches(3.8), LIGHT)
        box(s, MSO_SHAPE.RECTANGLE, x, Inches(2.0), Inches(3.7), Inches(0.12), c)
        text(s, x + Inches(0.3), Inches(2.5), Inches(3.2), Inches(0.6), k, 20, MUTED, font=font)
        text(s, x + Inches(0.3), Inches(3.2), Inches(3.2), Inches(1.2), v, 48, c, True, font=font)
        text(s, x + Inches(0.3), Inches(4.6), Inches(3.2), Inches(0.6), note, 18, INK, font=font)

    # 3 timeline
    s = p.slides.add_slide(blank)
    header(s, T['plan'], font)
    box(s, MSO_SHAPE.RECTANGLE, Inches(2.0), Inches(4.0), Inches(10.2), Inches(0.08), MUTED)
    for i, (step, m) in enumerate(zip(T['steps'], T['months'])):
        x = Inches(2.0 + i * 3.1)
        c = [ACCENT, ACCENT, ACCENT2, ACCENT3][i]
        box(s, MSO_SHAPE.OVAL, x - Inches(0.25), Inches(3.79), Inches(0.5), Inches(0.5), c)
        text(s, x - Inches(1.0), Inches(2.9), Inches(2.0), Inches(0.6), m, 20, MUTED, True, PP_ALIGN.CENTER, font)
        chev = box(s, MSO_SHAPE.ROUNDED_RECTANGLE, x - Inches(1.2), Inches(4.7), Inches(2.4), Inches(0.9), c)
        tf = chev.text_frame
        tf.paragraphs[0].alignment = PP_ALIGN.CENTER
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        r = tf.paragraphs[0].add_run(); r.text = step; r.font.size = Pt(16); r.font.bold = True; r.font.color.rgb = WHITE; r.font.name = font

    # 4 table
    s = p.slides.add_slide(blank)
    header(s, T['budget'], font)
    rows = T['rows']
    tbl = s.shapes.add_table(len(rows), 4, Inches(0.8), Inches(1.9), Inches(11.7), Inches(0.7 * len(rows))).table
    for ri, row in enumerate(rows):
        for ci, val in enumerate(row):
            cell = tbl.cell(ri, ci)
            cell.text = val
            para = cell.text_frame.paragraphs[0]
            para.runs[0].font.size = Pt(18)
            para.runs[0].font.name = font
            para.runs[0].font.bold = ri == 0
            para.runs[0].font.color.rgb = WHITE if ri == 0 else INK
            cell.fill.solid()
            cell.fill.fore_color.rgb = ACCENT if ri == 0 else (LIGHT if ri % 2 else WHITE)

    # 5 risks
    s = p.slides.add_slide(blank)
    header(s, T['risks'], font)
    for i, (risk, fix) in enumerate(T['r']):
        y = Inches(1.9 + i * 1.6)
        box(s, MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), y, Inches(5.2), Inches(1.2), LIGHT)
        text(s, Inches(1.1), y + Inches(0.3), Inches(4.8), Inches(0.7), risk, 20, INK, True, font=font)
        box(s, MSO_SHAPE.RIGHT_ARROW, Inches(6.3), y + Inches(0.35), Inches(0.9), Inches(0.5), ACCENT)
        box(s, MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.5), y, Inches(5.0), Inches(1.2), WHITE, ACCENT2)
        text(s, Inches(7.8), y + Inches(0.3), Inches(4.6), Inches(0.7), fix, 18, INK, font=font)

    p.save(out)


build(sys.argv[1], sys.argv[2])
