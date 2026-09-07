# -*- coding: utf-8 -*-
"""Break-time capacity of the marks on a character-grid line (faithful slice of
reference__0ea3ec86: its own styles/settings/fonts, one section copied from the
document -- 4 = 2 columns charSpace 2048 (cell 11.51, 20 cells), 18 = 2 columns
charSpace 3194 (11.76), 5 = 1 column 2048 (42 cells)).

    python _pb_unitcap_gen.py gen [sect]      # unitcap[sect].docx
    python _pb_unitcap_gen.py pdf [sect]      # Word COM -> unitcap[sect].pdf
    python _pb_unitcap_read.py <label-prefix...> [--sect=N] [--full] [--two]
    OXI_S1318=1 python _pb_agree.py unitcap.docx unitcap.pdf [--cols=1]

Arm families (S1345/S1346, 2026-09-07): m/b/e/f/k/g = marks x demand with the
unit 字、; z_/y_/w_ = a standalone mark's 0.5 and where the half-width digit
sits; v_/t_ = the digit-adjacent refusal, run boundaries, the tracked cap;
d_ = the middle dot at the line end; u_ = w:w character scale; L_/M_/N_/O_/P_/
Q_ = the （+word unit and its budget; I_ = indents and the floor; Z_ = ZWJ and
line-start-prohibited units; R_/S_ = X before a hanging mark (2- and 1-column).
Every arm is a paragraph of the slice; the reader walks the PDF in document
order so rotated-kana prefixes cannot collide.
"""
import os
import re
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
SRC = os.path.join(REPO, "pipeline_data", "docx_corpus", "ja", "reference", "0ea3ec86480140c2.docx")
OUT = os.path.join(REPO, "pipeline_data", "_pb_unitcap")
SECT_IDX = int(sys.argv[2]) if len(sys.argv) > 2 and sys.argv[2].isdigit() else 4
CHARSPACE = next((int(a.split("=", 1)[1]) for a in sys.argv if a.startswith("--charspace=")), None)
DOCNAME = ("unitcap" if SECT_IDX == 4 else "unitcap%d" % SECT_IDX) + ("" if CHARSPACE is None else "_cs%d" % CHARSPACE)
ONLY = next((a.split("=")[1] for a in sys.argv if a.startswith("--only=")), None)   # arm label prefix filter
sys.stdout.reconfigure(encoding="utf-8")

KANA = "あいうえおかきくけこさしすせそたちつてとなにぬねのはひふへほまみむめもやゆよらりるれろわ"
FILL = "以下は次の行に流れる本文であって行末の判定には関わらない文字列を続ける。"


def line20(marks, kind="、"):
    """20 characters: kana with `marks` marks at spread positions (never first/last)."""
    chars = list(KANA[:20])
    pos = {1: [10], 2: [6, 13], 3: [5, 10, 15], 4: [4, 8, 12, 16]}.get(marks, [])
    for i, p in enumerate(pos):
        chars[p] = kind[i % len(kind)]
    return "".join(chars)


KANA = "あいうえおかきくけこさしすせそたちつてとなにぬねのはひふへほまみむめもやゆよらりるれろわ"

_ROT = [0]

def kana(n):
    _ROT[0] += 3
    r = _ROT[0] % len(KANA)
    return ((KANA[r:] + KANA[:r]) * 3)[:n]

# a 2-column 2048 line (floor 20 cells): 19 - m kana with m marks spread through the
# line, h half-width digits (0.5 cell each), then the unit 字、 (1 + 0.96 cells).
# demand over the floor = 0.96 + 0.5 h cells; capacity offered = m marks.
ARMS = []
for m in (1, 2, 3, 4):
    for h in (0, 1, 2, 3):
        body = list(kana(19 - m))
        step = len(body) // (m + 1)
        for k in range(m):
            body.insert((k + 1) * step + k, "、")
        text = "".join(body) + "1" * h + "字、"
        ARMS.append(("m%d_h%d" % (m, h), 0, text))
# the same with brackets as the marks: （x） pairs = 2 marks each
for pairs in (1, 2):
    for h in (0, 1, 2, 3):
        body = list(kana(19 - 2 * pairs))
        step = len(body) // (pairs + 1)
        for k in range(pairs):
            pos = (k + 1) * step + 2 * k
            body.insert(pos, "（")
            body.insert(pos + 2, "）")
        text = "".join(body) + "1" * h + "字、"
        ARMS.append(("b%d_h%d" % (pairs, h), 0, text))
# the same m-arms with w:hint="eastAsia" on the run (the real document's runs carry it)
for m in (1, 2, 3, 4):
    for h in (1, 2):
        body = list(kana(19 - m))
        step = len(body) // (m + 1)
        for k in range(m):
            body.insert((k + 1) * step + k, "、")
        text = "".join(body) + "1" * h + "字、"
        ARMS.append(("e%d_h%d" % (m, h), "hint", text))
for pairs in (1, 2):
    body = list(kana(19 - 2 * pairs))
    step = len(body) // (pairs + 1)
    for k in range(pairs):
        pos = (k + 1) * step + 2 * k
        body.insert(pos, "（")
        body.insert(pos + 2, "）")
    ARMS.append(("f%d_h1" % pairs, "hint", "".join(body) + "1" + "字、"))
# the real paragraph (0ea3ec86 p4), whole and from its sixth line on
_P4 = open("C:/tmp/x0e/p4_234_para.txt", encoding="utf-8").read()
ARMS.append(("p4full", "hint", _P4))
ARMS.append(("p4from6", "hint", _P4[_P4.index("るための法律（障害者総合支援法）"):]))
ARMS.append(("p4line6", "hint", "るための法律（障害者総合支援法）」とされた。平成25年４月からは、障害者"))
ARMS.append(("p4kana", "hint", "るための法律（あいうえおかきくけ）」とされた。平成25年４月からは、障害者"))
ARMS.append(("p4nopair", "hint", "るための法律（障害者総合支援法）とされた。平成25年４月からは、障害者"))
# the unit's mark: 。 instead of 、 (does the hanging mark's identity decide?)
for m in (1, 2, 4):
    for h in (1, 2):
        body = list(kana(19 - m))
        step = len(body) // (m + 1)
        for k in range(m):
            body.insert((k + 1) * step + k, "、")
        ARMS.append(("k%d_h%d" % (m, h), "hint", "".join(body) + "1" * h + "字。"))
ARMS.append(("p4comma", "hint", "るための法律（障害者総合支援法）」とされた、平成25年４月からは、障害者"))
ARMS.append(("p4half", "hint", "るための法律（障害者総合支援法）」とされ1た。平成25年４月からは、障害者"))
# no digits: 20 cells of text (kana + marks) then 字、 -> 字 needs exactly 1.0 cell
KATA = "アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワ"
_KR = [0]
def kata(n):
    _KR[0] += 5
    r = _KR[0] % len(KATA)
    return ((KATA[r:] + KATA[:r]) * 3)[:n]
def spread(n_kana, marks):
    body = list(kata(n_kana))
    step = len(body) // (len(marks) + 1)
    for k, mk in enumerate(marks):
        body.insert((k + 1) * step + k, mk)
    return "".join(body)
ARMS.append(("g_m2", "hint", spread(18, ["、", "、"]) + "字、"))                       # 2 、 (1.0 if 、 counts)
ARMS.append(("g_m4", "hint", spread(16, ["、", "、", "、", "、"]) + "字、"))            # 4 、
ARMS.append(("g_b2", "hint", (lambda k: k[:8] + "（" + k[8:10] + "）" + k[10:18])(kata(18)) + "字、"))   # （xx） 2 brackets
ARMS.append(("g_b4", "hint", (lambda k: k[:4] + "（" + k[4:6] + "）" + k[6:10] + "（" + k[10:12] + "）" + k[12:16])(kata(16)) + "字、"))  # 4 brackets
ARMS.append(("g_bp", "hint", (lambda k: k[:6] + "（" + k[6:14] + "）」" + k[14:17])(kata(17)) + "字、"))  # （8 chars）」 like p4 : 3 brackets, one pair
ARMS.append(("g_bo", "hint", (lambda k: k[:6] + "「" + k[6:14] + "」" + k[14:18])(kata(18)) + "字、"))  # 「8 chars」 2 brackets
ARMS.append(("g_n2", "hint", spread(19, ["、", "、"]) + "字"))                          # normal char, needs 1.0
# normal character (no mark after it) pulled in by standalone marks of each kind, demand 0.5 (one digit) or 1.0
def solo_line(kind, n_marks, n_kana):
    body = list(kata(n_kana))
    step = len(body) // (n_marks + 1)
    for k in range(n_marks):
        body.insert((k + 1) * step + k, kind)
    return "".join(body)
for kind, tag in (("）", "close"), ("（", "open"), ("・", "dot"), ("、", "comma"), ("」", "cbr")):
    ARMS.append(("z_%s1_h1" % tag, "hint", solo_line(kind, 1, 18) + "1字"))    # 19 + 0.5 → 字 needs 0.5
    ARMS.append(("z_%s2_h1" % tag, "hint", solo_line(kind, 2, 17) + "1字"))    # two marks, needs 0.5
    ARMS.append(("z_%s2_n" % tag, "hint", solo_line(kind, 2, 18) + "字"))       # 20 + 字 → needs 1.0
ARMS.append(("z_pair_h1", "hint", (lambda k: k[:6] + "（" + k[6:12] + "）」" + k[12:16])(kata(16)) + "1字"))  # pair + digit: 19.5+字 → needs 0.5 (pair gives 1.0)
ARMS.append(("z_unit_close", "hint", solo_line("）", 2, 17) + "1字、"))           # unit with standalone ）: needs 0.5 then hang
# demand 0.5 with the digit MID-line (the wrap candidate follows a kana, not the digit)
def mid_line(kind, n_marks):
    body = list(kata(18))
    body.insert(9, "1")
    step = len(body) // (n_marks + 1)
    for k in range(n_marks):
        body.insert((k + 1) * step + k, kind)
    return "".join(body)
for kind, tag in (("）", "close"), ("（", "open"), ("・", "dot"), ("、", "comma"), ("」", "cbr"), ("。", "period")):
    ARMS.append(("y_%s1" % tag, "hint", mid_line(kind, 1) + "字"))
    ARMS.append(("y_%s2" % tag, "hint", mid_line(kind, 2) + "字"))
ARMS.append(("y_none", "hint", "".join(list(kata(19))[:9] + ["1"] + list(kata(19))[9:19]) + "字"))
ARMS.append(("y_unit_close1", "hint", mid_line("）", 1) + "字、"))
# w_ arms: is the 0.5 grant about the mark's ADJACENCY to the half-width digit, or about
# where the digit sits (mid-line vs the line end)?  plus the p8 doc line, under-full
# adjacency (unconditional collapse?), and tracked paragraphs (p31 fits 21 chars to the
# true edge 235.98 with （ ） compressed 0.16 cell each).
def sep_line(kind, gap):
    body = list(kata(18))
    body.insert(9, "1")
    body.insert(9 - gap, kind)   # mark `gap` kana before the digit (gap=0 adjacent)
    return "".join(body)
ARMS.append(("w_sep1", "hint", sep_line("）", 1) + "字"))
ARMS.append(("w_sep3", "hint", sep_line("）", 3) + "字"))
ARMS.append(("w_sep6", "hint", sep_line("）", 6) + "字"))
ARMS.append(("w_after1", "hint", (lambda b: "".join(b[:10] + ["）"] + b[10:]))(list(kata(18)[:9]) + ["1"] + list(kata(18)[9:])) + "字"))  # 1）
ARMS.append(("w_endadj", "hint", kata(18) + "）1字"))                       # mark adjacent to the END digit
ARMS.append(("w_endsep", "hint", kata(17) + "）" + kata(1) + "1字"))      # one kana between mark and end digit (same as z)
ARMS.append(("w_dfirst", "hint", "1" + (lambda k: k[:9] + "）" + k[9:])(kata(18)) + "字"))   # digit first, mark mid
ARMS.append(("w_dfirst_adj", "hint", "1）" + kata(18) + "字"))
ARMS.append(("w_kanji_end", "hint", "移動作等関連項目身体介護調理洗濯掃）家事1字"))    # kanji body (17+）+1 = 18.5), digit at end
ARMS.append(("w_kanji_mid", "hint", "移動作等関連項目身1体介護調理洗濯掃）家事字"))    # kanji body, digit mid, mark far
ARMS.append(("w_end1gap", "hint", kata(17) + "）" + "1" + kata(1) + "字"))          # digit one kana before the end
ARMS.append(("w_end2gap", "hint", kata(16) + "）" + "1" + kata(2) + "字"))          # digit two kana before the end
ARMS.append(("w_endhira", "hint", kata(18) + "）1の"))                              # hiragana after the end digit
ARMS.append(("w_endcomma", "hint", kata(18) + "）1、字"))                           # comma after the end digit (needs 0.5 for 、 then 字)
ARMS.append(("w_endalpha", "hint", kata(17) + "）ab字"))                            # Latin letters at the end
ARMS.append(("w_enddigit2", "hint", kata(17) + "）12字"))                           # two digits at the end: 17+1+1 = 19, 字 fits without compression
ARMS.append(("w_enddigit2b", "hint", kata(18) + "）12字"))                          # 18+1+1 = 20, 字 needs 1.0
ARMS.append(("w_p8", "hint", "な医療に関連する項目(12項目）の計80項目の" + FILL))
ARMS.append(("w_p8_fw", "hint", "な医療に関連する項目（12項目）の計80項目の" + FILL))   # full-width （: 21 cells, needs 1.0
ARMS.append(("w_p8_end", "hint", "な医療に関連する項目(12項目）の計項目の80" + FILL))  # digits at the end
ARMS.append(("w_under_adj", "hint", (lambda k: k[:8] + "）1" + k[8:17])(kata(17)) + "字"))     # 19.5 cells: no pressure
ARMS.append(("w_under_sep", "hint", (lambda k: k[:5] + "）" + k[5:11] + "1" + k[11:17])(kata(17)) + "字"))
# tracked paragraphs (w:spacing in twips): 21 chars with two standalone marks like p31
ARMS.append(("w_tr2_21m", -2, "及び（" + kata(2) + "）" + kata(15)))          # 21 x 11.39 = 239.2: needs 3.2 to the true edge
ARMS.append(("w_tr2_21", -2, kata(21)))                                        # no marks: refuse expected
ARMS.append(("w_tr6_21", -6, kata(21)))                                        # 21 x 11.19 = 235.0 < 235.98: floor or edge?
ARMS.append(("w_tr6_21m", -6, "及び（" + kata(2) + "）" + kata(15)))
ARMS.append(("w_tr4_21m", -4, "及び（" + kata(2) + "）" + kata(15)))          # 21 x 11.29 = 237.1: needs 1.1
ARMS.append(("w_tr0_21m", 0, "及び（" + kata(2) + "）" + kata(15)))           # untracked: 21 cells, 2 standalone marks, needs 1.0
ARMS.append(("w_tr2_20m1", -2, "及び（" + kata(2) + "）" + kata(14) + "1"))   # 20.5 x 11.39 ~ 233.5 < 235.98
# v_ arms: why is the rescue refused when the overflowing CJK char directly follows a
# half-width digit?  autospace off, two CJK chars after the digit, a w:r boundary at the
# overflow char, an under-full pair (is the pair's collapse unconditional?), and the
# standalone cap swept with tracked 21-char lines (over the floor by 0.46 / 0.64 / 0.83).
ARMS.append(("v_z_noas", "hint", solo_line("）", 1, 18) + "1字"))
ARMS.append(("v_endadj_noas", "hint", kata(18) + "）1字"))
ARMS.append(("v_chunk2", "hint", solo_line("）", 1, 18) + "1字字"))
ARMS.append(("v_runsplit", "hint", sep_line("）", 3) + "|字"))
ARMS.append(("v_runsplit1", "hint", (lambda t: t[:-1] + "|" + t[-1])(sep_line("）", 3)) + "字"))
ARMS.append(("v_under_pair", "hint", (lambda k: k[:7] + "（" + k[7:11] + "）」" + k[11:16])(kata(16)) + "字"))
ARMS.append(("v_tr3_21m", -3, "及び（" + kata(2) + "）" + kata(15)))          # over the floor by 0.46 cell, 2 standalone marks
ARMS.append(("v_tr3_21m1", -3, "及び" + kata(3) + "）" + kata(15)))           # 0.46 from ONE standalone mark
ARMS.append(("v_tr2_21m3", -2, "及び（" + kata(2) + "）" + kata(6) + "、" + kata(8)))   # 0.64, 3 standalone marks
ARMS.append(("v_tr2_21m4", -2, "及び（" + kata(2) + "）" + kata(4) + "、" + kata(4) + "、" + kata(3)))  # 0.64, 4 marks
ARMS.append(("v_tr1_21m", -1, "及び（" + kata(2) + "）" + kata(15)))          # 0.83, 2 marks: refuse expected
ARMS.append(("v_tr2_21pair", -2, "及び（" + kata(2) + "）」" + kata(14)))     # 0.64 with a pair: grant expected
# t_ arms: which "previous element" blocks the rescue?  Latin letter, full-width digit,
# a Latin word itself overflowing, a 19-cell line with a pair (unconditional collapse?),
# and the doc's p31 shape (untracked leading 　 run + tracked run).
ARMS.append(("t_endletter", "hint", kata(18) + "）a字"))                            # 18+1+0.5, 字 after a Latin letter
ARMS.append(("t_endfwdigit", "hint", (lambda k: k[:9] + "1" + k[9:17])(kata(17)) + "）１字"))   # 17+0.5+1+1, 字 after full-width １
ARMS.append(("t_latword", "hint", (lambda k: k[:9] + "1" + k[9:18])(kata(18)) + "）12"))       # 18+0.5+1, the word 12 overflows by 0.5
ARMS.append(("t_latword_end", "hint", kata(18) + "）" + "12"))                                # 18+1 = 19, 12 fits: control
ARMS.append(("t_latword3", "hint", kata(18) + "）" + "123"))                                  # 18+1+1.5 = 20.5: the word 123 overflows by 0.5
ARMS.append(("t_under_pair19", "hint", (lambda k: k[:7] + "（" + k[7:11] + "）」" + k[11:15])(kata(15)) + "字"))  # 19 cells
ARMS.append(("t_end2latin", "hint", kata(17) + "）" + "1" + "2字"))                          # same as 12字 with the digits split in two runs? no: control 17+1+1 = 19 + 字 fits
ARMS.append(("t_endlat_sp", "hint", kata(18) + "）1 字"))                                    # ASCII space between the digit and 字
ARMS.append(("t_p31shape", "hint", "SPACE" + "また、障害者の職業的自立を図るため、職業訓練を行う施設は、東京障害者職業能力開発校及び（公財）東京しごと財団障害者就業支援課がある。"))
ARMS.append(("t_p31tracked", -4, "SPACE" + "また、障害者の職業的自立を図るため、職業訓練を行う施設は、東京障害者職業能力開発校及び（公財）東京しごと財団障害者就業支援課がある。"))
# u_ arms: character scale (w:w) on the grid -- 0ea3ec86 p12/p13/p25 draw a 90% run at
# 10.34 (cell 11.49 x 0.9) / 10.6 (11.76 x 0.9): does the CELL scale, and the floor?
ARMS.append(("u_w90_23", "w90", kata(23)))                     # 22 x 10.34 = 227.5 fit, 23 = 237.8 over the floor
ARMS.append(("u_w90_22f", "w90", kata(22) + "字"))
ARMS.append(("u_w80_27", "w80", kata(27)))                     # 25 x 9.19 = 229.8 (floor 229.7)
ARMS.append(("u_w110_20", "w110", kata(20)))                   # 18 x 12.64 = 227.5, 19 = 240
ARMS.append(("u_w150_15", "w150", kata(15)))                   # 13 x 17.2 = 224, 14 = 241
ARMS.append(("u_w90_mix", "w90", "担当課　|" + kata(18)))       # 4 x 11.49 + 17 x 10.34 = 221.8, 18th = 232.1
ARMS.append(("u_w90_mark", "w90", kata(20) + "、" + kata(3)))   # 22 = 227.5, the 23rd needs 8.1 > 5.17
ARMS.append(("u_w90_half", "w90", kata(10) + "1" + kata(11) + "字"))          # 22 kana + 0.5 = 232.7: 字 over by 3.0, no mark
ARMS.append(("u_w90_halfm", "w90", kata(10) + "、1" + kata(10) + "字"))       # 21 kana + 、 + 1 = 232.7: 字 over by 3.0, 、 gives 5.17
ARMS.append(("u_w90_digits", "w90", kata(18) + "1234"))        # 18 x 10.34 + 4 x 5.17 = 206.8: control for the half cell at 90%
# d_ arms: the middle dot at the line end. 0ea3ec86 p30 「どに入院・入所中の児童・生徒の
# ために病院・」 keeps 21 (two mid ・ compressed 0.33 each, the final ・ past the floor) while
# p4 「…平成27年１|月・７月」 sends 月・ down with one 、 on the line. Hang, 追い出し or compress?
ARMS.append(("d_ctrl20", "hint", kata(19) + "・" + kata(3)))            # 19 + ・ = 20 cells: fits
ARMS.append(("d_dot_end", "hint", kata(20) + "・" + kata(3)))           # 20 + ・: hang (21) or 追い出し (19)?
ARMS.append(("d_comma_end", "hint", kata(20) + "、" + kata(3)))         # control: 、 hangs
ARMS.append(("d_dots3", "hint", kata(4) + "・" + kata(6) + "・" + kata(8) + "・" + kata(3)))   # 18 + 3 ・ = 21 like p30
ARMS.append(("d_commas_dot", "hint", kata(4) + "、" + kata(6) + "、" + kata(8) + "・" + kata(3)))  # 、、 mid, ・ end: 21
ARMS.append(("d_dots_norm", "hint", kata(4) + "・" + kata(6) + "・" + kata(9) + "字" + kata(3)))   # ・・ mid, normal char needs 1.0
ARMS.append(("d_dots3_norm", "hint", kata(3) + "・" + kata(5) + "・" + kata(5) + "・" + kata(5) + "字" + kata(3)))   # 3 ・ mid, normal needs 1.0
ARMS.append(("d_dots_commaend", "hint", kata(4) + "・" + kata(6) + "・" + kata(8) + "、" + kata(3)))  # ・・ mid, 、 end (hang)
ARMS.append(("d_p30", "hint", "どに入院・入所中の児童・生徒のために病院・施設内の学級で教育を受けることができる。"))
ARMS.append(("d_comma_dot", "hint", kata(18) + "、" + kata(1) + "・" + kata(3)))   # like p4: one 、 mid, ・ needs 1.0
ARMS.append(("d_dot_dot", "hint", kata(18) + "・" + kata(1) + "・" + kata(3)))     # one ・ mid, ・ end
ARMS.append(("d_close_dot", "hint", kata(18) + "）" + kata(1) + "・" + kata(3)))   # one ） mid, ・ end
ARMS.append(("d_dot_half", "hint", kata(9) + "1" + kata(10) + "・" + kata(3)))    # 19.5 + ・: needs 0.5 from itself
# natural (unjustified, under-full) advances of scaled runs: is the scaled cell exactly
# pitch x w/100 (80%: 9.19) or rounded?  u_w80_27 refused the 25th char at 25 x 9.19 = 229.8.
for w in ("50", "80", "90", "110", "125"):
    ARMS.append(("u_w%s_nat" % w, "w" + w, kata(10)))
ARMS.append(("u_w80_26", "w80", kata(26)))
ARMS.append(("u_w80_25f", "w80", kata(24) + "字"))     # 25 chars = 229.8: the equality case again
ARMS.append(("u_w80_25m", "w80", kata(12) + "、" + kata(11) + "字"))   # 25 with a 、: 0.5 elective at 80% = 4.6
# scale sweep: the natural advance of a scaled run against the grid, 10 chars each
for w in ("33", "40", "60", "66", "70", "75", "85", "95", "100", "105", "115", "120", "133", "150", "200"):
    ARMS.append(("u_w%s_nat" % w, "w" + w, kata(10)))
for w in ("50", "80", "90", "110", "125"):
    ARMS.append(("u_w%s_nat12", "w" + w + "s24", kata(10)))   # 12pt run under the 10.5pt grid
    ARMS[-1] = ("u_w%s_nat12" % w, "w" + w + "s24", kata(10))
ARMS.append(("u_w100_nat12", "w100s24", kata(10)))
ARMS.append(("u_w100_nat9", "w100s18", kata(10)))
ARMS.append(("u_w80_nat9", "w80s18", kata(10)))
# L_ arms: the elective cap for a LATIN word. 0ea3ec86 p13 「手　　続　区市役所、町村役場、
# 各島支庁（303」 (20 glyphs + 303) keeps 303 by halving 、 、 （ = 1.5 cells, three times the
# 0.5 a normal character gets.
ARMS.append(("L_a", "hint", kata(17) + "、" + kata(1) + "、" + "123"))                 # needs 1.5, two 、 (1.0)
ARMS.append(("L_b", "hint", kata(16) + "、" + kata(1) + "、" + kata(1) + "（" + "123"))  # needs 1.5, 、、（ (1.5) like p13
ARMS.append(("L_d", "hint", kata(17) + "、" + kata(1) + "、" + "12"))                  # needs 1.0, two 、
ARMS.append(("L_e", "hint", kata(18) + "、" + kata(1) + "12"))                        # needs 1.0, one 、
ARMS.append(("L_f", "hint", kata(18) + "、" + kata(1) + "1"))                         # needs 0.5, one 、
ARMS.append(("L_g", "hint", kata(15) + "、" + kata(1) + "、" + kata(1) + "）" + kata(1) + "字"))   # CJK needs 1.0, three marks
ARMS.append(("L_g2", "hint", kata(14) + "、" + kata(1) + "、" + kata(1) + "）" + kata(2) + "字"))  # CJK needs 1.0 (20 glyphs), three marks
ARMS.append(("L_h", "hint", "手　　続　区市役所、町村役場、各島支庁（303㌻）へ" + FILL))
ARMS.append(("L_i", "hint", "手　　続　区市役所、町村役場、各島支庁（303字）へ" + FILL))
ARMS.append(("L_j", "hint", kata(16) + "、" + kata(1) + "、" + kata(1) + "（" + "1234"))  # needs 2.0, 1.5 available: refuse?
ARMS.append(("L_k", "hint", kata(15) + "、" + kata(1) + "、" + kata(1) + "）" + kata(1) + "（" + "123"))  # needs 1.5, four marks
# L2_ arms: the same Latin words with CJK text FOLLOWING them (the doc's 303 is followed by ㌻）へ)
ARMS.append(("L2_a", "hint", kata(17) + "、" + kata(1) + "、" + "123" + "字あいう"))
ARMS.append(("L2_b", "hint", kata(16) + "、" + kata(1) + "、" + kata(1) + "（" + "123" + "字）あいう"))
ARMS.append(("L2_d", "hint", kata(17) + "、" + kata(1) + "、" + "12" + "字あいう"))
ARMS.append(("L2_e", "hint", kata(18) + "、" + kata(1) + "12" + "字あいう"))
ARMS.append(("L2_h2", "hint", "手あい続う区市役所、町村役場、各島支庁（303㌻）へ" + FILL))   # the doc line, spaces -> kana
ARMS.append(("L2_h3", "hint", "手あい続う区市役所、町村役場、各島支庁（303字）へ" + FILL))
ARMS.append(("L2_h4", "hint", "手あい続う区市役所、町村役場、各島支庁（303" + FILL))      # 303 then kana
# M_ arms: why does the doc line get 1.5 for 「303」 when the synthetic 「、ネ、ノ（123」 gets
# nothing?  mark spacing (5 apart vs 1-2), kanji vs kana body, the size of the need.
ARMS.append(("M_b", "hint", "手あい続う区市役所、町村役場、各島支庁（30" + FILL))        # doc shape, needs 1.0
ARMS.append(("M_c", "hint", "手あい続う区市役所、町村役場は各島支庁（303" + FILL))      # doc shape, 2 marks (1.0), needs 1.5
ARMS.append(("M_c2", "hint", "手あい続う区市役所は町村役場は各島支庁（303" + FILL))     # 1 mark (（), needs 1.5
ARMS.append(("M_c3", "hint", "手あい続う区市役所、町村役場、各島支庁は303" + FILL))      # 2 、, no （, needs 1.5
ARMS.append(("M_d", "hint", kata(4) + "、" + kata(4) + "、" + kata(4) + "、" + kata(4) + "（" + "303" + FILL))   # kana, 4 marks spaced 4, needs 1.5
ARMS.append(("M_d2", "hint", kata(5) + "、" + kata(5) + "、" + kata(5) + "（" + "303" + "字" + FILL))       # kana, marks spaced 5 (3 marks), needs 1.5
ARMS.append(("M_e", "hint", kata(9) + "、" + kata(5) + "、" + kata(4) + "12" + "字" + FILL))              # kana, 2 marks spaced 5, Latin needs 1.0
ARMS.append(("M_f", "hint", "手あい続う区市役所、町村役場、各島支庁）字あ" + FILL))       # doc shape, CJK char needs 1.0, 3 marks
ARMS.append(("M_i", "hint", "手あい続う区市役所、町村役場、各島支庁（12" + FILL))        # doc shape, Latin needs 1.0, 3 marks
ARMS.append(("M_j", "hint", "手あい続う区市役所、町村役場、各島支庁は12" + FILL))        # 2 、 spaced 5, Latin needs 1.0, no （
ARMS.append(("M_k", "hint", "手あい続う区市役所、町村役場、各島支庁は字あ" + FILL))      # 2 、 spaced 5, CJK needs 1.0
# N_ arms: the （+word unit -- is the （'s blank free and the elective cap one cell?  mark
# spacing, kanji vs kana, and a CJK character after the （.
ARMS.append(("N_1", "hint", "手あい続う区市役所町村役場、各島、支庁（303" + FILL))    # 、 at 15, 18
ARMS.append(("N_2", "hint", "手あい続う区市役所町村役場、各島支、庁（303" + FILL))    # 、 at 15, 19
ARMS.append(("N_3", "hint", "手あい続う区市役所町村役場各島、支、庁（303" + FILL))    # 、 at 16, 18
ARMS.append(("N_4", "hint", kata(9) + "、" + kata(4) + "、" + kata(4) + "（303" + FILL))   # kana, 、 at 10, 15
ARMS.append(("N_5", "hint", "区市役所町村役場各島支庁事務局、課、係（303" + FILL))      # kanji, 、 at 16, 18
ARMS.append(("N_6", "hint", "手あい続う区市役所、町村役場、各島支庁（字あ" + FILL))     # （ + CJK needs 1.0, 、、（
ARMS.append(("N_7", "hint", "手あい続う区市役所は町村役場は各島支庁（字あ" + FILL))     # （ + CJK needs 1.0, （ only
ARMS.append(("N_8", "hint", "手あい続う区市役所、町村役場、各島支庁（３０" + FILL))     # （ + full-width digits, needs 2.0
ARMS.append(("N_9", "hint", "手あい続う区市役所、町村役場、各島支庁「303" + FILL))     # 「 instead of （
ARMS.append(("N_10", "hint", "手あい続う区市役所、町村役場、各島支庁（3" + FILL))       # （ + 1 digit: needs 0.5
# O_ arms: 「…、ウ、ク（123」 (kana between close marks) refuses what 「…各島、支、庁（303」 grants
ARMS.append(("O_1", "hint", "手あい続う区市役所町村役場各島、ア、カ（303" + FILL))
ARMS.append(("O_2", "hint", "手あい続う区市役所町村役場各島、支、カ（303" + FILL))
ARMS.append(("O_3", "hint", "手あい続う区市役所町村役場各島、ア、庁（303" + FILL))
ARMS.append(("O_4", "hint", kata(15) + "、支、庁（303" + FILL))
ARMS.append(("O_5", "hint", kata(15) + "、あ、い（303" + FILL))
ARMS.append(("O_6", "hint", "アイウエオカキクケコサシスセソタ、ア、カ（303" + FILL))
ARMS.append(("O_7", "hint", "アイウエオカキクケコサシスセソタ、ナ、ハ（303" + FILL))
ARMS.append(("O_8", "hint", "アイウエオカキクケコサシスセソタ、ナ、ハ（123" + FILL))
ARMS.append(("O_9", "hint", "アイウエオカキクケコサシスセソタ、ナ、ハ（123"))          # paragraph ends after 123 (like L_b)
ARMS.append(("O_10", "hint", "手あい続う区市役所町村役場各島、支、庁（303"))          # N_3 but paragraph-final
# equality edge: the katakana lines refuse exactly at need == budget (1.5 = 3 x 0.5); is it
# sub-twip noise?  need 1.0 with 3 marks (slack 0.5), need 1.5 with 4 marks, equality at 1.0.
ARMS.append(("O_11", "hint", "アイウエオカキクケコサシスセソタ、ナ、ハ（30" + FILL))        # katakana, need 1.0, 3 marks
ARMS.append(("O_12", "hint", "アイウエオカキクケコサシ、セソ、タ、ナ、ハ（303" + FILL))     # katakana, need 1.5, 5 marks
ARMS.append(("O_13", "hint", "アイウエオカキクケコサシスセソタチツ、ハ（30" + FILL))        # katakana, equality 1.0 = 、（
ARMS.append(("O_14", "hint", "手あい続う区市役所町村役場各島支庁事務、局（30" + FILL))      # kanji, equality 1.0 = 、（
ARMS.append(("O_15", "hint", "あいうえおかきくけこさしすせそた、な、は（303" + FILL))        # hiragana body, equality 1.5
ARMS.append(("O_16", "hint", "アイウエオカキクケコサシスセソタ、ナ、ハ（3" + FILL))         # katakana, need 0.5, 3 marks
# P_ arms: minimal variations of the refused 「アイウエオカキクケコサシスセソタ、ナ、ハ（3」
ARMS.append(("P_1", "hint", "ネノハヒフヘホマミムメモヤユヨラ、ナ、ハ（3" + FILL))
ARMS.append(("P_2", "hint", "アイウエオカキクケコサシスセソタ、ナ、ハ（3" + FILL))
ARMS.append(("P_3", "hint", "イウエオカキクケコサシスセソタチ、ナ、ハ（3" + FILL))
ARMS.append(("P_4", "hint", "アイウエオカキクケコサシスセソ島、ナ、ハ（3" + FILL))
ARMS.append(("P_5", "hint", "手イウエオカキクケコサシスセソタ、ナ、ハ（3" + FILL))
ARMS.append(("P_6", "hint", "アイウエオカキクケコサシスセソタ、支、庁（3" + FILL))
ARMS.append(("P_7", "hint", "アイウエオカキクケコサシスセソタ、ナ、ハ（3"))
ARMS.append(("P_8", "hint", kata(15) + "、ナ、ハ（3" + FILL))
ARMS.append(("P_9", "hint", kata(16) + "、ナ、ハ（3" + FILL))
ARMS.append(("P_10", "hint", "アイウエオカキクケコサシスセソタ、ナ、ハ（3字" + FILL))
# Q_ arms: the （-unit budget = (n 、 + 1) x 0.5?  one 、 + （ refused a need of 1.0 (O_14) while
# two 、 + （ granted 1.5 (N_5).  (the O_6-O_16 / P_ / L_b arms had 21 glyphs before the （ -- miscounted)
ARMS.append(("Q_1", "hint", "手あい続う区市役所町村役場各島支庁事務、局（3" + FILL))       # one 、, need 0.5
ARMS.append(("Q_2", "hint", "手あい続う区市役所町村役場各島支庁事務、局（30" + FILL))      # one 、, need 1.0 (= O_14)
ARMS.append(("Q_3", "hint", "手あい続う区市役所町村役場各島支、庁事務局（30" + FILL))      # one 、 at 16, need 1.0
ARMS.append(("Q_4", "hint", "手あい続う区市役所、町村役場各島支庁事務局（30" + FILL))      # one 、 at 10, need 1.0
ARMS.append(("Q_5", "hint", "手あい続う区市役所、町村役場、各島支庁（30" + FILL))         # two 、, need 1.0 (= M_b)
ARMS.append(("Q_7", "hint", "手あい続う区市、役所、町村、役場各島支（303" + FILL))         # three 、, need 1.5
ARMS.append(("Q_8", "hint", "手あい続う区市、役所、町村、役場各島支（3033" + FILL))        # three 、, need 2.0
ARMS.append(("Q_9", "hint", "手あい続う区市役所、町村役場、各島支庁（3033" + FILL))        # two 、, need 2.0
ARMS.append(("Q_10", "hint", "手あい続う区市役所、町村役場、各島支庁（3033字" + FILL))     # two 、, need 2.0 then 字
ARMS.append(("Q_11", "hint", "手あい続う区市役所町村役場各島支庁事務、局（3字" + FILL))     # one 、, need 0.5 for 3, then 字 needs 1.0
# I_ arms: is the grid floor cut BEFORE the indent (20 cells - indent) or AFTER it
# (floor((true - indent)/cell))?  0ea3ec86 p19 「（グループホーム…地域活動」 (hangingChars 200 /
# hanging 471) holds 18 = 207.2 where 230.27 - 23.55 = 206.45 refuses.  ㋐'s natural width.
ARMS.append(("I_lc50", "hint", 'PPR{<w:ind w:leftChars="50" w:left="118"/>}' + kata(21)))      # before: 19, after: 20
ARMS.append(("I_l100", "hint", 'PPR{<w:ind w:left="100"/>}' + kata(21)))                        # 5pt: before 19, after 20
ARMS.append(("I_l200", "hint", 'PPR{<w:ind w:left="200"/>}' + kata(21)))                        # 10pt: 19 either way
ARMS.append(("I_rc50", "hint", 'PPR{<w:ind w:rightChars="50" w:right="118"/>}' + kata(21)))
ARMS.append(("I_h471c", "hint", 'PPR{<w:ind w:left="471" w:hangingChars="200" w:hanging="471"/>}' + kata(20) + "字" + kata(19) + "字"))
ARMS.append(("I_h471", "hint", 'PPR{<w:ind w:left="471" w:hanging="471"/>}' + kata(20) + "字" + kata(19) + "字"))
ARMS.append(("I_h500", "hint", 'PPR{<w:ind w:left="500" w:hanging="500"/>}' + kata(20) + "字" + kata(19) + "字"))   # 25pt: before 17, after 18
ARMS.append(("I_h300", "hint", 'PPR{<w:ind w:left="300" w:hanging="300"/>}' + kata(20) + "字" + kata(19) + "字"))   # 15pt: before 18 (215.3: 18 = 207.2 ok, 19 = 218.7 no), after floor(19.2) = 19
ARMS.append(("I_enc", "hint", "㋐" + kata(9)))
ARMS.append(("I_enc3", "hint", "㋐㋑㋒" + kata(7)))
ARMS.append(("I_enc19", "hint", 'PPR{<w:ind w:leftChars="100" w:left="236"/>}' + "㋐" + kata(18) + "字" + FILL))   # the doc line's shape: 19 cells, ㋐ + 18 + 字
# Z_ arms: a line-start-prohibited character after the character that fills the floor
ARMS.append(("Z_zwj", "hint", kata(19) + "ホ|‍|ーム" + FILL))
ARMS.append(("Z_nozwj", "hint", kata(19) + "ホーム" + FILL))
ARMS.append(("Z_small", "hint", kata(19) + "ホッと" + FILL))
ARMS.append(("Z_close0", "hint", kata(19) + "字」と" + FILL))
ARMS.append(("Z_close1", "hint", kata(9) + "、" + kata(9) + "字」と" + FILL))
ARMS.append(("Z_close2", "hint", kata(8) + "、" + kata(5) + "、" + kata(4) + "字」と" + FILL))
ARMS.append(("Z_dash2", "hint", kata(8) + "、" + kata(5) + "、" + kata(4) + "字ーと" + FILL))
ARMS.append(("Z_small2", "hint", kata(8) + "、" + kata(5) + "、" + kata(4) + "ホッと" + FILL))
# hanging-indent geometry: which of left / leftChars / hanging / hangingChars sets the
# continuation line (the doc's hangingChars=200 hanging=471 reads 23.1 = 2 cells, not 23.55)
ARMS.append(("I_hb", "hint", 'PPR{<w:ind w:left="600" w:hangingChars="200" w:hanging="471"/>}' + kata(20) + "字" + kata(19) + "字"))
ARMS.append(("I_hc", "hint", 'PPR{<w:ind w:leftChars="300" w:left="708" w:hangingChars="200" w:hanging="471"/>}' + kata(20) + "字" + kata(19) + "字"))
ARMS.append(("I_hd", "hint", 'PPR{<w:ind w:left="471" w:hangingChars="200" w:hanging="400"/>}' + kata(20) + "字" + kata(19) + "字"))
ARMS.append(("I_he", "hint", 'PPR{<w:ind w:left="471" w:hanging="400"/>}' + kata(20) + "字" + kata(19) + "字"))
ARMS.append(("I_hf", "hint", 'PPR{<w:ind w:leftChars="200" w:left="471" w:hangingChars="200" w:hanging="471"/>}' + kata(20) + "字" + kata(19) + "字"))
# ㋐ line: 19 cells with leftChars=100 and the twip varied -- equality at the 19th glyph
ARMS.append(("I_e230", "hint", 'PPR{<w:ind w:leftChars="100" w:left="230"/>}' + "㋐" + kata(17) + "字" + FILL))
ARMS.append(("I_e236", "hint", 'PPR{<w:ind w:leftChars="100" w:left="236"/>}' + "㋐" + kata(17) + "字" + FILL))
ARMS.append(("I_e250", "hint", 'PPR{<w:ind w:leftChars="100" w:left="250"/>}' + "㋐" + kata(17) + "字" + FILL))
ARMS.append(("I_e236k", "hint", 'PPR{<w:ind w:leftChars="100" w:left="236"/>}' + kata(18) + "字" + FILL))       # no ㋐
ARMS.append(("I_e0", "hint", kata(19) + "字" + FILL))                                                          # 20 cells, no indent: equality control
# R_ arms: X before a HANGING mark -- 0ea3ec86 p9 「…「一般２」となる。」 (-8) keeps る with three
# marks compressed 0.36 each (1.07 cells) and the 。 hanging; the v4 arms refused 1.0.  Tracked
# lines put the need at 0.46 / 0.64 / 0.83 / 1.0.
for sp, tag in ((-3, "046"), (-2, "064"), (-1, "083"), (0, "100")):
    ARMS.append(("R_m3_%s" % tag, sp, kata(4) + "、" + kata(5) + "、" + kata(5) + "」" + kata(3) + "字、" + FILL))
ARMS.append(("R_m2_064", -2, kata(5) + "、" + kata(7) + "、" + kata(6) + "字、" + FILL))
ARMS.append(("R_m1_064", -2, kata(9) + "、" + kata(10) + "字、" + FILL))
ARMS.append(("R_m4_083", -1, kata(3) + "、" + kata(4) + "、" + kata(4) + "」" + kata(3) + "、" + kata(3) + "字、" + FILL))
ARMS.append(("R_m3_064_norm", -2, kata(4) + "、" + kata(5) + "、" + kata(5) + "」" + kata(3) + "字あ" + FILL))
ARMS.append(("R_m3_064_period", -2, kata(4) + "、" + kata(5) + "、" + kata(5) + "」" + kata(3) + "字。" + FILL))
ARMS.append(("R_doc_mix", -8, "※|入所利用者(20歳以上）、グループホーム利用者は、区市町村民税課税世帯の場合、「一般２」となる。"))
# S_ arms (ONE-COLUMN section 5, 42 cells): X before a hanging mark with the need swept by
# tracking -- 42 glyphs (three marks) + 字、 at spacing 0 (need 1.0), -1 (0.55), -2 (0.17);
# 2 and 1 marks; and the doc's -8 shape (※ untracked + 44 tracked glyphs + 字。).
for sp, tag in ((0, "100"), (-1, "055"), (-2, "017")):
    ARMS.append(("S_m3_%s" % tag, sp, kata(10) + "、" + kata(12) + "、" + kata(12) + "」" + kata(4) + "字、" + FILL))
    ARMS.append(("S_m3_%s_norm" % tag, sp, kata(10) + "、" + kata(12) + "、" + kata(12) + "」" + kata(4) + "字あ" + FILL))
ARMS.append(("S_m2_100", 0, kata(12) + "、" + kata(14) + "、" + kata(13) + "字、" + FILL))
ARMS.append(("S_m1_100", 0, kata(20) + "、" + kata(20) + "字、" + FILL))
ARMS.append(("S_m1_055", -1, kata(20) + "、" + kata(20) + "字、" + FILL))
ARMS.append(("S_m2_055", -1, kata(12) + "、" + kata(14) + "、" + kata(13) + "字、" + FILL))
ARMS.append(("S_doc_tail", -8, "※|入所利用者(20歳以上）、グループホーム利用者は、区市町村民税課税世帯の場合、「一般２」となる。" + FILL))
ARMS.append(("S_doc_norm", -8, "※|入所利用者(20歳以上）、グループホーム利用者は、区市町村民税課税世帯の場合、「一般２」となるが" + FILL))
ARMS.append(("S_doc_m8", -8, "※|" + kata(20) + "、" + kata(12) + "、" + kata(8) + "」" + kata(3) + "字、" + FILL))   # 44 tracked glyphs + 字、 (need ~0.8)
# J_ arms: the unit of leftChars vs hangingChars when the paragraph is NOT the default size
# (d6fd9a51 8pt: hanging 160 = 1 char x 8; a1d6e4ef 9pt: hanging 380/203 = 9.36 but left (489-380)/50 = 10.9)
ARMS.append(("J_12_lh", "sz24", 'PPR{<w:ind w:leftChars="200" w:hangingChars="100"/>}' + kata(20) + "字" + kata(19) + "字"))
ARMS.append(("J_12_h", "sz24", 'PPR{<w:ind w:hangingChars="100"/>}' + kata(20) + "字" + kata(19) + "字"))
ARMS.append(("J_12_l", "sz24", 'PPR{<w:ind w:leftChars="200"/>}' + kata(20) + "字" + kata(19) + "字"))
ARMS.append(("J_8_lh", "sz16", 'PPR{<w:ind w:leftChars="200" w:hangingChars="100"/>}' + kata(28) + "字" + kata(27) + "字"))
ARMS.append(("J_8_h", "sz16", 'PPR{<w:ind w:hangingChars="100"/>}' + kata(28) + "字" + kata(27) + "字"))
ARMS.append(("J_8_l", "sz16", 'PPR{<w:ind w:leftChars="200"/>}' + kata(28) + "字" + kata(27) + "字"))
# K_ arms (S1348): the grid advance of an OFF-SIZE run, run size x charSpace (gen --charspace=N)
for sz in (14, 16, 18, 20, 21, 22, 24, 28, 32):
    ARMS.append(("K_sz%d" % sz, "sz%d" % sz, kata(10)))
# control: normal character instead of the unit (demand 0.96 + 0.5h, no mark to pull)
for m in (1, 2, 3):
    for h in (0, 1, 2):
        body = list(kana(19 - m))
        step = len(body) // (m + 1)
        for k in range(m):
            body.insert((k + 1) * step + k, "、")
        text = "".join(body) + "1" * h + "字"
        ARMS.append(("n%d_h%d" % (m, h), 0, text))
def gen():
    os.makedirs(OUT, exist_ok=True)
    z = zipfile.ZipFile(SRC)
    doc = z.read("word/document.xml").decode("utf-8")
    secs = re.findall(r"<w:sectPr[ >].*?</w:sectPr>", doc, re.S)
    sect = secs[SECT_IDX]  # default 4: 2 columns, charSpace 2048; 18: 2 columns, charSpace 3194
    if CHARSPACE is not None:
        sect = re.sub(r'w:charSpace="-?[0-9]+"', 'w:charSpace="%d"' % CHARSPACE, sect)
    sect = re.sub(r"<w:(header|footer)Reference[^>]*/>", "", sect)
    sect = re.sub(r'w:rsid\w*="[^"]*" ?', "", sect)
    head = doc[:doc.index("<w:body>") + len("<w:body>")]
    body = ""
    for label, sp, text in ARMS:
        if ONLY and not label.startswith(ONLY):
            continue
        if isinstance(sp, str) and sp.startswith("sz"):  # plain size override
            rpr = '<w:rPr><w:rFonts w:hint="eastAsia"/><w:sz w:val="%s"/></w:rPr>' % sp[2:]
        elif isinstance(sp, str) and sp.startswith("w"):   # character scale w:w, optional s<sz>
            wv, _, szv = sp[1:].partition("s")
            rpr = '<w:rPr><w:rFonts w:hint="eastAsia"/><w:w w:val="%s"/>%s</w:rPr>' % (wv, '<w:sz w:val="%s"/>' % szv if szv else "")
        else:
            rpr = '<w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>' if sp == "hint" else ('<w:rPr><w:spacing w:val="%d"/></w:rPr>' % sp if sp else "")
        if text.startswith("SPACE"):
            body += ('<w:p><w:pPr><w:jc w:val="both"/></w:pPr><w:r><w:t xml:space="preserve">　</w:t></w:r>'
                     '<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r></w:p>' % (rpr, text[5:]))
        else:
            ppr_extra = ""
            if text.startswith("PPR{"):
                ppr_extra = text[4:text.index("}")]
                text = text[text.index("}") + 1:]
            ppr = '<w:pPr><w:jc w:val="both"/>%s%s</w:pPr>' % ('<w:autoSpaceDE w:val="0"/><w:autoSpaceDN w:val="0"/>' if label.endswith("_noas") else "", ppr_extra)
            if label.endswith("_mix"):   # first run plain, the rest scaled
                parts = text.split("|")
                runs = '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t xml:space="preserve">%s</w:t></w:r>' % parts[0]
                runs += "".join('<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>' % (rpr, part) for part in parts[1:])
            else:
                runs = "".join('<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>' % (rpr, part) for part in text.split("|"))
            body += '<w:p>%s%s</w:p>' % (ppr, runs)
        body += "<w:p/>"
    new = head + body + sect + "</w:body></w:document>"
    out = os.path.join(OUT, DOCNAME + ".docx")
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as o:
        for item in z.infolist():
            if item.filename == "word/document.xml":
                o.writestr(item, new.encode("utf-8"))
            elif item.filename.startswith("word/header") or item.filename.startswith("word/footer"):
                continue
            else:
                data = z.read(item.filename)
                if item.filename == "word/_rels/document.xml.rels":
                    data = re.sub(rb'<Relationship [^>]*Target="(header|footer)\d*\.xml"[^>]*/>', b"", data)
                if item.filename == "[Content_Types].xml":
                    data = re.sub(rb'<Override [^>]*PartName="/word/(header|footer)\d*\.xml"[^>]*/>', b"", data)
                o.writestr(item, data)
    print("wrote", out)


def pdf():
    import fitz
    import win32com.client as w
    src = os.path.join(OUT, DOCNAME + ".docx")
    out = src[:-5] + ".pdf"
    app = w.Dispatch("Word.Application")
    app.Visible = False
    try:
        d = app.Documents.Open(src, ReadOnly=True)
        d.ExportAsFixedFormat(out, 17)
        d.Close(False)
    finally:
        app.Quit()
    doc = fitz.open(out)
    lines = []
    for pno in range(len(doc)):
        page = doc[pno]
        mid = page.rect.width / 2
        for b in page.get_text("dict")["blocks"]:
            for l in b.get("lines", []):
                t = "".join(s["text"] for s in l["spans"]).replace(" ", "")
                if t:
                    lines.append((pno, 0 if l["bbox"][0] < mid else 1, round(l["bbox"][1], 1), round(l["bbox"][2] - l["bbox"][0], 1), t))
    lines.sort()
    lines = [l for l in lines if l[4].strip()]
    for label, sp, text in ARMS:
        full = text.replace("SPACE", "　")
        key = full.replace("　", "")[:3]
        for i, (pno, col, y, wdt, t) in enumerate(lines):
            tt = t.replace(" ", "").replace("　", "")
            if tt.startswith(key):
                nxt = lines[i + 1][4].replace(" ", "") if i + 1 < len(lines) else ""
                unit_on_line1 = tt.endswith("字、") or (label.startswith("p4") and "とされた。" in tt)
                if label.startswith("p4"):
                    if label == "p4full":
                        # show the line that starts with るための
                        for k in range(i, min(i + 16, len(lines))):
                            l6 = lines[k][4].replace(" ", "")
                            if l6.startswith("るための"):
                                print("%-8s line6=%2d w=%6.1f %s" % (label, len(l6), lines[k][3], l6)); break
                        break
                    print("%-8s line1=%2d w=%6.1f %s" % (label, len(tt), wdt, tt))
                    break
                print("%-6s demand=%.2f line1=%2d/%2d w=%6.1f %s ...%s / %s" % (label, 0.96 + 0.5 * int(label[-1]), len(tt), len(full), wdt, "GRANT" if unit_on_line1 else "refuse", tt[-6:], nxt[:4]))
                break


if __name__ == "__main__":
    {"gen": gen, "pdf": pdf}[sys.argv[1]]()
