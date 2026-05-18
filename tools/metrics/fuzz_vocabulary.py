"""Fuzz Quirk Discovery — OOXML attribute vocabulary.

Vocabulary defines the SPACE of attribute values to randomly combine in
fuzz documents. The goal is to expose Word vs Oxi divergences without
hand-authoring minimal repros for each suspected quirk.

Categories (Session 96):
- Paragraph spacing: before, after, beforeLines, afterLines, line, lineRule
- Paragraph indent: leftChars, rightChars, hangingChars, firstLineChars
- Auto-space flags: autoSpaceDE, autoSpaceDN, wordWrap, adjustRightInd
- Run properties: font_size, character_spacing, kern
- Cell: tcW, gridSpan, vAlign, tcMar (margin)
- Row: trHeight, hRule, cantSplit
"""
from __future__ import annotations
import random
from dataclasses import dataclass
from typing import Optional


# Paragraph-level attributes (pPr)
PARA_SPACING_BEFORE = [None, 0, 87, 120, 200]            # twip (0 = explicit zero)
PARA_SPACING_AFTER = [None, 0, 87, 120, 200]
PARA_BEFORE_LINES = [None, 0, 20, 30, 50, 100]            # 1/100 of line
PARA_AFTER_LINES = [None, 0, 20, 30, 50, 100]
PARA_LINE = [None, 240, 280, 312, 360]                    # twip
PARA_LINE_RULE = [None, "exact", "atLeast", "auto"]
PARA_LEFT_CHARS = [None, 0, 11, 100, 200, -50]
PARA_RIGHT_CHARS = [None, 0, -51, 50, 100]
PARA_HANGING_CHARS = [None, 0, 35, 70]
PARA_FIRST_LINE_CHARS = [None, 0, 50, 100, 200]
PARA_AUTO_SPACE_DE = [None, 0, 1]
PARA_AUTO_SPACE_DN = [None, 0, 1]
PARA_WORD_WRAP = [None, 0, 1]
PARA_ADJUST_RIGHT_IND = [None, 0, 1]
PARA_JC = [None, "left", "right", "center", "both", "distribute"]

# Run-level attributes (rPr)
RUN_FONT_SIZE = [None, 18, 21, 22, 24, 28]                # 1/2 of pt (sz)
RUN_CHARACTER_SPACING = [None, 0, -1, -5, -10, 5, 10]      # 1/20 pt
RUN_KERN = [None, 0, 2]                                    # half-point threshold
RUN_FONT_FAMILY_EA = [None, "ＭＳ 明朝", "ＭＳ ゴシック", "游明朝", "游ゴシック"]

# Cell-level attributes (tcPr)
CELL_WIDTH = [3000, 4500, 6000, 9343]                       # twip
CELL_GRID_SPAN = [1, 2, 3, 4]
CELL_V_ALIGN = [None, "top", "center", "bottom"]
CELL_MAR_TOP = [None, 0, 12, 50, 100]                       # twip
CELL_MAR_BOTTOM = [None, 0, 12, 50, 100]

# Row-level attributes (trPr)
ROW_TR_HEIGHT = [None, 200, 437, 658, 1000, 1500]            # twip
ROW_H_RULE = [None, "atLeast", "exact"]
ROW_CANT_SPLIT = [None, 0, 1]


@dataclass
class FuzzPara:
    """Random paragraph attributes."""
    # spacing
    before: Optional[int] = None
    after: Optional[int] = None
    before_lines: Optional[int] = None
    after_lines: Optional[int] = None
    line: Optional[int] = None
    line_rule: Optional[str] = None
    # indent
    left_chars: Optional[int] = None
    right_chars: Optional[int] = None
    hanging_chars: Optional[int] = None
    first_line_chars: Optional[int] = None
    # flags
    auto_space_de: Optional[int] = None
    auto_space_dn: Optional[int] = None
    word_wrap: Optional[int] = None
    adjust_right_ind: Optional[int] = None
    jc: Optional[str] = None
    # run
    font_size: Optional[int] = None
    character_spacing: Optional[int] = None
    kern: Optional[int] = None
    font_family_ea: Optional[str] = None
    # text
    text: str = "サンプルテキスト"

    @classmethod
    def random(cls, rng: random.Random, text: str = "サンプルテキスト") -> "FuzzPara":
        return cls(
            before=rng.choice(PARA_SPACING_BEFORE),
            after=rng.choice(PARA_SPACING_AFTER),
            before_lines=rng.choice(PARA_BEFORE_LINES),
            after_lines=rng.choice(PARA_AFTER_LINES),
            line=rng.choice(PARA_LINE),
            line_rule=rng.choice(PARA_LINE_RULE),
            left_chars=rng.choice(PARA_LEFT_CHARS),
            right_chars=rng.choice(PARA_RIGHT_CHARS),
            hanging_chars=rng.choice(PARA_HANGING_CHARS),
            first_line_chars=rng.choice(PARA_FIRST_LINE_CHARS),
            auto_space_de=rng.choice(PARA_AUTO_SPACE_DE),
            auto_space_dn=rng.choice(PARA_AUTO_SPACE_DN),
            word_wrap=rng.choice(PARA_WORD_WRAP),
            adjust_right_ind=rng.choice(PARA_ADJUST_RIGHT_IND),
            jc=rng.choice(PARA_JC),
            font_size=rng.choice(RUN_FONT_SIZE),
            character_spacing=rng.choice(RUN_CHARACTER_SPACING),
            kern=rng.choice(RUN_KERN),
            font_family_ea=rng.choice(RUN_FONT_FAMILY_EA),
            text=text,
        )

    def spacing_attrs_xml(self) -> str:
        parts = []
        if self.before is not None:
            parts.append(f'w:before="{self.before}"')
        if self.after is not None:
            parts.append(f'w:after="{self.after}"')
        if self.before_lines is not None:
            parts.append(f'w:beforeLines="{self.before_lines}"')
        if self.after_lines is not None:
            parts.append(f'w:afterLines="{self.after_lines}"')
        if self.line is not None:
            parts.append(f'w:line="{self.line}"')
        if self.line_rule is not None:
            parts.append(f'w:lineRule="{self.line_rule}"')
        if not parts:
            return ""
        return f'<w:spacing {" ".join(parts)}/>'

    def indent_attrs_xml(self) -> str:
        parts = []
        if self.left_chars is not None:
            parts.append(f'w:leftChars="{self.left_chars}"')
        if self.right_chars is not None:
            parts.append(f'w:rightChars="{self.right_chars}"')
        if self.hanging_chars is not None:
            parts.append(f'w:hangingChars="{self.hanging_chars}"')
        if self.first_line_chars is not None:
            parts.append(f'w:firstLineChars="{self.first_line_chars}"')
        if not parts:
            return ""
        return f'<w:ind {" ".join(parts)}/>'

    def pPr_xml(self) -> str:
        spacing = self.spacing_attrs_xml()
        indent = self.indent_attrs_xml()
        flags = []
        if self.auto_space_de is not None:
            flags.append(f'<w:autoSpaceDE w:val="{self.auto_space_de}"/>')
        if self.auto_space_dn is not None:
            flags.append(f'<w:autoSpaceDN w:val="{self.auto_space_dn}"/>')
        if self.word_wrap is not None:
            flags.append(f'<w:wordWrap w:val="{self.word_wrap}"/>')
        if self.adjust_right_ind is not None:
            flags.append(f'<w:adjustRightInd w:val="{self.adjust_right_ind}"/>')
        jc = f'<w:jc w:val="{self.jc}"/>' if self.jc else ""
        return f'<w:pPr>{spacing}{indent}{"".join(flags)}{jc}</w:pPr>'

    def rPr_xml(self) -> str:
        parts = []
        if self.font_family_ea is not None:
            parts.append(f'<w:rFonts w:ascii="Century" w:eastAsia="{self.font_family_ea}" w:hAnsi="Century"/>')
        if self.font_size is not None:
            parts.append(f'<w:sz w:val="{self.font_size}"/><w:szCs w:val="{self.font_size}"/>')
        if self.character_spacing is not None:
            parts.append(f'<w:spacing w:val="{self.character_spacing}"/>')
        if self.kern is not None:
            parts.append(f'<w:kern w:val="{self.kern}"/>')
        if not parts:
            return ""
        return f'<w:rPr>{"".join(parts)}</w:rPr>'

    def to_xml(self) -> str:
        return f'<w:p>{self.pPr_xml()}<w:r>{self.rPr_xml()}<w:t>{self.text}</w:t></w:r></w:p>'


@dataclass
class FuzzCell:
    width: int = 3000
    grid_span: int = 1
    v_align: Optional[str] = None
    mar_top: Optional[int] = None
    mar_bottom: Optional[int] = None
    paragraphs: list = None  # list of FuzzPara

    @classmethod
    def random(cls, rng: random.Random, n_paras: int = 1) -> "FuzzCell":
        return cls(
            width=rng.choice(CELL_WIDTH),
            grid_span=rng.choice(CELL_GRID_SPAN),
            v_align=rng.choice(CELL_V_ALIGN),
            mar_top=rng.choice(CELL_MAR_TOP),
            mar_bottom=rng.choice(CELL_MAR_BOTTOM),
            paragraphs=[FuzzPara.random(rng) for _ in range(n_paras)],
        )

    def to_xml(self) -> str:
        span_xml = f'<w:gridSpan w:val="{self.grid_span}"/>' if self.grid_span > 1 else ""
        valign_xml = f'<w:vAlign w:val="{self.v_align}"/>' if self.v_align else ""
        mar_parts = []
        if self.mar_top is not None:
            mar_parts.append(f'<w:top w:w="{self.mar_top}" w:type="dxa"/>')
        if self.mar_bottom is not None:
            mar_parts.append(f'<w:bottom w:w="{self.mar_bottom}" w:type="dxa"/>')
        mar_xml = f'<w:tcMar>{"".join(mar_parts)}</w:tcMar>' if mar_parts else ""
        paras_xml = "".join(p.to_xml() for p in self.paragraphs)
        return (f'<w:tc><w:tcPr><w:tcW w:w="{self.width}" w:type="dxa"/>'
                f'{span_xml}{valign_xml}{mar_xml}</w:tcPr>{paras_xml}</w:tc>')


@dataclass
class FuzzRow:
    tr_height: Optional[int] = None
    h_rule: Optional[str] = None
    cant_split: Optional[int] = None
    cells: list = None

    @classmethod
    def random(cls, rng: random.Random, n_cells: int = 2, n_paras_per_cell: int = 1) -> "FuzzRow":
        return cls(
            tr_height=rng.choice(ROW_TR_HEIGHT),
            h_rule=rng.choice(ROW_H_RULE),
            cant_split=rng.choice(ROW_CANT_SPLIT),
            cells=[FuzzCell.random(rng, n_paras=n_paras_per_cell) for _ in range(n_cells)],
        )

    def to_xml(self) -> str:
        trh_parts = []
        if self.tr_height is not None:
            rule = f' w:hRule="{self.h_rule}"' if self.h_rule else ""
            trh_parts.append(f'<w:trHeight w:val="{self.tr_height}"{rule}/>')
        if self.cant_split:
            trh_parts.append(f'<w:cantSplit/>')
        trPr_xml = f'<w:trPr>{"".join(trh_parts)}</w:trPr>' if trh_parts else ""
        cells_xml = "".join(c.to_xml() for c in self.cells)
        return f'<w:tr>{trPr_xml}{cells_xml}</w:tr>'
