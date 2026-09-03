# -*- coding: utf-8 -*-
"""Which documents actually ASK the theme for their East Asian font, and what
does Word answer?

S323 suppresses the theme's `<a:font script="Jpan">` entry whenever the minor
font's `<a:ea typeface=""/>` is explicitly empty, on the reading that "an empty
EA slot means fall through to rPrDefault". Its two evidence documents inherit a
LITERAL `w:eastAsia="ＭＳ 明朝"` from docDefaults, so neither of them ever
reaches the theme at all -- the reading is untestable there.

This separates the population that can answer the question: docDefaults (or the
style chain) hands out `eastAsiaTheme`, the theme's `<a:ea>` is empty, and a
Jpan entry names some face. For each such document it then reads the Word truth
PDF's embedded font names and says whether the Jpan face is among them.

    python _theme_ea_census.py
"""
import os
import re
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _theme_ea_reach as REACH

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
REPO = Path(__file__).resolve().parents[2]

PDF_DIRS = [
    REPO / "pipeline_data" / "ja_benchmark" / "ssim_blind50" / "word_pdf",
    REPO / "pipeline_data" / "ja_benchmark" / "ssim_blindB50" / "word_pdf",
    REPO / "pipeline_data" / "ja_benchmark" / "ssim_blindC50" / "word_pdf",
]
# Family -> the name a PDF subset uses.
PDF_ALIAS = {"游明朝": "YuMincho", "游ゴシック": "YuGothic", "游ゴシック Light": "YuGothic",
             "ＭＳ 明朝": "MS-Mincho", "ＭＳ ゴシック": "MS-Gothic", "メイリオ": "Meiryo"}


def theme_of(z):
    names = [n for n in z.namelist() if n.startswith("word/theme/")]
    if not names:
        return None
    t = z.read(names[0]).decode("utf8", "replace")
    m = re.search(r"<a:minorFont>(.*?)</a:minorFont>", t, re.S)
    if not m:
        return None
    body = m.group(1)
    ea = re.search(r'<a:ea[^>]*typeface="([^"]*)"', body)
    jp = re.search(r'<a:font script="Jpan" typeface="([^"]*)"', body)
    return ("" if (ea and not ea.group(1)) else (ea.group(1) if ea else None),
            jp.group(1) if jp else None)


def asks_theme(z):
    """How many CJK characters actually resolve THROUGH the theme.

    Reading docDefaults alone is not enough (the style chain and the run both
    sit in between), and reading `<w:t>` with a regex is not enough either --
    see `_theme_ea_reach` for the two nesting traps that invent evidence.
    """
    styles, default_style, docdef = REACH.styles_of(z)
    doc = z.read("word/document.xml").decode("utf8", "replace")
    n_theme, n_lit, _ = REACH.scan(doc, styles, default_style, docdef)
    return n_theme > 0, n_theme


def pdf_fonts(doc_id):
    import fitz
    for d in PDF_DIRS:
        p = d / f"{doc_id}.pdf"
        if p.exists():
            f = fitz.open(p)
            seen = set()
            for i in range(f.page_count):
                for ent in f[i].get_fonts(full=False):
                    seen.add(re.sub(r"^[A-Z]{6}\+", "", ent[3]))
            return sorted(seen)
    return None


def main():
    roots = [REPO / "pipeline_data" / "docx_corpus" / "ja",
             REPO / "tools" / "golden-test" / "documents" / "docx"]
    rows = []
    for root in roots:
        for p in sorted(root.rglob("*.docx")):
            try:
                z = zipfile.ZipFile(p)
                th = theme_of(z)
                if not th:
                    continue
                ea, jp = th
                if ea != "" or not jp:
                    continue           # only the explicit-empty-ea + Jpan shape
                via_default, n_runs = asks_theme(z)
                if not via_default and n_runs <= 0:
                    continue           # nothing ever reaches the theme
                did = f"{p.parent.name}__{p.stem}" if root.name == "ja" else p.stem
                rows.append((did, jp, via_default, n_runs, pdf_fonts(did)))
            except Exception as e:
                print(f"  !! {p.name}: {str(e)[:60]}")
    print(f"documents whose EA font can only come from the theme: {len(rows)}\n")
    print(f"{'doc':<44} {'Jpan':<14} {'reach':<6} {'cjk':>6}  PDF says")
    for did, jp, via, n, fonts in rows:
        want = PDF_ALIAS.get(jp, jp)
        if fonts is None:
            verdict = "(no truth PDF)"
        else:
            hit = [f for f in fonts if want.lower() in f.lower()]
            verdict = ("USES " + ", ".join(hit)) if hit else ("NOT among: " + ", ".join(fonts)[:70])
        print(f"{did:<44} {jp:<14} {str(via):<6} {n:>6}  {verdict}")


if __name__ == "__main__":
    main()
