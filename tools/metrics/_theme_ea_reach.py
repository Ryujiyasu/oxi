# -*- coding: utf-8 -*-
"""Does the East Asian THEME font actually reach any visible CJK text?

The coarse census asks "does docDefaults hand out `eastAsiaTheme`", which is not
the question: the default paragraph style, the paragraph's own style chain, the
character style and the run's own rFonts each sit between docDefaults and a
glyph, and any literal `w:eastAsia` along the way ends the search. A document
where every one of those names a face has NO OPINION about what an empty
`<a:ea>` means -- counting it as a counterexample is counting silence as a "no".

★Two nesting traps, both of which produced phantom evidence on the first pass:
 - A run that carries a drawing WRAPS the textbox's own runs. Reading `<w:t>`
   out of the outer run credits the inner text to the anchor's rPr (which is
   usually just `<w:noProof/>`, i.e. "inherits everything") and invents a
   theme-resolved run that does not exist. So this walks TOKENS and keeps a
   stack: text belongs to the innermost open run.
 - `<mc:Fallback>` repeats the same text as `<mc:Choice>` in VML. Word paints
   the Choice; counting both double-counts, and the copies can disagree.

    python _theme_ea_reach.py <docx> [docx ...]
"""
import re
import sys
import zipfile
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
CJK = re.compile(r"[぀-ヿ㐀-鿿]")
TOKEN = re.compile(
    r"<(/?)(w:p|w:r|mc:Fallback|w:pPr|w:rPr|w:txbxContent)([ >])|"
    r"<w:t[^>]*>([^<]*)</w:t>|"
    r"<w:pStyle w:val=\"([^\"]*)\"|<w:rStyle w:val=\"([^\"]*)\"|"
    r"(<w:rFonts[^>]*>)")


def rfonts_ea(frag):
    """(literal, uses_theme) from the first <w:rFonts> in a fragment."""
    if not frag:
        return None, False
    m = re.search(r"<w:rFonts[^>]*>", frag)
    if not m:
        return None, False
    tag = m.group(0)
    lit = re.search(r'w:eastAsia="([^"]+)"', tag)
    return (lit.group(1) if lit else None), ("eastAsiaTheme" in tag)


def styles_of(z):
    st = z.read("word/styles.xml").decode("utf8", "replace") if "word/styles.xml" in z.namelist() else ""
    out, default = {}, None
    for m in re.finditer(r"<w:style [^>]*>.*?</w:style>", st, re.S):
        s = m.group(0)
        sid = re.search(r'w:styleId="([^"]*)"', s)
        if not sid:
            continue
        based = re.search(r'<w:basedOn w:val="([^"]*)"', s)
        rpr = re.search(r"<w:rPr>.*?</w:rPr>", s, re.S)
        out[sid.group(1)] = (based.group(1) if based else None, rpr.group(0) if rpr else None)
        if 'w:default="1"' in s and 'w:type="paragraph"' in s:
            default = sid.group(1)
    dd = re.search(r"<w:docDefaults>.*?</w:docDefaults>", st, re.S)
    return out, default, (dd.group(0) if dd else None)


def resolve(chain_ids, styles):
    """Walk a style chain outward; the first literal or theme marker wins."""
    seen = set()
    for sid in chain_ids:
        while sid and sid in styles and sid not in seen:
            seen.add(sid)
            based, rpr = styles[sid]
            lit, theme = rfonts_ea(rpr)
            if lit:
                return lit, False
            if theme:
                return None, True
            sid = based
    return None, False


def scan(doc, styles, default_style, docdef, show=False):
    dd_lit, dd_theme = rfonts_ea(docdef)
    paras, runs = [], []          # stacks of {pstyle} / {rstyle, rfonts, in_rpr}
    in_fallback = 0
    n_theme = n_lit = 0
    faces = {}
    for m in TOKEN.finditer(doc):
        close, name, _, text, pstyle, rstyle, rfonts = m.groups()
        if name == "mc:Fallback":
            in_fallback += -1 if close else 1
            continue
        if name == "w:p":
            if close:
                paras and paras.pop()
            elif not doc[m.start():m.start() + 8].startswith("<w:pPr"):
                paras.append({"pstyle": None})
            continue
        if name == "w:r":
            if close:
                runs and runs.pop()
            else:
                runs.append({"rstyle": None, "rfonts": None, "in_rpr": False})
            continue
        if name == "w:rPr" and runs:
            runs[-1]["in_rpr"] = not close
            continue
        if pstyle is not None and paras:
            paras[-1]["pstyle"] = pstyle
            continue
        if rstyle is not None and runs:
            runs[-1]["rstyle"] = rstyle
            continue
        if rfonts is not None and runs and runs[-1]["in_rpr"] and runs[-1]["rfonts"] is None:
            runs[-1]["rfonts"] = rfonts
            continue
        if text is None or in_fallback or not runs:
            continue
        n = len(CJK.findall(text))
        if not n:
            continue
        run = runs[-1]
        lit, theme = rfonts_ea(run["rfonts"])
        if not lit and not theme:
            chain = ([run["rstyle"]] if run["rstyle"] else [])
            chain += [paras[-1]["pstyle"]] if (paras and paras[-1]["pstyle"]) else \
                     ([default_style] if default_style else [])
            lit, theme = resolve(chain, styles)
        if not lit and not theme:
            lit, theme = dd_lit, dd_theme
        if theme:
            n_theme += n
            if show:
                print(f"    theme run: {text[:30]!r}")
        elif lit:
            n_lit += n
            faces[lit] = faces.get(lit, 0) + n
    return n_theme, n_lit, faces


def main():
    show = "--show" in sys.argv
    for path in [a for a in sys.argv[1:] if a != "--show"]:
        z = zipfile.ZipFile(path)
        styles, default_style, docdef = styles_of(z)
        doc = z.read("word/document.xml").decode("utf8", "replace")
        n_theme, n_lit, faces = scan(doc, styles, default_style, docdef, show)
        top = sorted(faces.items(), key=lambda kv: -kv[1])[:3]
        verdict = "THEME REACHES TEXT" if n_theme else "no opinion (every run names a face)"
        print(f"{Path(path).stem:<26} cjk via theme={n_theme:>6}  via literal={n_lit:>6}  "
              f"{', '.join(f'{k}:{v}' for k, v in top):<34}  {verdict}")


if __name__ == "__main__":
    main()
