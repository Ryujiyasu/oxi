# -*- coding: utf-8 -*-
"""How many characters does one line hold — Word's answer against this engine's.

Half the Phase-1 failures start on an ordinary paragraph, which means no
construct is to blame: some earlier paragraph simply took one line too few or
too many, and the page slip is the error surfacing later. Two JA documents show
the direction — Word gives a paragraph 4 lines where this engine gives 3 — so
the engine is fitting one character too many.

Rather than read a blind document's line breaks, this asks the question
directly: a paragraph of exactly K characters, K swept, and the K at which each
side goes from one line to two IS its capacity. A one-character gap between the
two answers is the whole bug, visible in one number.

    python tools/metrics/_pb_line_capacity.py
    LC_GRID=0 LC_JC=left python tools/metrics/_pb_line_capacity.py

Sweeps the settings that could move the capacity — grid on/off, justified or
not, the indent, and what character sits at the break — because a capacity that
is right in one of those and wrong in another names the rule.
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
GDI = REPO / "tools" / "oxi-gdi-renderer" / "target" / "release" / "oxi-gdi-renderer.exe"
OUT = REPO / "tests" / "fixtures" / "line_capacity"

FONT = "ＭＳ 明朝"
FILLER = "あいうえおかきくけこさしすせそたちつてとなにぬねのはひふへほまみむめも"

CONTENT_TYPES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                 '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                 '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                 '<Default Extension="xml" ContentType="application/xml"/>'
                 '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
                 '</Types>')
ROOT_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
             '</Relationships>')


def filler(n: int, tail: str = "") -> str:
    """`n` characters of kana, optionally ending on a given character.

    `LC_PUNCT=k` sprinkles a comma every k characters. Punctuation is what the
    two engines can disagree about: a line-end comma is compressible, and how
    much of it each side is willing to squeeze decides whether one more
    character fits.
    """
    every = int(os.environ.get("LC_PUNCT", "0"))
    pool = FILLER
    # `LC_MIX=k` puts a short Latin token every k characters. A CJK line that
    # carries Latin is where the two engines were seen to disagree, and the
    # Latin advance is the half of the line neither side measures in ems.
    mix = int(os.environ.get("LC_MIX", "0"))
    if mix > 0:
        token = os.environ.get("LC_TOKEN", "AB")
        pool = "".join((token if (i + 1) % mix == 0 else c)
                       for i, c in enumerate(FILLER * 4))
    if every > 0:
        pool = "".join(("、" if (i + 1) % every == 0 else c)
                       for i, c in enumerate(FILLER * 4))
    body = (pool * (n // len(pool) + 1))[:max(0, n - len(tail))]
    return body + tail


def document(counts, grid: bool, jc: str, indent_chars: int, size: float, tail: str) -> str:
    rpr = (f'<w:rPr><w:rFonts w:ascii="{FONT}" w:eastAsia="{FONT}" w:hAnsi="{FONT}"/>'
           f'<w:sz w:val="{round(size * 2)}"/><w:szCs w:val="{round(size * 2)}"/></w:rPr>')
    # `leftChars` is in HUNDREDTHS of a character, and it OVERRIDES `left`.
    # Writing leftChars="4" for a four-character indent asks for 0.04 of one —
    # 0.42pt — and silently kills the twips beside it. Two sweeps reported
    # "an indent changes nothing" on the strength of that.
    # The four documents that wrap a one-line paragraph in two all carry an
    # indent, and none of them is a plain LEFT one: two set `rightChars`, one
    # sets `hangingChars`, one sets a NEGATIVE `leftChars`. So the sweep has to
    # reach those, not just the left edge.
    right = int(os.environ.get("LC_RIGHT", "0"))
    hang = int(os.environ.get("LC_HANG", "0"))
    # Real documents state leftChars and left that DISAGREE: one carries
    # leftChars="800" (eight characters) beside left="1885" twips (94.25pt),
    # which is 11.78pt per character against a 10.5pt body. Which one each side
    # follows is the whole question, and a sweep that keeps them consistent —
    # as the first ones did — can never ask it. `LC_TWIP_PER_CHAR` sets the
    # twips half of the pair independently.
    per_char = float(os.environ.get("LC_TWIP_PER_CHAR", "0")) or size * 20
    bits = []
    if indent_chars:
        bits.append(f'w:leftChars="{indent_chars * 100}" '
                    f'w:left="{round(indent_chars * per_char)}"')
    if right:
        bits.append(f'w:rightChars="{right * 100}" '
                    f'w:right="{round(right * per_char)}"')
    if hang:
        bits.append(f'w:hangingChars="{hang * 100}" '
                    f'w:hanging="{round(hang * size * 20)}"')
    ind = f'<w:ind {" ".join(bits)}/>' if bits else ""
    paras = []
    for n in counts:
        # The order of pPr's children is fixed by the schema — spacing, then
        # ind, then jc — and Word SILENTLY DROPS an element that arrives out of
        # turn. The first sweep put jc first and spent three runs reporting
        # that an indent changes nothing, because neither side ever saw it.
        # Two of the four documents pair their indent with a 40pt EXACT line.
        # An exact rule is its own placement regime, so it belongs in the sweep.
        line = os.environ.get("LC_LINE", "240")
        rule = os.environ.get("LC_RULE", "auto")
        ppr = ('<w:pPr>'
               f'<w:spacing w:before="0" w:after="0" w:line="{line}" w:lineRule="{rule}"/>'
               f'{ind}<w:jc w:val="{jc}"/>{rpr}</w:pPr>')
        paras.append(f'<w:p>{ppr}<w:r>{rpr}<w:t>{filler(n, tail)}</w:t></w:r></w:p>')
    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" '
            'w:header="0" w:footer="0" w:gutter="0"/>'
            + ('<w:docGrid w:type="lines" w:linePitch="360"/>' if grid else '')
            + '</w:sectPr>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'<w:body>{"".join(paras)}{sect}</w:body></w:document>')


def write(path: Path, xml: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("word/document.xml", xml)


def oxi_line_counts(path: Path, n_paras: int) -> list:
    with tempfile.TemporaryDirectory() as tmp:
        dump = Path(tmp) / "l.json"
        subprocess.run([str(GDI), str(path), str(Path(tmp) / "p"), "150",
                        f"--dump-layout={dump}"], capture_output=True)
        if not dump.is_file():
            return [None] * n_paras
        data = json.loads(dump.read_text(encoding="utf-8"))
    seen: dict[int, set] = {}
    for page in data.get("pages", []):
        for e in page.get("elements", []):
            if e.get("type") == "text" and e.get("text") and e.get("para_idx") is not None:
                seen.setdefault(e["para_idx"], set()).add(round(float(e["y"]), 1))
    return [len(seen.get(i, ())) for i in range(n_paras)]


def word_line_counts(app, path: Path, n_paras: int) -> list:
    doc = app.Documents.Open(str(path), False, True)
    try:
        out = []
        for para in doc.Paragraphs:
            try:
                out.append(int(para.Range.ComputeStatistics(1)))
            except Exception:  # noqa: BLE001
                out.append(None)
        return (out + [None] * n_paras)[:n_paras]
    finally:
        doc.Close(False)


def main() -> int:
    grid = os.environ.get("LC_GRID", "1") == "1"
    jc = os.environ.get("LC_JC", "both")
    indent = int(os.environ.get("LC_INDENT", "0"))
    size = float(os.environ.get("LC_SIZE", "10.5"))
    tail = os.environ.get("LC_TAIL", "")
    lo = int(os.environ.get("LC_LO", "30"))
    hi = int(os.environ.get("LC_HI", "48"))
    counts = list(range(lo, hi + 1))

    at = OUT / f"cap_g{int(grid)}_{jc}_i{indent}_s{size:g}_{tail or 'plain'}.docx"
    write(at, document(counts, grid, jc, indent, size, tail))

    import win32com.client
    app = win32com.client.DispatchEx("Word.Application")
    for attr, value in (("Visible", False), ("DisplayAlerts", False)):
        try:
            setattr(app, attr, value)
        except Exception:  # noqa: BLE001
            pass
    try:
        wl = word_line_counts(app, at, len(counts))
    finally:
        app.Quit()
    ol = oxi_line_counts(at, len(counts))

    print(f"grid={int(grid)} jc={jc} indent={indent} size={size:g} "
          f"tail={tail or '(none)'}  font={FONT}")
    print(f"{'chars':>6} {'word':>5} {'oxi':>5}")
    w_break = o_break = None
    for n, w, o in zip(counts, wl, ol):
        mark = "" if w == o else "   <<<"
        print(f"{n:6} {str(w):>5} {str(o):>5}{mark}")
        if w_break is None and w and w > 1:
            w_break = n
        if o_break is None and o and o > 1:
            o_break = n
    print(f"\nfirst two-line at: word {w_break}, oxi {o_break}"
          f"  -> capacity word {(w_break or 0) - 1}, oxi {(o_break or 0) - 1}")

    # Every transition, not just the first. The disagreement on real documents
    # was on a FOUR-line paragraph, and a per-line capacity that is right on
    # line 1 can still be wrong on line 3 — a first-line indent, a hanging
    # indent or per-line rounding all show up only later.
    def steps(series):
        out, prev = {}, None
        for k, v in zip(counts, series):
            if v and prev is not None and v > prev:
                out[v] = k
            if v:
                prev = v
        return out

    ws, os_ = steps(wl), steps(ol)
    rows = sorted(set(ws) | set(os_))
    if rows:
        print("\nthe K at which each line first appears:")
        print(f"{'line':>5} {'word':>6} {'oxi':>6}")
        for k in rows:
            w, o = ws.get(k), os_.get(k)
            print(f"{k:5} {str(w):>6} {str(o):>6}{'' if w == o else '   <<<'}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    raise SystemExit(main())
