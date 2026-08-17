# -*- coding: utf-8 -*-
"""How much does Word actually compress 約物 on an overflowing CJK line?

The char-budget wall is now localised to one line of layout/mod.rs: the render
pass rebuilds each standalone 、。 width as (natural - comp) and overwrites the
breaker's own (more compressed) width, so a line grows in proportion to its 約物
count (r = 0.873; c7b923e5's worst line goes 428.60 -> 479.50). Before touching
it, settle which side is Word: on that line Word's effective advance is 9.73pt
per character, Oxi's breaker assumes 7.90 and its emitter draws 9.05.

Read Word's per-character advances directly (PDF rawdict gives per-glyph boxes)
on lines that differ ONLY in how many 約物 they carry.

    python _pb_yakucomp_gen.py gen
    python _pb_yakucomp_gen.py pdf      # Word truth
    python _pb_yakucomp_gen.py oxi      # Oxi, same arms
"""
import json
import os
import re
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_yakucomp")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")

sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, DRELS, NS, RELS  # noqa: E402

FACE = "ＭＳ 明朝"
SZ_HP = 21                 # 10.5pt
COMPAT = os.environ.get("OXI_PB_COMPAT", "15")
YAK = set("。、）」")

# Each arm is one paragraph whose text is a fixed CJK filler with N inserted
# 、。 pairs. The text is long enough to wrap, so the first line is the one that
# has to absorb the compression.
BASE = "本規程は労働者の就業に関する事項を定めるものであり関係者はこれを遵守しなければならない"


def arm_text(n_yak):
    out = []
    for i, ch in enumerate(BASE * 3):
        out.append(ch)
        if n_yak and i % max(1, (len(BASE) * 3) // n_yak) == 0 and i:
            out.append("、" if (i // 7) % 2 else "。")
    return "".join(out)


# ★v1 learned only that Word does NOT compress when nothing demands it (all arms
# 40 chars / 10.51pt, 約物 included, Oxi identical). The content width is 425.2pt
# and 40 chars are 420.0, so no arm ever crossed the "one more character only
# fits if something squeezes" boundary. v2 puts a 、 exactly at the boundary and
# walks it across, with a variable number of EARLIER 約物 to share the squeeze.
DEMAND = os.environ.get("OXI_PB_DEMAND")
FILL = "本規程労働者就業関事項定関係者遵守義務履行責任範囲明確化目的作成周知徹底図"


def demand_text(pos, n_prior):
    """`pos` filler chars (with n_prior 約物 mixed in) then the trigger 、."""
    body = []
    step = max(2, pos // (n_prior + 1)) if n_prior else 0
    for i in range(pos):
        body.append(FILL[i % len(FILL)])
        if n_prior and step and (i + 1) % step == 0 and                 sum(1 for c in body if c in "、。") < n_prior:
            body.append("、")
    return "".join(body[:pos]) + "、" + (FILL * 3)


if DEMAND:
    ARMS = [(p, k) for k in (0, 2, 4) for p in (38, 39, 40, 41)]
else:
    ARMS = [0, 2, 4, 8, 16]


def docx():
    return os.path.join(OUT, "yakucomp.docx")


def para(text, ppr=""):
    return ('<w:p><w:pPr>%s<w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s"/>'
            '<w:sz w:val="%d"/></w:rPr></w:pPr><w:r><w:rPr>'
            '<w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s"/>'
            '<w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr>'
            '<w:t xml:space="preserve">%s</w:t></w:r></w:p>'
            % (ppr, FACE, FACE, FACE, SZ_HP, FACE, FACE, FACE, SZ_HP, SZ_HP, text))


def gen():
    os.makedirs(OUT, exist_ok=True)
    body = []
    for ai, n in enumerate(ARMS):
        body.append(para("A%02dZ" % ai, "<w:pageBreakBefore/>" if ai else ""))
        body.append(para(demand_text(*n) if DEMAND else arm_text(n)))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS +
           "><w:body>" + "".join(body) +
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838" w:code="9"/>'
           '<w:pgMar w:top="1985" w:right="1701" w:bottom="1701" w:left="1701" '
           'w:header="851" w:footer="992" w:gutter="0"/>'
           '<w:docGrid w:type="lines" w:linePitch="360"/>'
           "</w:sectPr></w:body></w:document>")
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + ">"
              "<w:docDefaults><w:rPrDefault><w:rPr>"
              '<w:rFonts w:ascii="%s" w:eastAsia="%s" w:hAnsi="%s"/>'
              "</w:rPr></w:rPrDefault></w:docDefaults>"
              '<w:style w:type="paragraph" w:default="1" w:styleId="a">'
              '<w:name w:val="Normal"/><w:rPr><w:sz w:val="%d"/></w:rPr></w:style>'
              "</w:styles>" % (FACE, FACE, FACE, SZ_HP))
    # ★c7b923e5 (whose lines ARE flush to 0.11pt with no w:jc anywhere) declares
    # <w:useFELayout/>; this probe did not, and Word left every arm ragged and
    # uncompressed. OXI_PB_FE=1 adds it — the suspected discriminator.
    fe = "<w:useFELayout/>" if os.environ.get("OXI_PB_FE") else ""
    settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + NS +
                '>' + fe + '<w:compat><w:compatSetting w:name="compatibilityMode"'
                ' w:uri="http://schemas.microsoft.com/office/word"'
                ' w:val="%s"/></w:compat></w:settings>' % COMPAT)
    ct = CT.replace("</Types>",
                    '<Override PartName="/word/settings.xml" ContentType="application/'
                    'vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
                    "</Types>")
    drels = DRELS.replace("</Relationships>",
                          '<Relationship Id="rIdSet" Type="http://schemas.openxmlformats.org/'
                          'officeDocument/2006/relationships/settings" Target="settings.xml"/>'
                          "</Relationships>")
    with zipfile.ZipFile(docx(), "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", drels)
        z.writestr("word/styles.xml", styles)
        z.writestr("word/settings.xml", settings)
        z.writestr("word/document.xml", doc)
    print("wrote", docx(), len(ARMS), "arms; compat", COMPAT)


def report(rows, who):
    print("== %s ==" % who)
    print("%-10s %-6s %-7s %-9s %-9s %-9s %s"
          % ("arm", "nyak", "nchars", "line_w", "per_char", "yak_adv", "cjk_adv"))
    for n, r in rows:
        lbl = ("p%d_k%d" % n) if isinstance(n, tuple) else str(n)
        if not r:
            print("%-10s MISSING" % lbl)
            continue
        print("%-10s %-6d %-7d %-9.2f %-9.3f %-9s %s"
              % (lbl, r["nyak"], r["nch"], r["w"], r["w"] / max(1, r["nch"]),
                 "%.2f" % r["yak"] if r["yak"] else "-",
                 "%.2f" % r["cjk"] if r["cjk"] else "-"))


def pdf():
    import fitz
    import win32com.client as w
    out = docx().replace(".docx", ".pdf")
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    d = app.Documents.Open(docx(), ReadOnly=True)
    try:
        d.ExportAsFixedFormat(out, 17)
    finally:
        d.Close(False)
        app.Quit()
    doc = fitz.open(out)
    rows = []
    for ai, n in enumerate(ARMS):
        r = None
        for pi in range(doc.page_count):
            txt = doc[pi].get_text()
            if "A%02dZ" % ai not in txt:
                continue
            chars = []
            for b in doc[pi].get_text("rawdict")["blocks"]:
                for ln in b.get("lines", []):
                    for s in ln.get("spans", []):
                        for c in s.get("chars", []):
                            chars.append((round(ln["bbox"][1], 1), c["bbox"][0],
                                          c["bbox"][2], c["c"]))
            ys = sorted({y for y, _, _, _ in chars})
            if len(ys) < 2:
                break
            y = ys[1]                      # the paragraph's FIRST wrapped line
            line = sorted([c for c in chars if c[0] == y], key=lambda c: c[1])
            if len(line) < 3:
                break
            advs = [line[i + 1][1] - line[i][1] for i in range(len(line) - 1)]
            yk = [a for a, c in zip(advs, line) if c[3] in YAK]
            cj = [a for a, c in zip(advs, line) if c[3] not in YAK]
            r = {"nch": len(line), "w": line[-1][2] - line[0][1],
                 "nyak": sum(1 for c in line if c[3] in YAK),
                 "yak": sum(yk) / len(yk) if yk else 0.0,
                 "cjk": sum(cj) / len(cj) if cj else 0.0}
            break
        rows.append((n, r))
    report(rows, "WORD")


def oxi(envs=""):
    env = dict(os.environ)
    for kv in [s for s in envs.split(",") if s]:
        k, _, v = kv.partition("=")
        env[k] = v or "1"
    out = os.path.join(tempfile.gettempdir(), "yakucomp_oxi.json")
    subprocess.run([GDI, docx(), os.path.join(tempfile.gettempdir(), "yc"),
                    "--dump-layout=" + out], check=True, capture_output=True, env=env)
    pages = json.load(open(out, encoding="utf-8"))["pages"]
    rows = []
    for ai, n in enumerate(ARMS):
        r = None
        for pg in pages:
            ts = [e for e in pg["elements"] if e["type"] == "text"]
            if not any("A%02dZ" % ai in (e.get("text") or "") for e in ts):
                continue
            ys = sorted({round(e["y"], 1) for e in ts})
            if len(ys) < 2:
                break
            line = sorted([e for e in ts if round(e["y"], 1) == ys[1]],
                          key=lambda e: e["x"])
            if len(line) < 3:
                break
            yk = [e.get("w") or 0 for e in line if (e.get("text") or "") in YAK]
            cj = [e.get("w") or 0 for e in line if (e.get("text") or "") not in YAK]
            r = {"nch": len(line),
                 "w": line[-1]["x"] + (line[-1].get("w") or 0) - line[0]["x"],
                 "nyak": len(yk),
                 "yak": sum(yk) / len(yk) if yk else 0.0,
                 "cjk": sum(cj) / len(cj) if cj else 0.0}
            break
        rows.append((n, r))
    report(rows, "OXI " + (envs or "(default)"))


if __name__ == "__main__":
    if sys.argv[1] == "oxi":
        oxi(sys.argv[2] if len(sys.argv) > 2 else "")
    elif sys.argv[1] == "pdf":
        pdf()
    else:
        gen()
