# -*- coding: utf-8 -*-
"""S1594 probe: which East Asian face does Word give an UNSTYLED package?

Full-package controls of blind-G creative__6fd5a307 (no styles.xml, no
settings.xml, theme1.xml with an empty <a:ea> and no Jpan script, runs with
ascii/hAnsi "Hiragino Mincho Pro" only). One part changed per arm:
  A  original
  B  theme1.xml removed (part, relationship, content-type override)
  C  settings.xml added with compatibilityMode 15
  D  B + C
  E  settings.xml added with compatibilityMode 14
Read: Word PDF span font of the Japanese text and the line pitch of page 1.

    python tools/metrics/_pb_unstyled_ea_gen.py gen
    python tools/metrics/_pb_unstyled_ea_gen.py read      (after exporting PDFs)
"""
import copy, io, re, sys, zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
SRC = REPO / "pipeline_data/docx_corpus/ja/creative/6fd5a3073c03a7b9.docx"
OUT = REPO / "tests/fixtures/unstyled_ea"
SETTINGS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="{v}"/></w:compat>'
            '</w:settings>')
SET_REL = '<Relationship Id="rIdOxiSet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
SET_CT = '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'


def build(drop_theme, compat):
    zin = zipfile.ZipFile(SRC)
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            name = item.filename
            data = zin.read(name)
            if drop_theme and name == "word/theme/theme1.xml":
                continue
            if name == "word/_rels/document.xml.rels":
                s = data.decode("utf-8")
                if drop_theme:
                    s = re.sub(r'<Relationship [^>]*Target="theme/theme1.xml"[^>]*/>', "", s)
                if compat:
                    s = s.replace("</Relationships>", SET_REL + "</Relationships>")
                data = s.encode("utf-8")
            if name == "[Content_Types].xml":
                s = data.decode("utf-8")
                if drop_theme:
                    s = re.sub(r'<Override PartName="/word/theme/theme1.xml"[^>]*/>', "", s)
                if compat:
                    s = s.replace("</Types>", SET_CT + "</Types>")
                data = s.encode("utf-8")
            zout.writestr(copy.copy(item), data)
        if compat:
            zout.writestr("word/settings.xml", SETTINGS.format(v=compat))
    return buf.getvalue()


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    arms = {"A": (False, None), "B": (True, None), "C": (False, 15), "D": (True, 15), "E": (False, 14)}
    for k, (dt, cm) in arms.items():
        (OUT / f"ue_{k}.docx").write_bytes(build(dt, cm))
        print("wrote", k)


def read():
    import fitz
    from collections import Counter
    for pdf in sorted(OUT.glob("ue_*.pdf")):
        d = fitz.open(pdf)
        fonts = Counter()
        ys = []
        for b in d[0].get_text("dict")["blocks"]:
            for l in b.get("lines", []):
                for s in l["spans"]:
                    fonts[(s["font"], round(s["size"], 1))] += len(s["text"])
                if l["spans"]:
                    ys.append(round(l["spans"][0]["origin"][1], 2))
        steps = sorted(set(round(b - a, 2) for a, b in zip(ys, ys[1:]) if 0 < b - a < 40))
        print(pdf.stem, len(d), "pages", fonts.most_common(3), "steps", steps[:4])


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "read": read}[sys.argv[1]]()
