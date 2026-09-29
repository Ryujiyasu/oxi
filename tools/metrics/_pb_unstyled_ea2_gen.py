# -*- coding: utf-8 -*-
"""S1594 probe 2: the East Asian default of an UNSTYLED package on other docs.

For each source doc: styles.xml removed (part, relationship, content type) and
every run-level eastAsia / eastAsiaTheme attribute stripped, then
  e  theme minor/major <a:ea> emptied and their Jpan script entries removed
  j  theme minor Jpan script set to "ＭＳ Ｐゴシック" (ea emptied)
  n  theme1.xml removed as well
Read: Word PDF span fonts of page 1.

    python tools/metrics/_pb_unstyled_ea2_gen.py gen
    python tools/metrics/_pb_unstyled_ea2_gen.py read
"""
import copy, io, re, sys, zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
SRCS = {"pol": "ja/policies/1f014c0fdd5ce4e1", "leg": "ja/legal/0adfa2505c48d1e8",
        "rep": "ja/reports/5823d5a85b3b39f1"}
OUT = REPO / "tests/fixtures/unstyled_ea2"


def theme_edit(t, jpan):
    def fix(block):
        block = re.sub(r'<a:ea typeface="[^"]*"', '<a:ea typeface=""', block)
        block = re.sub(r'<a:font script="Jpan" typeface="[^"]*"/>', "", block)
        if jpan:
            block = block.replace("</a:minorFont>", f'<a:font script="Jpan" typeface="{jpan}"/></a:minorFont>')
        return block
    t = re.sub(r"<a:minorFont>.*?</a:minorFont>", lambda m: fix(m.group(0)), t, flags=re.S)
    t = re.sub(r"<a:majorFont>.*?</a:majorFont>", lambda m: re.sub(r'<a:font script="Jpan" typeface="[^"]*"/>', "",
               re.sub(r'<a:ea typeface="[^"]*"', '<a:ea typeface=""', m.group(0))), t, flags=re.S)
    return t


def build(src, arm):
    zin = zipfile.ZipFile(REPO / "pipeline_data/docx_corpus" / f"{src}.docx")
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            name = item.filename
            data = zin.read(name)
            if (name == "word/styles.xml" and arm != "s") or (arm == "n" and name == "word/theme/theme1.xml"):
                continue
            if name == "word/styles.xml" and arm == "s":
                st = data.decode("utf-8")
                st = re.sub(r' w:eastAsia(Theme)?="[^"]*"', "", st)
                dd = re.search(r"<w:rPrDefault>.*?</w:rPrDefault>", st, re.S)
                blk = dd.group(0)
                if "<w:rFonts" in blk:
                    blk2 = blk.replace("<w:rFonts", '<w:rFonts w:eastAsia="Times New Roman"', 1)
                elif "<w:rPr>" in blk:
                    blk2 = blk.replace("<w:rPr>", '<w:rPr><w:rFonts w:eastAsia="Times New Roman"/>', 1)
                else:
                    blk2 = '<w:rPrDefault><w:rPr><w:rFonts w:eastAsia="Times New Roman"/></w:rPr></w:rPrDefault>'
                st = st[:dd.start()] + blk2 + st[dd.end():]
                _unused = re.sub(r"(<w:docDefaults>.*?<w:rPrDefault>\s*<w:rPr>)", r'<w:rFonts w:eastAsia="Times New Roman"/>', st, count=1, flags=re.S)
                data = st.encode("utf-8")
            s = None
            if name == "word/document.xml":
                s = data.decode("utf-8")
                s = re.sub(r' w:eastAsia(Theme)?="[^"]*"', "", s)
            elif name == "word/_rels/document.xml.rels":
                s = data.decode("utf-8")
                if arm != "s":
                    s = re.sub(r'<Relationship [^>]*Target="styles.xml"[^>]*/>', "", s)
                if arm == "n":
                    s = re.sub(r'<Relationship [^>]*Target="theme/theme1.xml"[^>]*/>', "", s)
            elif name == "[Content_Types].xml":
                s = data.decode("utf-8")
                if arm != "s":
                    s = re.sub(r'<Override PartName="/word/styles.xml"[^>]*/>', "", s)
                if arm == "n":
                    s = re.sub(r'<Override PartName="/word/theme/theme1.xml"[^>]*/>', "", s)
            elif name == "word/theme/theme1.xml":
                s = theme_edit(data.decode("utf-8"), "ＭＳ Ｐゴシック" if arm == "j" else None)
            if s is not None:
                data = s.encode("utf-8")
            zout.writestr(copy.copy(item), data)
    return buf.getvalue()


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    for k, src in SRCS.items():
        for arm in ("e", "j", "n", "s"):
            (OUT / f"u2_{k}_{arm}.docx").write_bytes(build(src, arm))
    print("ok")


def read():
    import fitz
    from collections import Counter
    for pdf in sorted(OUT.glob("u2_*.pdf")):
        d = fitz.open(pdf)
        fonts = Counter()
        for b in d[0].get_text("dict")["blocks"]:
            for l in b.get("lines", []):
                for s in l["spans"]:
                    if any(ord(c) > 0x3000 for c in s["text"]):
                        fonts[s["font"]] += len(s["text"])
        print(pdf.stem, fonts.most_common(3))


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "read": read}[sys.argv[1]]()
