# -*- coding: utf-8 -*-
"""S1592 probe: per-glyph character-grid increment for PROPORTIONAL CJK faces.

Faithful slice of blind-G policies__1f014c0f (its styles, settings, fonts and the
linesAndChars / charSpace 4626 section are kept; only the body is replaced). Each
arm is one left-aligned body paragraph AND the same text in a one-cell table, so
the body and the cell law are read off the same Word PDF.

Arms: face (HGPｺﾞｼｯｸM / ＭＳ Ｐゴシック / ＭＳ Ｐ明朝) x charSpace (4626 / 0 / 9252).
Read: glyph origin steps from Word's PDF minus the face's design advance.

    python tools/metrics/_pb_propgrid_gen.py gen
    python tools/metrics/_pb_propgrid_gen.py read
"""
import copy, io, json, re, sys, zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
SRC = REPO / "pipeline_data/docx_corpus/ja/policies/1f014c0fdd5ce4e1.docx"
OUT = REPO / "tests/fixtures/propgrid"
FACES = {"hgp": "HGPｺﾞｼｯｸM", "pgo": "ＭＳ Ｐゴシック", "pmi": "ＭＳ Ｐ明朝"}
SPACES = [4626, 0, 9252]
LINES = {
    "hira": "あいうえおかきくけこさしすせそたちつてとなにぬねの",
    "kata": "アイウエオカキクケコサシスセソタチツテトナニヌネノ",
    "punc": "漢、漢。漢「漢」漢（漢）漢・漢ー漢！漢？漢：漢",
    "kanj": "漢字試験文書書式段落表組行間字間調整確認",
    "latn": "漢abcdeABCDE12345漢ＡＢＣ１２３漢",
}


def run(face, text):
    return (f'<w:r><w:rPr><w:rFonts w:ascii="{face}" w:eastAsia="{face}" w:hAnsi="{face}" '
            f'w:hint="eastAsia"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>'
            f'<w:t xml:space="preserve">{text}</w:t></w:r>')


def para(face, text):
    return f'<w:p><w:pPr><w:jc w:val="left"/></w:pPr>{run(face, text)}</w:p>'


def cell_table(face, text):
    return ('<w:tbl><w:tblPr><w:tblW w:w="9072" w:type="dxa"/><w:tblInd w:w="99" w:type="dxa"/>'
            '<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders>'
            '<w:tblCellMar><w:left w:w="99" w:type="dxa"/><w:right w:w="99" w:type="dxa"/></w:tblCellMar>'
            '</w:tblPr><w:tblGrid><w:gridCol w:w="9072"/></w:tblGrid><w:tr><w:tc><w:tcPr>'
            '<w:tcW w:w="9072" w:type="dxa"/></w:tcPr>' + para(face, text) + '</w:tc></w:tr></w:tbl>')


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    zin = zipfile.ZipFile(SRC)
    doc = zin.read("word/document.xml").decode("utf-8")
    sect = re.findall(r"<w:sectPr\b.*?</w:sectPr>", doc, re.S)[-1]
    sect = re.sub(r'<w:footerReference[^>]*/>', '', sect)
    head = doc[:doc.index("<w:body>") + len("<w:body>")]
    for fk, face in FACES.items():
        for cs in SPACES:
            body = ""
            for key, text in LINES.items():
                body += para(face, text) + cell_table(face, text) + para(face, "")
            s = re.sub(r'w:charSpace="-?\d+"', f'w:charSpace="{cs}"', sect)
            xml = head + body + s + "</w:body></w:document>"
            name = OUT / f"pg_{fk}_cs{cs}.docx"
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = xml.encode("utf-8") if item.filename == "word/document.xml" else zin.read(item.filename)
                    zout.writestr(copy.copy(item), data)
            name.write_bytes(buf.getvalue())
            print("wrote", name.name)


def read():
    import fitz
    from fontTools.ttLib import TTCollection, TTFont
    files = {"hgp": ("C:/Windows/Fonts/HGRGM.TTC", "HGPGothicM"),
             "pgo": ("C:/Windows/Fonts/msgothic.ttc", "MS PGothic"),
             "pmi": ("C:/Windows/Fonts/msmincho.ttc", "MS PMincho")}
    design = {}
    for fk, (path, name) in files.items():
        for f in TTCollection(path).fonts:
            if f["name"].getDebugName(4) == name:
                design[fk] = (f.getBestCmap(), f["hmtx"], f["head"].unitsPerEm)
    res = {}
    for pdf in sorted(OUT.glob("pg_*.pdf")):
        fk = pdf.stem.split("_")[1]
        cm, hm, upm = design[fk]
        d = fitz.open(pdf)
        rows = []
        for page in d:
            for b in page.get_text("rawdict")["blocks"]:
                for l in b.get("lines", []):
                    chars = [c for s in l["spans"] for c in s["chars"]]
                    for a, nxt in zip(chars, chars[1:]):
                        ch = a["c"]
                        if ord(ch) not in cm:
                            continue
                        adv = nxt["origin"][0] - a["origin"][0]
                        nat = hm[cm[ord(ch)]][0] * 12 / upm
                        rows.append((ch, round(nat, 3), round(adv, 3), round(adv - nat, 3), round(a["origin"][0], 2)))
        res[pdf.stem] = rows
    (OUT / "propgrid_result.json").write_text(json.dumps(res, ensure_ascii=False, indent=0), encoding="utf-8")
    for k, rows in res.items():
        by = {}
        for ch, nat, adv, add, x in rows:
            by.setdefault(ch, set()).add(add)
        print(k, " ".join(f"{c}:{sorted(v)[0]:+.2f}" for c, v in list(by.items())[:80]))


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "read": read}[sys.argv[1]]()
