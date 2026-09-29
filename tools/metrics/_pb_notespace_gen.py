# -*- coding: utf-8 -*-
"""S1603 probe: the gap between consecutive endnotes.

Full-package controls of blind-G EN forms__005d851e (compat 15; every endnote
paragraph carries spacing before=120 after=120 and Word steps note 1 -> note 2
by one line + 6.0, not + 12). One edit per arm, applied to every endnote
paragraph's own <w:spacing w:before="120" w:after="120"/>:
  A  before=240 after=60     (12 / 3)
  B  before=60  after=240    (3 / 12)
  C  before=0   after=240
  D  before=240 after=0
Read: Word PDF, page of the endnotes, the origin y of each note number (7pt).

    python tools/metrics/_pb_notespace_gen.py gen
    python tools/metrics/_pb_notespace_gen.py read   (after exporting PDFs)
"""
import copy, io, sys, zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
SRC = REPO / "pipeline_data/docx_corpus/en/forms/005d851e7948bcf4.docx"
OUT = REPO / "tests/fixtures/notespace"
ARMS = {"A": (240, 60), "B": (60, 240), "C": (0, 240), "D": (240, 0)}


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    zin = zipfile.ZipFile(SRC)
    en = zin.read("word/endnotes.xml").decode("utf-8")
    key = '<w:spacing w:before="120" w:after="120"/>'
    assert key in en
    for k, (b, a) in ARMS.items():
        new = en.replace(key, f'<w:spacing w:before="{b}" w:after="{a}"/>')
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = new.encode("utf-8") if item.filename == "word/endnotes.xml" else zin.read(item.filename)
                zout.writestr(copy.copy(item), data)
        (OUT / f"ns_{k}.docx").write_bytes(buf.getvalue())
    print("ok")


def read():
    import fitz
    for pdf in sorted(OUT.glob("ns_*.pdf")):
        d = fitz.open(pdf)
        ys = []
        for p in d:
            for b in p.get_text("dict")["blocks"]:
                for l in b.get("lines", []):
                    s0 = l["spans"][0]
                    if round(s0["size"], 1) == 7.0 and s0["text"].strip()[:2].strip().isdigit():
                        ys.append((p.number + 1, round(s0["origin"][1], 1), s0["text"].strip()[:3]))
        ys.sort()
        gaps = [round(b[1] - a[1], 1) for a, b in zip(ys, ys[1:]) if a[0] == b[0]]
        print(pdf.stem, len(d), "pages", ys[:3], "gaps", gaps[:4])


if __name__ == "__main__":
    {"gen": gen, "read": read}[sys.argv[1]]()
