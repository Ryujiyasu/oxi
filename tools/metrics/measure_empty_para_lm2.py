"""Measure Word's empty-paragraph height in LM2 docGrid for various rPr sz values.

Tests hypothesis: Oxi over-allocates empty 12pt MS Gothic para in LM2 linePitch=360.
Expected Word behavior: 1 grid pitch = 18pt regardless of rPr size (up to some
threshold where it snaps to 2 pitches).
"""
import io, json, os, sys, time, zipfile
from pathlib import Path
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT = Path(__file__).with_name("output") / "empty_para_lm2.json"
OUT.parent.mkdir(parents=True, exist_ok=True)
TMP = Path("pipeline_data") / "_empty_para_lm2_tmp.docx"
TMP.parent.mkdir(parents=True, exist_ok=True)

CT = '<?xml version="1.0"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
RELS = '<?xml version="1.0"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'

# Test matrix: pmark size (half-points) × font
FONTS = [("ＭＳ ゴシック", "MS Gothic")]
# Target sizes in half-points: sz=21 (10.5pt), 22 (11pt), 24 (12pt), 26 (13pt), 28 (14pt)
SZS = [21, 22, 24, 26, 28]
LINE_PITCHES = [360, 350]


def build(font, pmark_sz, line_pitch):
    pmark_rpr = f'<w:rPr><w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}"/><w:sz w:val="{pmark_sz}"/><w:szCs w:val="{pmark_sz}"/></w:rPr>'
    content_rpr = f'<w:rPr><w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr>'
    # Three paras: content(anchor) | empty(with pmark sz) | content(anchor)
    anchor = lambda label: f'<w:p><w:pPr>{content_rpr}</w:pPr><w:r>{content_rpr}<w:t>{label}</w:t></w:r></w:p>'
    empty = f'<w:p><w:pPr>{pmark_rpr}</w:pPr></w:p>'
    body = anchor("P1") + empty + anchor("P3")
    sect = f'<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="851" w:bottom="1134" w:left="851" w:header="851" w:footer="992" w:gutter="0"/><w:docGrid w:type="lines" w:linePitch="{line_pitch}"/></w:sectPr>'
    xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:body>{body}{sect}</w:body></w:document>'
    )
    with zipfile.ZipFile(TMP, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/document.xml", xml)


def measure(word, font, pmark_sz, line_pitch):
    try: os.remove(TMP)
    except FileNotFoundError: pass
    build(font, pmark_sz, line_pitch)
    last_err = None
    for attempt in range(4):
        try:
            doc = word.Documents.Open(str(TMP.resolve()), ReadOnly=True)
            time.sleep(0.3)
            y1 = doc.Paragraphs(1).Range.Information(6)
            y2 = doc.Paragraphs(2).Range.Information(6)
            y3 = doc.Paragraphs(3).Range.Information(6)
            doc.Close(False)
            return {
                "font": font, "pmark_sz_half": pmark_sz, "pmark_pt": pmark_sz / 2.0,
                "line_pitch_tw": line_pitch, "line_pitch_pt": line_pitch / 20.0,
                "y1": round(y1, 2), "y2": round(y2, 2), "y3": round(y3, 2),
                "empty_para_h": round(y3 - y2, 2),
                "p1_to_p2_gap": round(y2 - y1, 2),
            }
        except Exception as e:
            last_err = e
            time.sleep(0.8 + attempt * 0.5)
    return {"font": font, "pmark_sz_half": pmark_sz, "line_pitch_tw": line_pitch, "error": str(last_err)}


def main():
    word = win32com.client.Dispatch("Word.Application")
    time.sleep(1.0)
    word.Visible = False
    word.DisplayAlerts = False
    results = []
    try:
        for font, pretty in FONTS:
            for lp in LINE_PITCHES:
                for sz in SZS:
                    m = measure(word, font, sz, lp)
                    m["font_pretty"] = pretty
                    results.append(m)
                    if "error" in m:
                        print(f"{pretty} sz={sz/2}pt pitch={lp}tw: ERR {m['error']}")
                    else:
                        print(f"{pretty} sz={sz/2:4.1f}pt pitch={lp}tw: empty_para_h={m['empty_para_h']}  (P1-P2={m['p1_to_p2_gap']} anchors)")
    finally:
        try: word.Quit()
        except: pass
        try: os.remove(TMP)
        except: pass
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, ensure_ascii=False)
    print(f"\nSaved -> {OUT}")


if __name__ == "__main__":
    main()
