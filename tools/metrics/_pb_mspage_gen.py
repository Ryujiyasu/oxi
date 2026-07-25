"""Multiple-spacing page-capacity: is Oxi's first-line-of-page ~11pt too high (fits
1 extra line/page) GENERAL to Latin multiple-spacing, or specific to the Bidi
tall-line (educational__002354115a)?

Plain TNR (Latin, no Bidi), no-type docGrid linePitch=360, line spacing swept
(276/360/480 = 1.15/1.5/2.0x auto). N single-line body paras. Measure Word
per-page line count + first-line baseline (COM Info 3/6) and compare to Oxi.
If Oxi fits 1 more line / first line higher -> GENERAL (S695-family page-capacity,
gen2/contract canary-heavy, worth a multi-session fix). If they match -> the bug
needs the Bidi tall-line source (rare, low value).

Usage: python _pb_mspage_gen.py gen | measure | oxi
"""
import os, sys, json, zipfile

OUTDIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                      "pipeline_data", "_pb_mspage")
FONT = "Times New Roman"


def para(i, sz):
    r = f'<w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:cs="{FONT}"/><w:sz w:val="{sz}"/>'
    return (f'<w:p><w:pPr><w:jc w:val="left"/><w:rPr>{r}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{r}</w:rPr><w:t>Body line number {i} of the multiple spacing sweep</w:t></w:r></w:p>')


def build(line, sz, n=45):
    # per-paragraph line spacing (auto, multiple)
    body = ""
    for i in range(n):
        r = f'<w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:cs="{FONT}"/><w:sz w:val="{sz}"/>'
        body += (f'<w:p><w:pPr><w:spacing w:after="0" w:line="{line}" w:lineRule="auto"/>'
                 f'<w:jc w:val="left"/><w:rPr>{r}</w:rPr></w:pPr>'
                 f'<w:r><w:rPr>{r}</w:rPr><w:t>Body line number {i} of the multiple spacing sweep test</w:t></w:r></w:p>')
    grid = '<w:docGrid w:linePitch="360"/>'
    mar = ('<w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418" '
           'w:header="851" w:footer="992" w:gutter="0"/>')
    sect = f'<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>{mar}{grid}</w:sectPr>'
    return body + sect


def docxml(body):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'<w:body>{body}</w:body></w:document>')


def write_docx(path, body, sz):
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
          '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
          '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
          '</Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
            '</Relationships>')
    drels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
             '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
             '</Relationships>')
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
              f'<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:cs="{FONT}"/>'
              f'<w:sz w:val="{sz}"/></w:rPr></w:rPrDefault></w:docDefaults>'
              '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
              '</w:styles>')
    settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:settings>')
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/_rels/document.xml.rels", drels)
        z.writestr("word/document.xml", docxml(body))
        z.writestr("word/styles.xml", styles)
        z.writestr("word/settings.xml", settings)


CASES = [(line, 24) for line in (276, 360, 480)]  # 1.15x / 1.5x / 2.0x, TNR 12pt


def nm(line, sz):
    return f"ms_l{line}_s{sz}.docx"


def gen():
    os.makedirs(OUTDIR, exist_ok=True)
    for line, sz in CASES:
        write_docx(os.path.join(OUTDIR, nm(line, sz)), build(line, sz), sz)
    print(f"wrote {len(CASES)} docs to {OUTDIR}")


def measure():
    import win32com.client
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False; word.DisplayAlerts = 0
    out = []
    try:
        for line, sz in CASES:
            path = os.path.abspath(os.path.join(OUTDIR, nm(line, sz)))
            d = word.Documents.Open(path, ReadOnly=True); d.Repaginate()
            n = d.Paragraphs.Count
            # first para on p2 -> p1 capacity; first-line baseline of p1
            y1 = d.Range(d.Paragraphs(1).Range.Start, d.Paragraphs(1).Range.Start).Information(6)
            cap = None; y_p2 = None
            for i in range(1, n + 1):
                st = d.Range(d.Paragraphs(i).Range.Start, d.Paragraphs(i).Range.Start)
                if int(st.Information(3)) >= 2:
                    cap = i - 1; y_p2 = st.Information(6); break
            out.append({"line": line, "sz": sz, "p1_cap": cap, "y1": round(y1, 2),
                        "y_p2_first": round(y_p2, 2) if y_p2 else None})
            print(f"  line={line} ({line/240:.2f}x): Word p1 holds {cap} lines, first baseline y={y1:.2f}")
            d.Close(False)
    finally:
        word.Quit()
    json.dump(out, open(os.path.join(OUTDIR, "_word.json"), "w"), indent=1)
    print("-> _word.json")


def oxi():
    """Render each in Oxi GDI, count lines on p1, first-line y."""
    import subprocess
    GDI = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..",
                       "oxi-gdi-renderer", "target", "release", "oxi-gdi-renderer.exe")
    word = json.load(open(os.path.join(OUTDIR, "_word.json"))) if os.path.exists(os.path.join(OUTDIR, "_word.json")) else []
    wmap = {(w["line"], w["sz"]): w for w in word}
    for line, sz in CASES:
        docx = os.path.abspath(os.path.join(OUTDIR, nm(line, sz)))
        dump = os.path.join(OUTDIR, f"oxi_l{line}.json")
        subprocess.run([GDI, docx, os.path.join(OUTDIR, "p_"), f"--dump-layout={dump}"],
                       capture_output=True)
        d = json.load(open(dump, encoding="utf-8"))
        pgs = d.get("pages", d)
        pgs = pgs.get("pages", pgs) if isinstance(pgs, dict) else pgs
        p1 = [e for e in pgs[0].get("elements", pgs[0]) if (e.get("text") or "").strip()]
        ys = sorted(set(round(e.get("y", 0), 2) for e in p1))
        oy1 = ys[0] if ys else None
        ocap = len(ys)
        w = wmap.get((line, sz), {})
        wc = w.get("p1_cap"); wy1 = w.get("y1")
        flag = ""
        if wc is not None and ocap != wc:
            flag = f"  <== Oxi {ocap} vs Word {wc} ({'GENERAL bug' if ocap>wc else 'Oxi fewer'})"
        print(f"  line={line} ({line/240:.2f}x): Oxi p1 {ocap} lines (first y={oy1}) | Word {wc} lines (y1={wy1}){flag}")


if __name__ == "__main__":
    a = sys.argv[1:]
    if a == ["gen"]: gen()
    elif a == ["measure"]: measure()
    elif a == ["oxi"]: oxi()
    else: print(__doc__)
