"""
Measure Word's actual footnote separator + footnote block height for b837-like docs.

Goal: find the real per-page overhead Oxi is missing. Hypothesis: Word reserves
a separator region (~line-height) once per page with footnotes, plus each
footnote's actual rendered height.

Test matrix: 1, 2, 3, 5, 7 footnotes per page with MS Mincho 10.5pt body/fn.
Measure:
  - last_body_y: y of last body paragraph's BOTTOM
  - first_fn_y: y of first footnote (top)
  - last_fn_y: y of last footnote (top)
  - separator_gap = first_fn_y − last_body_y (should include separator area)
  - total_fn_block_h = last_fn_y + fn_line_h − first_fn_y
"""
import io, json, os, sys, time, zipfile
from pathlib import Path
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT = Path(__file__).with_name("output") / "footnote_separator.json"
OUT.parent.mkdir(parents=True, exist_ok=True)
TMP = Path("pipeline_data") / "_footnote_separator_tmp.docx"
TMP.parent.mkdir(parents=True, exist_ok=True)

CT = '<?xml version="1.0"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>'
RELS_DOC = '<?xml version="1.0"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
# document.xml.rels links to footnotes.xml
DOC_RELS = '<?xml version="1.0"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rFn" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>'

FONT = "ＭＳ 明朝"
SZ_HALF = 21  # 10.5pt


def rpr(sz_half=SZ_HALF):
    return f'<w:rPr><w:rFonts w:ascii="{FONT}" w:eastAsia="{FONT}" w:hAnsi="{FONT}"/><w:sz w:val="{sz_half}"/><w:szCs w:val="{sz_half}"/></w:rPr>'


def build(n_footnotes):
    """Single-page doc with n_footnotes refs in one body paragraph."""
    # Body: one paragraph with N footnote refs
    refs = "".join(
        f'<w:r>{rpr()}<w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="{i+1}"/></w:r>'
        for i in range(n_footnotes)
    )
    # Simpler: each ref in its own w:r, text "ref1 ref2 ..."
    body_runs = ""
    for i in range(n_footnotes):
        body_runs += (
            f'<w:r>{rpr()}<w:t xml:space="preserve">ref{i+1}</w:t></w:r>'
            f'<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="{i+2}"/></w:r>'
        )
    body = f'<w:p><w:pPr>{rpr()}</w:pPr>{body_runs}</w:p>'
    sect = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="851" w:bottom="1134" w:left="851" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>'
    doc_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:body>{body}{sect}</w:body></w:document>'
    )

    # Footnotes: standard separator (id=0) and continuationSeparator (id=1), plus user footnotes id=2..n+1
    fn_entries = [
        '<w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>',
        '<w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>',
    ]
    for i in range(n_footnotes):
        txt = f"note {i+1} body text"
        fn_entries.append(
            f'<w:footnote w:id="{i+2}"><w:p><w:pPr>{rpr()}</w:pPr>'
            f'<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>'
            f'<w:r>{rpr()}<w:t xml:space="preserve"> {txt}</w:t></w:r></w:p></w:footnote>'
        )
    footnotes_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        + "".join(fn_entries) +
        '</w:footnotes>'
    )

    with zipfile.ZipFile(TMP, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS_DOC)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/document.xml", doc_xml)
        z.writestr("word/footnotes.xml", footnotes_xml)


def measure(word, n):
    try: os.remove(TMP)
    except FileNotFoundError: pass
    build(n)
    last_err = None
    for attempt in range(4):
        try:
            doc = word.Documents.Open(str(TMP.resolve()), ReadOnly=True)
            time.sleep(0.4)
            # Body paragraph info
            body_p = doc.Paragraphs(1)
            body_y = body_p.Range.Information(6)  # top Y
            body_end_y = doc.Range(body_p.Range.End - 1, body_p.Range.End).Information(6)

            # Footnotes collection
            fns = doc.Footnotes
            fn_count = fns.Count
            fn_ys = []
            for i in range(1, fn_count + 1):
                fn = fns(i)
                fy = fn.Range.Information(6)
                fn_ys.append(round(fy, 2))

            doc.Close(False)
            return {
                "n_footnotes": n,
                "body_top_y": round(body_y, 2),
                "body_end_y": round(body_end_y, 2),
                "fn_count": fn_count,
                "fn_ys": fn_ys,
                "first_fn_y": fn_ys[0] if fn_ys else None,
                "last_fn_y": fn_ys[-1] if fn_ys else None,
                "separator_gap": round(fn_ys[0] - body_end_y, 2) if fn_ys else None,
                "fn_block_h": round(fn_ys[-1] - fn_ys[0], 2) if len(fn_ys) >= 2 else 0.0,
            }
        except Exception as e:
            last_err = e
            time.sleep(0.8)
    return {"n_footnotes": n, "error": str(last_err)}


def main():
    word = win32com.client.Dispatch("Word.Application")
    time.sleep(1.0)
    word.Visible = False
    word.DisplayAlerts = False
    results = []
    try:
        for n in [1, 2, 3, 5, 7]:
            m = measure(word, n)
            results.append(m)
            if "error" in m:
                print(f"n={n}: ERR {m['error']}")
            else:
                print(f"n={n}: body_end={m['body_end_y']}, first_fn={m['first_fn_y']}, "
                      f"last_fn={m['last_fn_y']}, separator_gap={m['separator_gap']}, "
                      f"fn_block_h={m['fn_block_h']}, fn_ys={m['fn_ys']}")
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
