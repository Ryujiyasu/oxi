"""Extract chars-indent paragraphs on d77a pages {3, 6, 8, 9, 10}.

Pre-filter at XML level so we only need to query Word COM for paragraphs
that actually use chars-based indents (avoids Word session timeout on
207 paragraphs).
"""
import json, re, zipfile, sys, time
from pathlib import Path
import win32com.client as w32

DOC = Path(r"C:\Users\ryuji\oxi-1\tools\golden-test\documents\docx\d77a58485f16_20240705_resources_data_outline_08.docx")
TARGET_PAGES = [3, 6, 8, 9, 10]
OUT = Path(r"C:\Users\ryuji\oxi-1\pipeline_data\d77a_chars_indent_extract.json")


def parse_paragraphs_xml(docx_path):
    with zipfile.ZipFile(docx_path) as z:
        xml = z.read("word/document.xml").decode("utf-8")
        try:
            styles_xml = z.read("word/styles.xml").decode("utf-8")
        except KeyError:
            styles_xml = ""

    docdef_sz = None
    m = re.search(
        r'<w:docDefaults>.*?<w:rPrDefault>.*?<w:rPr>.*?<w:sz w:val="(\d+)"',
        styles_xml, re.S
    )
    if m:
        docdef_sz = int(m.group(1))

    paras = []
    for m in re.finditer(r'<w:p\b[^>]*>(.*?)</w:p>', xml, re.S):
        body = m.group(1)
        ppr_m = re.search(r'<w:pPr\b[^>]*>(.*?)</w:pPr>', body, re.S)
        ppr = ppr_m.group(1) if ppr_m else ""

        ind = {}
        ind_m = re.search(r'<w:ind\b([^/>]*)/?>', ppr)
        if ind_m:
            attrs = ind_m.group(1)
            for k in ("left","leftChars","right","rightChars",
                      "firstLine","firstLineChars",
                      "hanging","hangingChars",
                      "start","startChars","end","endChars"):
                a = re.search(rf'w:{k}="(-?\d+)"', attrs)
                if a: ind[k] = int(a.group(1))

        numpr = None
        nm = re.search(r'<w:numPr>(.*?)</w:numPr>', ppr, re.S)
        if nm:
            inner = nm.group(1)
            ilvl = re.search(r'<w:ilvl w:val="(\d+)"', inner)
            numid = re.search(r'<w:numId w:val="(\d+)"', inner)
            numpr = {
                "ilvl": int(ilvl.group(1)) if ilvl else None,
                "numId": int(numid.group(1)) if numid else None,
            }

        ppr_sz = None
        m2 = re.search(r'<w:rPr>.*?<w:sz w:val="(\d+)"', ppr, re.S)
        if m2: ppr_sz = int(m2.group(1))

        run_sz = None
        run_m = re.search(r'<w:r\b[^>]*>(.*?)</w:r>', body, re.S)
        run_text = ""
        if run_m:
            run_body = run_m.group(1)
            sz_m = re.search(r'<w:rPr>.*?<w:sz w:val="(\d+)"', run_body, re.S)
            if sz_m: run_sz = int(sz_m.group(1))
            txt_parts = re.findall(r'<w:t[^>]*>([^<]*)</w:t>', body)
            run_text = "".join(txt_parts)[:60]

        style = None
        sm = re.search(r'<w:pStyle w:val="([^"]*)"', ppr)
        if sm: style = sm.group(1)

        has_chars = any(k in ind for k in ("leftChars","rightChars","firstLineChars","hangingChars","startChars","endChars"))

        paras.append({
            "ind": ind,
            "numPr": numpr,
            "ppr_sz": ppr_sz,
            "run_sz": run_sz,
            "style": style,
            "text": run_text,
            "has_chars_indent": has_chars,
        })
    return paras, docdef_sz


def main():
    parsed, docdef_sz = parse_paragraphs_xml(DOC)
    cw_pt = (docdef_sz / 2.0) if docdef_sz else None
    chars_indent_indices = [i+1 for i, p in enumerate(parsed) if p.get("has_chars_indent")]
    print(f"Total paragraphs in XML: {len(parsed)}", file=sys.stderr)
    print(f"chars-indent paragraphs: {len(chars_indent_indices)}", file=sys.stderr)
    print(f"docDefault.sz = {docdef_sz} (char_width = {cw_pt}pt)", file=sys.stderr)

    out = {
        "docDefault_sz": docdef_sz,
        "char_width_pt": cw_pt,
        "total_paras_xml": len(parsed),
        "n_chars_indent_paras": len(chars_indent_indices),
        "pages": {},
    }

    word = w32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(str(DOC.resolve()), ReadOnly=True)
        try:
            print(f"Word reports {doc.Paragraphs.Count} paragraphs", file=sys.stderr)
            for idx in chars_indent_indices:
                try:
                    p = doc.Paragraphs(idx)
                    page = p.Range.Information(3)
                    if page not in TARGET_PAGES:
                        continue
                    fmt = p.Format
                    li = fmt.LeftIndent
                    ri = fmt.RightIndent
                    fli = fmt.FirstLineIndent
                    y = p.Range.Information(6)
                    txt = (p.Range.Text or "")[:60].replace("\r","\\r").replace("\x07","\\x07")
                    parsed_p = parsed[idx-1]
                    rec = {
                        "para_idx_word": idx,
                        "page": page,
                        "y": y,
                        "format_LI_pt": li,
                        "format_RI_pt": ri,
                        "format_FLI_pt": fli,
                        "xml_ind": parsed_p.get("ind"),
                        "xml_numPr": parsed_p.get("numPr"),
                        "xml_ppr_sz": parsed_p.get("ppr_sz"),
                        "xml_run_sz": parsed_p.get("run_sz"),
                        "xml_style": parsed_p.get("style"),
                        "text": txt,
                    }
                    out["pages"].setdefault(page, []).append(rec)
                    # Save incrementally every 10 hits
                    if sum(len(v) for v in out["pages"].values()) % 5 == 0:
                        OUT.parent.mkdir(parents=True, exist_ok=True)
                        with open(OUT, "w", encoding="utf-8") as f:
                            json.dump(out, f, indent=2, ensure_ascii=False)
                except Exception as e:
                    print(f"  ERROR para {idx}: {e}", file=sys.stderr)
                    # save what we have and stop
                    break
        finally:
            try: doc.Close(SaveChanges=0)
            except: pass
    finally:
        try: word.Quit()
        except: pass

    OUT.parent.mkdir(parents=True, exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(out, f, indent=2, ensure_ascii=False)

    print()
    print(f"docDefault sz = {docdef_sz} (char_width = {cw_pt}pt)")
    print()
    for pg in sorted(out["pages"]):
        rows = out["pages"][pg]
        print(f"=== Page {pg}: {len(rows)} chars-indent paras ===")
        for r in rows:
            ind = r["xml_ind"] or {}
            ind_summary = ", ".join(f"{k}={v}" for k,v in sorted(ind.items()))
            np = r["xml_numPr"]
            np_str = f"numId={np['numId']} ilvl={np['ilvl']}" if np else "—"
            print(f"  p{r['para_idx_word']:3d} y={r['y']:6.1f} "
                  f"LI={r['format_LI_pt']:5.1f} FLI={r['format_FLI_pt']:+6.1f} "
                  f"sty={(r['xml_style'] or '-'):>8s} {np_str:18s} "
                  f"ppr_sz={r['xml_ppr_sz'] or '-'} run_sz={r['xml_run_sz'] or '-'}")
            print(f"        ind={{{ind_summary}}}")
            print(f"        text={r['text']!r}")


if __name__ == "__main__":
    main()
