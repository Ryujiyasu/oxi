"""S467: localize the gen2-EN p1 accumulating vertical drift.
Measure EVERY page-1 paragraph Y via Word COM (collapsed-start Information(6),
R30) including list items, and diff against the Oxi GDI --dump-layout tops.
Per-paragraph drift = oxi_top - word_y pinpoints WHICH element loses height."""
import json, io, subprocess, os, sys
import win32com.client as win32

VPOS = 6  # wdVerticalPositionRelativeToPage
PAGE = 3  # wdActiveEndPageNumber
REPO = r"C:\Users\ryuji\oxi-main"
DOCS = {
    "gen2_067": r"gen2_067_SLA_Template.docx",
    "gen2_055": r"gen2_055_Risk_Management_Report.docx",
    "gen2_056": r"gen2_056_Product_Plan.docx",
}
RENDERER = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release", "oxi-gdi-renderer.exe")


def oxi_tops(dump):
    d = json.load(io.open(dump, encoding="utf-8"))
    rows = []  # (top, text) for page-1 body+cell text, sorted by top
    for pg in d["pages"]:
        if pg["page"] != 1:
            continue
        # group consecutive elements into visual lines by y
        for el in pg["elements"]:
            if el.get("type") != "text":
                continue
            t = (el.get("text") or "").strip()
            if not t:
                continue
            rows.append((round(el["y"], 1), t, el["x"]))
    rows.sort()
    return rows


def main():
    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    allout = []
    for tag, fname in DOCS.items():
        docx = os.path.join(REPO, "tools", "golden-test", "documents", "docx", fname)
        dump = os.path.join(r"C:\Users\ryuji\AppData\Local\Temp", "s467_%s.json" % tag)
        out_prefix = os.path.join(r"C:\Users\ryuji\AppData\Local\Temp", "s467_%s" % tag)
        subprocess.run([RENDERER, docx, out_prefix, "150", "--dump-layout=" + dump],
                       capture_output=True, text=True)
        otops = oxi_tops(dump)

        doc = word.Documents.Open(docx, ReadOnly=True)
        winfo = []
        for p in doc.Paragraphs:
            rng = p.Range
            st = doc.Range(rng.Start, rng.Start)
            if st.Information(PAGE) != 1:
                continue
            t = p.Range.Text.strip()
            if not t:
                t = "(empty)"
            pf = p.Format
            lt = p.Range.ListFormat.ListString  # bullet/number marker text
            winfo.append(dict(t=t, y=round(st.Information(VPOS), 2),
                              sb=round(pf.SpaceBefore, 2), sa=round(pf.SpaceAfter, 2),
                              lsr=pf.LineSpacingRule, ls=round(pf.LineSpacing, 2),
                              marker=lt))
        doc.Close(False)

        lines = ["=== %s (Word page-1 paragraphs) ===" % tag,
                 "%-28s %7s %5s %5s %4s %6s %6s %8s" % ("text", "word_y", "sb", "sa", "lsr", "ls", "oxi_top", "drift")]
        prev_y = None
        for rec in winfo:
            # match oxi top: nearest oxi line whose text startswith the word text prefix
            wp = rec["t"][:8]
            cand = [(abs(oy - rec["y"]), oy, ot) for oy, ot, ox in otops
                    if ot[:8] == wp or (wp and ot.startswith(wp[:5]))]
            oxi_top = min(cand)[1] if cand else None
            drift = round(oxi_top - rec["y"], 2) if oxi_top is not None else None
            gap = round(rec["y"] - prev_y, 2) if prev_y is not None else 0.0
            prev_y = rec["y"]
            mk = ("[%s]" % rec["marker"]) if rec["marker"] else ""
            lines.append("%-28s %7.2f %5.1f %5.1f %4d %6.2f %6s %8s  gap=%+.2f %s" % (
                (rec["t"][:26] + mk)[:28], rec["y"], rec["sb"], rec["sa"], rec["lsr"],
                rec["ls"], ("%.1f" % oxi_top) if oxi_top else "-",
                ("%+.2f" % drift) if drift is not None else "-", gap, mk))
        allout.append("\n".join(lines))
    word.Quit()
    txt = "\n\n".join(allout)
    io.open(os.path.join(REPO, "tools", "metrics", "_s467_gen2_drift.out"), "w", encoding="utf-8").write(txt)
    # also print ascii-safe
    print(txt.encode("ascii", "replace").decode())


if __name__ == "__main__":
    main()
