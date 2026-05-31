"""S468 confirmation: on REAL regressing CJK docs, verify the unification:
Word quantizes line-box tops to the 0.75pt grid; Oxi quantizes to 0.5pt.
Per-paragraph: Word Information(6) (collapsed-start, R30) vs Oxi dump el.y
(LINE-BOX-TOP convention; subtract text_y_off is NOT needed since el.y is
already the box top). Report each side's grid residual distribution.
"""
import json, io, subprocess, os, glob
import win32com.client as win32

VPOS = 6
PAGE = 3
REPO = r"C:\Users\ryuji\oxi-main"
RENDERER = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release", "oxi-gdi-renderer.exe")
TMP = r"C:\Users\ryuji\AppData\Local\Temp"
DOCS = ["0e7af1ae8f21", "3a4f9fbe1a83", "2ea81a8441cc", "34140b9c5662"]


def grid_res(y, pitch):
    return round(abs(round(y / pitch) * pitch - y), 4)


def oxi_box_tops(path):
    dump = os.path.join(TMP, "s468rd.json")
    subprocess.run([RENDERER, path, os.path.join(TMP, "s468rd"), "150", "--dump-layout=" + dump],
                   capture_output=True, text=True)
    d = json.load(io.open(dump, encoding="utf-8"))
    seen = set()
    for pg in d["pages"]:
        if pg["page"] != 1:
            continue
        for el in pg["elements"]:
            if el.get("type") != "text":
                continue
            if not (el.get("text") or "").strip():
                continue
            seen.add(round(el["y"], 3))  # LINE-BOX-TOP per dump convention
    return sorted(seen)


def main():
    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    out = []
    for stub in DOCS:
        matches = glob.glob(os.path.join(REPO, "tools", "golden-test", "documents", "docx", "*%s*.docx" % stub))
        matches = [m for m in matches if "layout" not in m]
        if not matches:
            out.append("%s: NOT FOUND" % stub); continue
        path = os.path.normpath(matches[0])
        otops = oxi_box_tops(path)

        try:
            doc = word.Documents.Open(path, ReadOnly=True, AddToRecentFiles=False)
        except Exception as e:
            out.append("%s: OPEN FAILED %s" % (stub, e)); out.append(""); continue
        wtops = []
        for p in doc.Paragraphs:
            rng = p.Range
            st = doc.Range(rng.Start, rng.Start)
            if st.Information(PAGE) != 1:
                continue
            if not p.Range.Text.strip():
                continue
            wtops.append(round(st.Information(VPOS), 3))
        doc.Close(False)

        def summarize(tops, pitch):
            res = [grid_res(y, pitch) for y in tops]
            on = sum(1 for r in res if r < 0.02)
            return on, len(res), (max(res) if res else 0)

        w075 = summarize(wtops, 0.75); w05 = summarize(wtops, 0.5)
        o075 = summarize(otops, 0.75); o05 = summarize(otops, 0.5)
        out.append("=== %s ===" % os.path.basename(path)[:40])
        out.append("  Word tops n=%d : on-0.75 %d/%d (max res %.3f) | on-0.50 %d/%d (max %.3f)"
                   % (len(wtops), w075[0], w075[1], w075[2], w05[0], w05[1], w05[2]))
        out.append("  Oxi  tops n=%d : on-0.75 %d/%d (max res %.3f) | on-0.50 %d/%d (max %.3f)"
                   % (len(otops), o075[0], o075[1], o075[2], o05[0], o05[1], o05[2]))
        out.append("  Word tops[:6]=%s" % [round(y, 2) for y in wtops[:6]])
        out.append("  Oxi  tops[:6]=%s" % [round(y, 2) for y in otops[:6]])
        out.append("")
    word.Quit()
    txt = "\n".join(out)
    io.open(os.path.join(REPO, "tools", "metrics", "_s468_realdoc_grid.out"), "w", encoding="utf-8").write(txt)
    print(txt.encode("ascii", "replace").decode())


if __name__ == "__main__":
    main()
