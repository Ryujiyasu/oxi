# -*- coding: utf-8 -*-
"""Which real lines does Word hang?  Render each document with the line-end hang
on (default) and off (OXI_S809_DISABLE=1), take every run of lines where the two
Oxi layouts break differently, and ask the Word PDF which of the two it matches.
Prints, per differing group, the ON/OFF line ends and the verdict, plus the face,
size and last glyph of the line that the hang changed.

    python tools/metrics/hang_census.py <renderer.exe> <doc_id> [<doc_id> ...]

Word PDFs are exported once per document in a fresh Word into
pipeline_data/hang_census/<doc>.pdf (re-used afterwards).
"""
import difflib, json, os, shutil, subprocess, sys, tempfile
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import feature_census as FC  # noqa: E402

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "hang_census")


def word_pdf(path, did):
    os.makedirs(OUT, exist_ok=True)
    pdf = os.path.join(OUT, did.replace("/", "__") + ".pdf")
    if not os.path.exists(pdf):
        import win32com.client
        w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
        try:
            d = w.Documents.Open(os.path.abspath(path), ReadOnly=True); d.ExportAsFixedFormat(pdf, 17); d.Close(False)
        finally:
            w.Quit()
    return pdf


def word_line_ends(pdf):
    import fitz
    ends = set()
    for pg in fitz.open(pdf):
        for b in pg.get_text("rawdict")["blocks"]:
            for l in b.get("lines", []):
                ch = [c for s in l["spans"] for c in s["chars"] if c["c"].strip()]
                if ch:
                    ends.add("".join(c["c"] for c in ch)[-18:])
    return ends


def oxi_lines(exe, path, env_extra):
    t = tempfile.mkdtemp()
    try:
        env = dict(os.environ); env.update(env_extra)
        g = os.path.join(t, "g.json")
        subprocess.run([exe, path, os.path.join(t, "p"), "--dump-glyphs=" + g], env=env, capture_output=True, timeout=900)
        out = []
        for p in json.load(open(g, encoding="utf-8"))["pages"]:
            rows = {}
            for x in p["glyphs"]:
                rows.setdefault(round(x["baseline"]), []).append(x)
            for _, v in sorted(rows.items()):
                v = sorted(v, key=lambda c: c["x"])
                out.append(("".join(c["char"] for c in v), v[-1]["font_family"], v[-1]["font_size"]))
        return out
    finally:
        shutil.rmtree(t, ignore_errors=True)


def main():
    exe, docs = sys.argv[1], sys.argv[2:]
    rows = FC.load()
    tally = {"on": 0, "off": 0, "neither": 0}
    for did in docs:
        path = rows[did]["path"]
        ends = word_line_ends(word_pdf(path, did))
        on, off = oxi_lines(exe, path, {}), oxi_lines(exe, path, {"OXI_S809_DISABLE": "1"})
        sm = difflib.SequenceMatcher(None, [x[0] for x in on], [x[0] for x in off], autojunk=False)
        print("==", did)
        for tag, i1, i2, j1, j2 in sm.get_opcodes():
            if tag == "equal":
                continue
            a, b = on[i1:i2], off[j1:j2]
            ok_on = sum(x[0][-18:] in ends for x in a) == len(a)
            ok_off = sum(x[0][-18:] in ends for x in b) == len(b)
            v = "on" if ok_on and not ok_off else "off" if ok_off and not ok_on else "neither"
            tally[v] += 1
            hung = a[0] if a else None
            print("  %-7s ON %-40r OFF %-40r face %s %s" % (
                v, hung[0][-38:] if hung else "", b[0][0][-38:] if b else "", hung[1] if hung else "", hung[2] if hung else ""))
    print("verdicts:", tally)


if __name__ == "__main__":
    main()
