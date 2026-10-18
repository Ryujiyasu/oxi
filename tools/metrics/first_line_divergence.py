# -*- coding: utf-8 -*-
"""First line where Oxi breaks differently from Word.

Exports the document once with Word (fresh process, cached under
pipeline_data/first_line_div/<doc>.pdf), renders it with the given Oxi exe
(--dump-glyphs), groups both into lines by baseline, and walks the pages in order
until a line's last 20 non-space characters differ.  Prints the two lines, the next
ones, and both lines' x extents -- the starting point for a line-fit derivation.

    python tools/metrics/first_line_divergence.py <doc_id> <renderer.exe> [FLAG=1,...]
"""
import json, os, shutil, subprocess, sys, tempfile
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import feature_census as FC  # noqa: E402

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "first_line_div")


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


def word_lines(pdf):
    import fitz
    pages = []
    for pg in fitz.open(pdf):
        rows = {}
        for b in pg.get_text("rawdict")["blocks"]:
            for l in b.get("lines", []):
                for s in l["spans"]:
                    for c in s["chars"]:
                        if c["c"].strip():
                            rows.setdefault(round(c["origin"][1]), []).append(c)
        lines = []
        for y, v in sorted(rows.items()):
            v = sorted(v, key=lambda c: c["origin"][0])
            lines.append((y, round(v[0]["bbox"][0], 2), round(v[-1]["bbox"][2], 2), "".join(c["c"] for c in v)))
        pages.append(lines)
    return pages


def oxi_lines(exe, path, env_extra):
    t = tempfile.mkdtemp()
    try:
        env = dict(os.environ); env.update(env_extra)
        g = os.path.join(t, "g.json")
        subprocess.run([exe, path, os.path.join(t, "p"), "--dump-glyphs=" + g], env=env, capture_output=True, timeout=1200)
        pages = []
        for p in json.load(open(g, encoding="utf-8"))["pages"]:
            rows = {}
            for x in p["glyphs"]:
                if x["char"].strip():
                    rows.setdefault(round(x["baseline"]), []).append(x)
            pages.append([(y, round(v[0]["x"], 2), round(max(c["x"] for c in v), 2),
                           "".join(c["char"] for c in sorted(v, key=lambda c: c["x"]))) for y, v in sorted(rows.items())])
        return pages
    finally:
        shutil.rmtree(t, ignore_errors=True)


def main():
    did, exe = sys.argv[1], sys.argv[2]
    env = dict(kv.split("=", 1) for kv in (sys.argv[3] if len(sys.argv) > 3 else "").split(",") if kv)
    path = FC.load()[did]["path"]
    W, O = word_lines(word_pdf(path, did)), oxi_lines(exe, path, env)
    print("pages: word %d oxi %d" % (len(W), len(O)))
    for pi in range(min(len(W), len(O))):
        for k, (a, b) in enumerate(zip(W[pi], O[pi])):
            if a[3][-20:] != b[3][-20:]:
                print("page %d line %d" % (pi + 1, k))
                print("  W y%-5s x %.1f..%.1f  %s" % (a[0], a[1], a[2], a[3][-80:]))
                print("  O y%-5s x %.1f..%.1f  %s" % (b[0], b[1], b[2], b[3][-80:]))
                for tag, L in (("W", W[pi]), ("O", O[pi])):
                    if k + 1 < len(L):
                        print("  %s next: %s" % (tag, L[k + 1][3][-80:]))
                return
    print("no line divergence on the common pages")


if __name__ == "__main__":
    main()
