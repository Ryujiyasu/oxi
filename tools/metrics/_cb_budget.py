# -*- coding: utf-8 -*-
"""Break budget vs rendered advance, on the same line, next to Word's own line.

The char-budget wall shows up as lines Oxi draws PAST the text edge.  The
`OXI_DUMP_LINEW` instrument prints what the breaker thought the line was worth
(its running width, the sum of the fragment widths it stored, and the budget it
fit to); the layout dump prints what the emitter actually drew.  When those two
disagree the line is broken to one width and drawn at another, which is the
[[latin_text_wrap_compression]] failure in its CJK form.

This joins the three by the line's leading characters:

    breaker (cur_w / frag_sum / avail)  |  emitter (x0..x1)  |  Word (x0..x1)

    python _cb_budget.py c7b923e5            # every over-wide line, worst first
    python _cb_budget.py c7b923e5 --page 3   # one page, every line
"""
import collections
import json
import os
import re
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
os.environ.setdefault("PYTHONIOENCODING", "utf-8")
sys.stdout.reconfigure(encoding="utf-8")

REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_cb_budget")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")
DOCS = os.path.join(REPO, "tools", "golden-test", "documents", "docx")

LINEW = re.compile(
    r"\[LINEW\] nch=(\d+) frag_sum=([-\d.]+) nat_sum=([-\d.]+) avail=([-\d.]+) "
    r"over=([-\d.]+) head=(.*)$")


def docx_for(prefix):
    return next(os.path.join(DOCS, f) for f in sorted(os.listdir(DOCS))
                if f.startswith(prefix) and f.endswith(".docx")
                and not f.startswith("~$"))


def run_oxi(docx, tag="base", envs=""):
    """Render once with the breaker dump on; return (linew records, layout)."""
    os.makedirs(OUT, exist_ok=True)
    layout = os.path.join(OUT, "%s_%s.json" % (os.path.basename(docx)[:12], tag))
    env = dict(os.environ)
    env["OXI_DUMP_LINEW"] = "1"
    for kv in [s for s in envs.split(",") if s]:
        k, _, v = kv.partition("=")
        env[k] = v if v != "" else "1"
    r = subprocess.run([GDI, docx, os.path.join(OUT, "png_" + tag),
                        "--dump-layout=" + layout],
                       capture_output=True, env=env)
    if r.returncode != 0:
        sys.exit("renderer failed: %s" % r.stderr.decode("utf-8", "replace")[-2000:])
    recs = []
    for line in r.stderr.decode("utf-8", "replace").splitlines():
        m = LINEW.search(line)
        if m:
            recs.append({
                "nch": int(m.group(1)), "frag_sum": float(m.group(2)),
                "nat_sum": float(m.group(3)), "avail": float(m.group(4)),
                "over": float(m.group(5)), "head": m.group(6),
            })
    return recs, json.load(open(layout, encoding="utf-8"))


def oxi_lines(layout):
    """Emitted text lines.  ★Group by (cell, y): y alone fuses the cells of a
    table row into one line and invents an over-wide line (the trap the census
    hit).  Returns page, x0, x1, width sum, text."""
    out = []
    for pi, page in enumerate(layout["pages"], 1):
        rows = collections.OrderedDict()
        for e in page["elements"]:
            # ★Whitespace-only lines are dropped on BOTH sides. The Word reader
            # skips them (`txt.strip()`), so keeping Oxi's shifted every line of
            # a page down by one and scored a page that matches exactly as a
            # total miss -- c7b923e5's page 2 opens with an ideographic space in
            # both engines.
            if e["type"] != "text" or not (e.get("text") or "").strip():
                continue
            key = (round(e["y"], 1), e.get("cell_row_idx"), e.get("cell_col_idx"))
            rows.setdefault(key, []).append(e)
        for (y, _r, _c), els in rows.items():
            els.sort(key=lambda e: e["x"])
            out.append({
                "page": pi, "y": y,
                "x0": min(e["x"] for e in els),
                "x1": max(e["x"] + (e.get("w") or 0.0) for e in els),
                "wsum": sum(e.get("w") or 0.0 for e in els),
                "text": "".join(e["text"] for e in els),
            })
    return out


def word_lines(docx):
    """Word's own lines from its PDF, per page: x0, x1, text."""
    import fitz
    # The corpus ships Word's own export for many documents; use it rather than
    # driving Word again.
    rt = docx[:-5] + "_rt.pdf"
    pdf = rt if os.path.exists(rt) else os.path.join(
        OUT, os.path.basename(docx)[:12] + ".pdf")
    if not os.path.exists(pdf):
        os.makedirs(OUT, exist_ok=True)
        import win32com.client as w
        app = w.DispatchEx("Word.Application")
        app.Visible = False
        d = app.Documents.Open(docx, ReadOnly=True)
        try:
            d.ExportAsFixedFormat(pdf, 17)
        finally:
            d.Close(False)
            app.Quit()
    doc = fitz.open(pdf)
    out = []
    for pi, pg in enumerate(doc, 1):
        for b in pg.get_text("dict")["blocks"]:
            for ln in b.get("lines", []):
                txt = "".join(s["text"] for s in ln["spans"])
                if not txt.strip():
                    continue
                x0, _y0, x1, _y1 = ln["bbox"]
                out.append({"page": pi, "x0": x0, "x1": x1, "text": txt,
                            "y": round(ln["bbox"][1], 2)})
    return out


def norm(s):
    return "".join(ch for ch in (s or "") if not ch.isspace())


def match_report(docx, envs, label, quiet=False):
    """How many of Word's own lines does Oxi reproduce, character for character?

    The break is the thing under test, so score it directly: normalise both
    sides' line text and count the lines that agree.  A credit the render cannot
    honour shows up here as a line that starts one character late and never
    recovers.

    ★Counted as a per-page MULTISET intersection, not position by position: one
    extra line early on shifts every later line and would score a document that
    breaks identically from line 2 onwards as a total miss."""
    _recs, layout = run_oxi(docx, tag="m" + label, envs=envs)
    olines = oxi_lines(layout)
    wlines = word_lines(docx)
    per_page = {}
    tot = collections.Counter()
    obypage = collections.defaultdict(collections.Counter)
    for o in olines:
        obypage[o["page"]][norm(o["text"])] += 1
    wbypage = collections.defaultdict(collections.Counter)
    for w in wlines:
        wbypage[w["page"]][norm(w["text"])] += 1
    for pg in sorted(wbypage):
        w, o = wbypage[pg], obypage.get(pg, collections.Counter())
        hit = sum((w & o).values())
        per_page[pg] = (hit, sum(w.values()), sum(o.values()))
        tot["hit"] += hit
        tot["word"] += sum(w.values())
    if not quiet:
        print("%-28s lines matched %d/%d  %s"
              % (label, tot["hit"], tot["word"],
                 " ".join("p%d %d/%d(oxi %d)" % (p, v[0], v[1], v[2])
                          for p, v in sorted(per_page.items()))))
    return tot["hit"], tot["word"], per_page


def miss_report(docx, envs=""):
    """Word's line next to Oxi's, wherever the two disagree.

    Line agreement is the break's own score; this is what is left of it. Pages
    are walked in parallel and the first divergence on each page is what matters
    -- everything after it is that break repeated, so the head of each run is
    printed with the character Word moved and Oxi kept (or the reverse)."""
    _recs, layout = run_oxi(docx, tag="miss", envs=envs)
    obypage = collections.defaultdict(list)
    for o in oxi_lines(layout):
        obypage[o["page"]].append(o)
    wbypage = collections.defaultdict(list)
    for w in word_lines(docx):
        wbypage[w["page"]].append(w)
    for pg in sorted(wbypage):
        w, o = wbypage[pg], obypage.get(pg, [])
        for i in range(min(len(w), len(o))):
            wt, ot = norm(w[i]["text"]), norm(o[i]["text"])
            if wt == ot:
                continue
            # what the two did differently at the break: the common prefix, then
            # the first character each side put next
            n = 0
            while n < min(len(wt), len(ot)) and wt[n] == ot[n]:
                n += 1
            print("p%-2d L%-3d word %3d ch |%s|\n        %8s oxi  %3d ch |%s|   split after %d: word %r / oxi %r"
                  % (pg, i, len(wt), wt[-14:], "", len(ot), ot[-14:], n,
                     wt[n:n + 1], ot[n:n + 1]))


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    page = None
    for a in sys.argv[1:]:
        if a.startswith("--page"):
            page = int(a.split("=")[1] if "=" in a else args.pop())
    prefix = args[0] if args else "c7b923e5"
    docx = docx_for(prefix)

    if "--miss" in sys.argv:
        miss_report(docx)
        return

    if "--match" in sys.argv:
        arms = [("base", "")]
        for a in sys.argv[1:]:
            if a.startswith("--arm="):
                envs = a[len("--arm="):]
                arms.append((envs[:26], envs))
        for label, envs in arms:
            match_report(docx, envs, label)
        return

    recs, layout = run_oxi(docx)
    olines = oxi_lines(layout)
    wlines = word_lines(docx)

    # breaker record -> emitted line, by the leading characters of the line.
    # The breaker's head is 14 raw chars, the emitted text keeps its spaces, so
    # both sides are keyed on the first 8 non-space characters.
    def key(s):
        return norm(s)[:8]

    by_head = collections.defaultdict(list)
    for r in recs:
        by_head[key(r["head"])].append(r)
    wby = collections.defaultdict(list)
    for w in wlines:
        wby[key(w["text"])].append(w)

    rows = []
    for o in olines:
        k = key(o["text"])
        rec = by_head[k].pop(0) if by_head.get(k) else None
        wl = wby[k].pop(0) if wby.get(k) else None
        rows.append((o, rec, wl))

    if page:
        rows = [r for r in rows if r[0]["page"] == page]
        rows.sort(key=lambda t: t[0]["y"])
    else:
        rows = [t for t in rows if t[1] and
                t[0]["x1"] - t[0]["x0"] - t[1]["avail"] > 0.5]
        rows.sort(key=lambda t: -(t[0]["x1"] - t[0]["x0"] - t[1]["avail"]))
        rows = rows[:25]

    print("== %s ==  breaker vs emitter vs Word (pt)" % os.path.basename(docx))
    print("%-3s %-7s %-4s %-8s %-8s %-8s %-8s %-8s %-8s %s"
          % ("pg", "y", "nch", "avail", "frag_s", "nat_s", "emit_w",
             "emit-frag", "word_w", "head"))
    for o, rec, wl in rows:
        emit_w = o["x1"] - o["x0"]
        av = rec["avail"] if rec else float("nan")
        fsum = rec["frag_sum"] if rec else float("nan")
        nsum = rec["nat_sum"] if rec else float("nan")
        ww = (wl["x1"] - wl["x0"]) if wl else float("nan")
        print("%-3d %-7.1f %-4d %-8.2f %-8.2f %-8.2f %-8.2f %-8.2f %-8.2f %s"
              % (o["page"], o["y"], len(norm(o["text"])), av, fsum, nsum,
                 emit_w, emit_w - fsum, ww, norm(o["text"])[:16]))


if __name__ == "__main__":
    main()
