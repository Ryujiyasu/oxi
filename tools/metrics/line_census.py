# -*- coding: utf-8 -*-
"""Line census: page-1 lines of Word (PDF) vs Oxi (--dump-layout) over a doc set.

The stop criterion (2) of 2026-10-02: on the Blind-G 100 documents, no page-1
line whose baseline is more than 2pt from Word's.

    python tools/metrics/line_census.py pdf                  # Word -> PDF (cached)
    python tools/metrics/line_census.py run <renderer.exe> <tag>
    python tools/metrics/line_census.py report <tag> [N]     # worst N docs

Pairing: every Word line (chars grouped by baseline within 1.5pt) is matched
to the Oxi line with the most shared leading text within +-40pt; Oxi baselines
are the rendered glyph tops + the used face's winAscent (--dump-glyphs).
"""
import glob, json, os, re, shutil, subprocess, sys, tempfile
from pathlib import Path
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "line_census"
BLIND = REPO / "pipeline_data" / "blindG_20260929"


def docs():
    out = []
    for lang in ("en", "ja"):
        for j in sorted((BLIND / lang / "word").glob("*.json")):
            typ, sha = j.stem.split("__", 1)
            p = REPO / "pipeline_data" / "docx_corpus" / lang / typ / f"{sha}.docx"
            if p.exists():
                out.append((f"{lang}__{typ}__{sha[:8]}", p))
    return out


def make_pdfs():
    import win32com.client
    (OUT / "pdf").mkdir(parents=True, exist_ok=True)
    w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
    try:
        for key, p in docs():
            pdf = OUT / "pdf" / f"{key}.pdf"
            if pdf.exists():
                continue
            tmp = os.path.join(tempfile.mkdtemp(), p.name); shutil.copy(p, tmp)
            try:
                d = w.Documents.Open(tmp, ReadOnly=False, AddToRecentFiles=False)
                d.TrackRevisions = False; d.Revisions.AcceptAll()
                d.SaveAs2(str(pdf), 17); d.Close(0)
                print("ok", key, flush=True)
            except Exception as e:
                print("ERR", key, e, flush=True)
    finally:
        w.Quit()


def group_lines(items):
    """items: (baseline, x0, x1, text_char) -> lines sorted by y."""
    items = sorted(items, key=lambda g: (round(g[0]), g[1]))
    out = []
    for y, x0, x1, ch in items:
        if out and abs(out[-1]["y"] - y) < 1.5:
            L = out[-1]; L["x1"] = max(L["x1"], x1); L["x0"] = min(L["x0"], x0); L["t"] += ch
        else:
            out.append({"y": y, "x0": x0, "x1": x1, "t": ch})
    for L in out:
        L["t"] = L["t"].replace(" ", "")
    return out


def word_lines(pdf):
    import fitz
    page = fitz.open(pdf)[0]
    items = []
    for b in page.get_text("rawdict")["blocks"]:
        for l in b.get("lines", []):
            if abs(l["dir"][1]) > 1e-3:
                continue
            for s in l["spans"]:
                for c in s["chars"]:
                    if c["c"].strip():
                        items.append((c["origin"][1], c["origin"][0], c["bbox"][2], c["c"]))
    return group_lines(items)


_ASC = None


def win_ascent(family):
    """Ascent ratio the GDI renderer draws with (tmAscent = OS/2 winAscent)."""
    global _ASC
    if _ASC is None:
        _ASC = {}
        for e in json.load(open(REPO / "crates/oxidocs-core/src/font/data/font_metrics_compact.json", encoding="utf-8")):
            _ASC[e["family"]] = e["win_ascent"] / e["units_per_em"]
        # the renderer's localized face names
        _ASC.setdefault("游明朝", _ASC.get("Yu Mincho Regular", 0.995))
        _ASC.setdefault("游ゴシック", _ASC.get("Yu Gothic Regular", 1.0))
        _ASC.setdefault("ＭＳ 明朝", _ASC.get("MS Mincho", 0.859))
        _ASC.setdefault("ＭＳ ゴシック", _ASC.get("MS Gothic", 0.859))
        _ASC.setdefault("ＭＳ Ｐ明朝", _ASC.get("MS PMincho", 0.859))
        _ASC.setdefault("ＭＳ Ｐゴシック", _ASC.get("MS PGothic", 0.859))
    if family in _ASC:
        return _ASC[family]
    for k, v in _ASC.items():
        if family and (k.startswith(family) or family.startswith(k)):
            return v
    return 0.859


def oxi_lines(exe, docx):
    """2026-10-02: from --dump-glyphs (rendered glyph box tops + the face the
    renderer used); baseline = top + winAscent(face) * size. The earlier
    --dump-layout estimate used a flat 0.859 (MS Mincho's ascent), which read
    Cambria / Calibri / Yu Mincho lines 1-2pt off in either direction."""
    t = tempfile.mkdtemp(); out = os.path.join(t, "g.json")
    subprocess.run([exe, str(docx), os.path.join(t, "p"), "--dump-glyphs=" + out], capture_output=True, timeout=600)
    if not os.path.exists(out):
        return None
    d = json.load(open(out, encoding="utf-8"))
    if not d.get("pages"):
        return []
    items = []
    for g in d["pages"][0]["glyphs"]:
        fs = g.get("font_size", 0) or 0
        base = g["top"] + win_ascent(g.get("font_family", "")) * fs
        items.append((base, g["x"], g["x"] + fs, g["char"]))
    items.sort(key=lambda it: (round(it[0]), it[1]))
    return group_lines(items)


def pair(W, O):
    """For each Word line the best Oxi line: shared leading chars, within 40pt."""
    rows = []
    used = set()
    for w in W:
        best, score = None, 0
        for i, o in enumerate(O):
            if i in used or abs(o["y"] - w["y"]) > 40:
                continue
            k = 0
            for a, b in zip(w["t"], o["t"]):
                if a != b:
                    break
                k += 1
            if k > score and k >= min(3, len(w["t"])):
                best, score = i, k
        if best is not None:
            used.add(best); o = O[best]
            rows.append({"wt": w["t"][:20], "wy": w["y"], "oy": o["y"], "dy": o["y"] - w["y"], "dx": o["x0"] - w["x0"], "k": score})
        else:
            rows.append({"wt": w["t"][:20], "wy": w["y"], "oy": None, "dy": None, "dx": None, "k": 0})
    return rows


def run(exe, tag):
    (OUT / tag).mkdir(parents=True, exist_ok=True)
    summary = {}
    for key, p in docs():
        pdf = OUT / "pdf" / f"{key}.pdf"
        if not pdf.exists():
            summary[key] = {"error": "no pdf"}; continue
        try:
            W = word_lines(pdf); O = oxi_lines(exe, p)
        except Exception as e:
            summary[key] = {"error": str(e)[:120]}; continue
        if O is None:
            summary[key] = {"error": "oxi failed"}; continue
        rows = pair(W, O)
        dys = [r["dy"] for r in rows if r["dy"] is not None]
        bad = [r for r in rows if r["dy"] is not None and abs(r["dy"]) > 2.0]
        first = next((r for r in rows if r["dy"] is not None and abs(r["dy"]) > 2.0), None)
        summary[key] = {"n_word": len(W), "n_oxi": len(O), "paired": len(dys), "unpaired": len(rows) - len(dys),
                        "n_dy_gt2": len(bad), "max_dy": max((abs(d) for d in dys), default=0.0),
                        "first_bad": first, "rows": rows}
        print(f"{key:44} lines {len(W):3} paired {len(dys):3} |dy|>2: {len(bad):3} max {summary[key]['max_dy']:.2f}", flush=True)
    json.dump(summary, open(OUT / tag / "summary.json", "w", encoding="utf-8"), ensure_ascii=False, indent=0)
    report(tag, 15)


def report(tag, n=15):
    s = json.load(open(OUT / tag / "summary.json", encoding="utf-8"))
    ok = [k for k, v in s.items() if "error" not in v]
    tot_bad = sum(s[k]["n_dy_gt2"] for k in ok)
    clean = sum(1 for k in ok if s[k]["n_dy_gt2"] == 0)
    print(f"\n{tag}: {len(ok)} docs measured ({len(s) - len(ok)} errors); clean docs {clean}/{len(ok)}; lines |dy|>2 total {tot_bad}")
    worst = sorted(ok, key=lambda k: (-s[k]["n_dy_gt2"], -s[k]["max_dy"]))[:n]
    for k in worst:
        v = s[k]; fb = v["first_bad"]
        print(f"  {k:44} |dy|>2 {v['n_dy_gt2']:3}/{v['paired']:3} max {v['max_dy']:6.2f}  first: {fb['wt'] if fb else ''!r} dy={fb['dy'] if fb else 0:+.2f} at y={fb['wy'] if fb else 0:.1f}")


if __name__ == "__main__":
    cmd = sys.argv[1]
    if cmd == "pdf":
        make_pdfs()
    elif cmd == "run":
        run(os.path.abspath(sys.argv[2]), sys.argv[3])
    elif cmd == "report":
        report(sys.argv[2], int(sys.argv[3]) if len(sys.argv) > 3 else 15)
