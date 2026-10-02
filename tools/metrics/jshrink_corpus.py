"""Compat-15 justify-shrink ALLOW, measured from every justified line of real docs.

Every justified line Word wrapped gives an interval on the shrink it may apply:

    used  = natural(line) - avail             (Word did shrink this much)  -> allow >= used
    need  = natural(line + space + next word) - avail  (Word refused it)   -> allow <  need

natural = font advances (fontTools, the font file Word embedded) at the run size;
avail   = the rendered span of the justified line (first glyph left .. last glyph
right), which is the line's available width.  A line is accepted only when its
words sit where one uniform per-space delta puts them (residual <= RES_TOL), which
checks the natural widths and that the line really is justified.

    python tools/metrics/jshrink_corpus.py pdf            # Word -> PDF for the selection
    python tools/metrics/jshrink_corpus.py lines          # -> pipeline_data/jshrink_corpus/lines.json
    python tools/metrics/jshrink_corpus.py check [model]  # violations of a model
"""
import glob, json, os, re, shutil, sys, tempfile, zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data/jshrink_corpus"
SEL = REPO / "pipeline_data/jshrink_corpus_sel.json"
RES_TOL = 0.6  # Arial lines in Word PDFs carry ~0.3pt position noise (legal__0027c9c1)


def select():
    rows = json.load(open(SEL))
    return [Path(p) for p, cm, _ in rows if cm == "15"]


def doc_id(p):
    return f"{p.parent.name}__{p.stem[:8]}"


def make_pdfs():
    import win32com.client
    (OUT / "pdf").mkdir(parents=True, exist_ok=True)
    w = win32com.client.DispatchEx("Word.Application")
    w.Visible = False
    w.DisplayAlerts = 0
    try:
        for p in select():
            pdf = OUT / "pdf" / (doc_id(p) + ".pdf")
            if pdf.exists():
                continue
            tmp = os.path.join(tempfile.mkdtemp(), p.name)
            shutil.copy(REPO / p, tmp)
            try:
                d = w.Documents.Open(tmp, ReadOnly=False, AddToRecentFiles=False)
                d.TrackRevisions = False
                d.Revisions.AcceptAll()
                d.SaveAs2(str(pdf), 17)
                d.Close(0)
                print("ok", pdf.name, flush=True)
            except Exception as e:
                print("ERR", p, e, flush=True)
    finally:
        w.Quit()


# ---- fonts ---------------------------------------------------------------
_FONTS = None


def font_index():
    global _FONTS
    if _FONTS is not None:
        return _FONTS
    from fontTools.ttLib import TTFont, TTCollection
    cache = OUT / "font_index.json"
    if cache.exists():
        idx = json.load(open(cache))
    else:
        idx = {}
        dirs = [r"C:\Windows\Fonts", os.path.expandvars(r"%LOCALAPPDATA%\Microsoft\Windows\Fonts")]
        for d in dirs:
            for f in glob.glob(os.path.join(d, "*")):
                if not f.lower().endswith((".ttf", ".otf", ".ttc")):
                    continue
                try:
                    fonts = TTCollection(f).fonts if f.lower().endswith(".ttc") else [TTFont(f, lazy=True)]
                except Exception:
                    continue
                for i, t in enumerate(fonts):
                    try:
                        ps = t["name"].getDebugName(6)
                    except Exception:
                        continue
                    if ps and ps not in idx:
                        idx[ps] = [f, i]
        OUT.mkdir(parents=True, exist_ok=True)
        json.dump(idx, open(cache, "w"))
    _FONTS = {"idx": idx, "loaded": {}}
    return _FONTS


def advances(psname):
    """char -> advance in em, or None if the font file is not found."""
    from fontTools.ttLib import TTFont
    F = font_index()
    name = psname.split("+", 1)[-1]
    if name in F["loaded"]:
        return F["loaded"][name]
    hit = F["idx"].get(name)
    res = None
    if hit:
        t = TTFont(hit[0], fontNumber=hit[1], lazy=True)
        upem = t["head"].unitsPerEm
        cmap = t.getBestCmap() or {}
        hm = t["hmtx"].metrics
        res = {chr(c): hm[g][0] / upem for c, g in cmap.items() if g in hm} or None
    F["loaded"][name] = res
    return res


# ---- lines ---------------------------------------------------------------
def page_lines(page):
    """Visual lines: list of dict(chars=[(c,x0,x1,font,size)], y, x0, x1)."""
    out = []
    for b in page.get_text("rawdict")["blocks"]:
        for l in b.get("lines", []):
            if abs(l["dir"][1]) > 1e-3:
                continue
            chars = []
            for s in l["spans"]:
                for c in s["chars"]:
                    chars.append((c["c"], c["bbox"][0], c["bbox"][2], s["font"], round(s["size"], 2), c["origin"][1]))
            while chars and chars[-1][0] == " ":
                chars.pop()
            if chars:
                out.append({"chars": chars, "y": chars[0][5], "x0": chars[0][1], "x1": chars[-1][2], "blk": b["number"]})
    return out


def first_word(line):
    cs = []
    for c in line["chars"]:
        if c[0] == " ":
            break
        cs.append(c)
        if c[0] in "-\u2010\u2013\u2014/":
            break
    return cs


_PARAS = {}


def para_texts(doc):
    """Whitespace-normalised text of every paragraph of the source docx."""
    if doc in _PARAS:
        return _PARAS[doc]
    import html
    rows = json.load(open(SEL))
    path = next((p for p, cm, _ in rows if doc_id(Path(p)) == doc), None)
    out = []
    if path:
        xml = zipfile.ZipFile(REPO / path).read("word/document.xml").decode("utf-8", "ignore")
        for p in re.findall(r"<w:p[ >].*?</w:p>", xml, re.S):
            # w:br / w:tab become markers so a forced break is never read as a refusal
            t = "".join(m.group(1) if m.group(1) is not None else ("\t" if "tab" in m.group(0) else "\n")
                        for m in re.finditer(r"<w:t(?:\s[^>]*)?>(.*?)</w:t>|<w:tab/>|<w:br\b[^>]*/>|<w:cr/>", p, re.S))
            t = html.unescape(t).strip()  # keep double spaces: "end.  Next" is a 2-space gap
            if t:
                out.append(t)
    _PARAS[doc] = out
    return out


_KERN = {}


def doc_kerned(doc):
    if doc not in _KERN:
        rows = json.load(open(SEL))
        path = next((p for p, cm, _ in rows if doc_id(Path(p)) == doc), None)
        k = False
        if path:
            z = zipfile.ZipFile(REPO / path)
            for part in ("word/document.xml", "word/styles.xml"):
                if part in z.namelist() and "<w:kern " in z.read(part).decode("utf-8", "ignore"):
                    k = True
        _KERN[doc] = k
    return _KERN[doc]


def next_chunk(doc, text, nxt_word):
    """The chunk Word had to fit after `text` to pull `nxt_word` up: the source
    paragraph's text from the following word up to the next breakable gap.
    Spaces break; NBSP and '/' do not (Word keeps "family/carer" and
    "N$113\u00a0306" whole); a hyphen breaks after itself. None when the
    junction is not found or is ambiguous (several paragraphs, different chunks)."""
    probe = text + " " + nxt_word
    chunks = set()
    for p in para_texts(doc):
        i = p.find(probe)
        while i >= 0:
            rest = p[i + len(text) + 1:]
            m = re.match(r"[^ \t\n]*?(?:-|\u2010|\u2013(?=[^ \t\n])|$|(?=[ \t\n]))", rest)
            chunks.add(m.group(0) if m and m.group(0) else rest.split(" ")[0])
            i = p.find(probe, i + 1)
    return next(iter(chunks)) if len(chunks) == 1 else None


def analyse_line(line, nxt, xs_right):
    chars = line["chars"]
    fonts = {(c[3], c[4]) for c in chars}
    if len(fonts) != 1:
        return None
    font, size = next(iter(fonts))
    # Word's PDF writes Tf sizes a hair off (Corbel sz=20 -> 9.96, Calibri 11 -> 11.04)
    if abs(size - round(size * 2) / 2) < 0.08:
        size = round(size * 2) / 2
    adv = advances(font)
    if not adv or " " not in adv:
        return None
    if any(c[0] not in adv for c in chars):
        return None
    sp = adv[" "] * size
    nsp = sum(1 for c in chars if c[0] == " ")
    if nsp < 2 or nsp > 40 or "  " in "".join(c[0] for c in chars):
        return None
    # right edge must be shared with >= 3 other lines (the column edge)
    if sum(1 for x in xs_right if abs(x - line["x1"]) < 0.25) < 4:
        return None
    natural = sum(adv[c[0]] * size for c in chars)
    avail = line["x1"] - line["x0"]
    delta = (avail - natural) / nsp
    # residual: every glyph must sit at x0 + nat_prefix + k*delta
    res = 0.0
    acc = 0.0
    k = 0
    for c in chars:
        pred = line["x0"] + acc + k * delta
        if c[0] != " ":
            res = max(res, abs(c[1] - pred))
        acc += adv[c[0]] * size
        if c[0] == " ":
            k += 1
    if res > RES_TOL or abs(delta) < 0.02:
        return None
    words = "".join(c[0] for c in chars).split(" ")
    last = words[-1]
    rec = {"font": font.split("+", 1)[-1], "size": size, "nsp": nsp, "sp": sp,
           "natural": natural, "avail": avail, "used": natural - avail,
           "last": last, "last_w": sum(adv[ch] * size for ch in last),
           "text": "".join(c[0] for c in chars), "res": res}
    if nxt is not None:
        fw = first_word(nxt)
        if fw and all((c[3], c[4]) == (font, size) and c[0] in adv for c in fw):
            w = "".join(c[0] for c in fw)
            rec["next"] = w
            rec["next_w"] = sum(adv[ch] * size for ch in w)
            rec["need"] = natural + sp + rec["next_w"] - avail
    return rec


def collect():
    import fitz
    recs = []
    for pdf in sorted((OUT / "pdf").glob("*.pdf")):
        try:
            doc = fitz.open(pdf)
        except Exception:
            continue
        for pno, page in enumerate(doc):
            ls = page_lines(page)
            xs = [l["x1"] for l in ls]
            for i, l in enumerate(ls):
                nxt = None
                if i + 1 < len(ls):
                    n = ls[i + 1]
                    if 0 < n["y"] - l["y"] < 3 * max(c[4] for c in l["chars"]) and n["x0"] < l["x1"] and n["x1"] > l["x0"]:
                        nxt = n
                r = analyse_line(l, nxt, xs)
                if r:
                    r["doc"] = pdf.stem
                    r["page"] = pno + 1
                    r["kern"] = doc_kerned(r["doc"])
                    # the refused chunk is read from the source paragraph (NBSP / slash
                    # joins, forced breaks, double spaces are invisible in the PDF line)
                    if "need" in r:
                        chunk = next_chunk(r["doc"], r["text"], r["next"])
                        adv = advances(r["font"])
                        if chunk and all(ch in adv for ch in chunk):
                            r["next"] = chunk
                            r["next_w"] = sum(adv[ch] * r["size"] for ch in chunk)
                            r["need"] = r["natural"] + r["sp"] + r["next_w"] - r["avail"]
                        else:
                            for k in ("next", "next_w", "need"):
                                r.pop(k)
                    recs.append(r)
    json.dump(recs, open(OUT / "lines.json", "w"), indent=0)
    n_need = sum(1 for r in recs if "need" in r)
    print(f"{len(recs)} justified lines ({n_need} with next word), "
          f"{len({r['doc'] for r in recs})} docs, {sum(1 for r in recs if r['used'] > 0)} shrunk")


# ---- models --------------------------------------------------------------
def m_s825_s1475(r, for_need=False):
    """S825 space cap + S1475 last-word cap. For a refusal the candidate word
    becomes the line's last word and adds one more compressible space."""
    if for_need:
        return min(0.25 * (r["nsp"] + 1) * r["sp"], 0.35 * (r["next_w"] + r["sp"]))
    return min(0.25 * r["nsp"] * r["sp"], 0.35 * (r["last_w"] + r["sp"]))


MODELS = {"s825_s1475": m_s825_s1475}
TOL = 0.3  # PDF glyph positions carry up to ~0.3pt noise (see RES_TOL)


def check(name):
    recs = json.load(open(OUT / "lines.json"))
    f = MODELS[name]
    lo = [r for r in recs if r["used"] > f(r) + TOL]
    hi = [r for r in recs if "need" in r and r["need"] <= f(r, True) - TOL]
    sh = [r for r in recs if r["used"] > 0.3]
    import statistics
    print(f"shrunk lines {len(sh)}: used/cap quantiles "
          f"{[round(q, 3) for q in statistics.quantiles([r['used'] / f(r) for r in sh], n=20)[::4]]} max {max(r['used'] / f(r) for r in sh):.3f}")
    print(f"{name}: {len(recs)} lines; under-allow (Word shrank more) {len(lo)}; "
          f"over-allow (Word refused a fit the model allows) {len(hi)}")
    for tag, rs in (("UNDER", lo), ("OVER", hi)):
        for r in rs[:15]:
            print(f"  {tag} {r['doc']} p{r['page']} {r['font']} {r['size']} nsp={r['nsp']} res={r['res']:.2f} used={r['used']:.2f} cap={f(r):.2f} "
                  f"need={r.get('need', float('nan')):.2f} cap_next={f(r, True) if 'need' in r else float('nan'):.2f} "
                  f"last={r['last']!r} next={r.get('next')!r} | {r['text'][-36:]}")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    cmd = sys.argv[1]
    if cmd == "pdf":
        make_pdfs()
    elif cmd == "lines":
        collect()
    elif cmd == "check":
        check(sys.argv[2] if len(sys.argv) > 2 else "s825_s1475")
