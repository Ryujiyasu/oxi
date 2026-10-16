# -*- coding: utf-8 -*-
"""Face census: which face did Word draw each character in, and which did Oxi?

The Word PDF embeds the face Word actually used for every span; Oxi's glyph
dump names the family it resolved.  The two spell names differently (subset
prefixes, PostScript names, «Family,Bold» synthetic styles, ＭＳ 明朝 vs
MS-Mincho), so both are mapped to a canonical family read from the font
files' own name tables: Windows\\Fonts, the per-user font directory and the
Office cloud cache (the inventory Word draws from on this machine).  The
inventory itself is written next to the result, because Word's choice
depends on it.

Characters are paired per page by identity and position (same char, nearest
origin within 4pt vertically and 12pt horizontally), so a reflowed line still
pairs while a face that changes the advance does not lose the match.

    CENSUS_BLIND=blindG_20260929 python tools/metrics/face_census.py run <renderer.exe> <tag> [pages]
    python tools/metrics/face_census.py report <tag> [<tag2> ...]

Output: pipeline_data/face_census/<tag>/{pairs.json, inventory.json}
"""
import collections, glob, json, os, re, subprocess, sys, tempfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "face_census"


def font_dirs():
    local = os.environ.get("LOCALAPPDATA", "")
    win = os.environ.get("WINDIR", r"C:\Windows")
    return [("cloud", os.path.join(local, r"Microsoft\FontCache\4\CloudFonts")),
            ("user", os.path.join(local, r"Microsoft\Windows\Fonts")),
            ("system", os.path.join(win, "Fonts"))]


def build_inventory():
    """name (lowercased, any spelling) -> canonical family; plus the file list."""
    from fontTools.ttLib import TTFont, TTCollection
    names, files = {}, []
    for src, d in font_dirs():
        if not os.path.isdir(d):
            continue
        paths = glob.glob(os.path.join(d, "*")) + glob.glob(os.path.join(d, "*", "*"))
        for p in paths:
            ext = p.lower().rsplit(".", 1)[-1]
            if ext not in ("ttf", "otf", "ttc"):
                continue
            try:
                fonts = TTCollection(p, lazy=True).fonts if ext == "ttc" else [TTFont(p, lazy=True)]
            except Exception:
                continue
            for f in fonts:
                try:
                    nt = f["name"]
                except Exception:
                    continue
                def get(i, lang=None):
                    out = []
                    for r in nt.names:
                        if r.nameID == i:
                            try:
                                out.append((r.langID, r.toUnicode()))
                            except Exception:
                                pass
                    return out
                fam = [v for l, v in get(16)] or [v for l, v in get(1)]
                en = ([v for l, v in get(16) if l == 0x409] or [v for l, v in get(1) if l == 0x409] or fam or ["?"])[0]
                canon = en.strip()
                for v in [v for l, v in get(1)] + [v for l, v in get(16)] + [v for l, v in get(4)] + [v for l, v in get(6)]:
                    names.setdefault(v.strip().lower(), canon)
                files.append({"src": src, "file": p, "family": canon,
                              "style": ([v for l, v in get(2) if l == 0x409] or ["?"])[0]})
    return names, files


def canon(name, names):
    if name and name.startswith("="):
        # a family read from an embedded program: map it through the inventory
        # so localized / English spellings meet, else keep it as is
        return names.get(name[1:].strip().lower(), name[1:].strip())
    n = re.sub(r"^[A-Z]{6}\+", "", name or "").strip()
    for cand in (n, n.split(",")[0], re.sub(r"-(Bold|Italic|BoldItalic|Regular|Light|Medium|Semibold|BoldMT|ItalicMT|BoldItalicMT)$", "", n.split(",")[0]),
                 re.sub(r"MT$", "", n.split(",")[0])):
        k = cand.strip().lower()
        if k in names:
            return names[k]
    return "?" + n.split(",")[0]


def embedded_families(doc, page):
    """PDF base-font name (subset prefix stripped) -> the family in the EMBEDDED
    program's own name table.  Word labels a substituted face with the name the
    document asked for (en/creative__03ea0a4c: «Times» whose program is Times New
    Roman), so the label is not the face; and fitz truncates span font names to
    24 characters.  Unreadable programs keep the label."""
    import io
    from fontTools.ttLib import TTFont
    # One label can name SEVERAL programs: Word draws the CJK characters of a
    # Calibri run in its fallback face and embeds that program under the label
    # «Calibri» too (ja/reports__59eadfc4: «Calibri» = Calibri + MS PGothic).
    # Keep every program with the characters it maps, and pick per character.
    out = {}
    for f in page.get_fonts():
        label = re.sub(r"^[A-Z]{6}\+", "", f[3])
        try:
            buf = doc.extract_font(f[0])[3]
            t = TTFont(io.BytesIO(buf)); nt = t["name"]
            fam = None
            for nid in (16, 1):
                for r in nt.names:
                    if r.nameID == nid and r.langID == 0x409:
                        fam = r.toUnicode(); break
                if fam:
                    break
            fam = fam or nt.getDebugName(1)
            try:
                cps = set((t.getBestCmap() or {}).keys())
            except Exception:
                cps = set()
        except Exception:
            fam, cps = None, set()
        if fam:
            out.setdefault(label, []).append(("=" + fam.strip(), cps))
    return out


def pick_program(progs, label, ch):
    """The program among those sharing a label that maps `ch`; prefer the one
    whose family is the label itself when both do."""
    if not progs:
        return None
    have = [p for p in progs if ord(ch) in p[1]]
    if not have:
        have = [p for p in progs if not p[1]] or progs
    base = label.split(",")[0].split("-")[0].lower()
    for fam, _ in have:
        if fam[1:].replace(" ", "").lower().startswith(base.replace(" ", "")):
            return fam
    return have[0][0]


def word_chars(pdf, pages):
    import fitz
    doc = fitz.open(pdf); out = []
    for pi in range(min(pages, len(doc))):
        emb = embedded_families(doc, doc[pi])
        def real(span_font, ch):
            n = re.sub(r"^[A-Z]{6}\+", "", span_font)
            progs = emb.get(n)
            if progs is None:
                keys = [k for k in emb if k.startswith(n)]
                progs = [p for k in keys for p in emb[k]] if keys else None
            return pick_program(progs, n, ch) if progs else span_font
        for b in doc[pi].get_text("rawdict")["blocks"]:
            for l in b.get("lines", []):
                for s in l["spans"]:
                    for c in s["chars"]:
                        if c["c"].strip():
                            out.append((pi, c["c"], c["origin"][0], c["origin"][1], real(s["font"], c["c"])))
    return out


def oxi_chars(exe, docx, pages):
    t = tempfile.mkdtemp(); go = os.path.join(t, "g.json")
    subprocess.run([exe, str(docx), os.path.join(t, "p"), "--dump-glyphs=" + go], capture_output=True, timeout=900)
    if not os.path.exists(go):
        return None
    d = json.load(open(go, encoding="utf-8")); out = []
    for pi, pg in enumerate(d["pages"][:pages]):
        for g in pg["glyphs"]:
            if g["char"].strip():
                out.append((pi, g["char"], g["x"], g["baseline"], g.get("font_family") or ""))
    return out


def pair(W, O):
    by = collections.defaultdict(list)
    for o in O:
        by[(o[0], o[1])].append(o)
    used = set(); res = []
    for w in W:
        best = None
        for o in by.get((w[0], w[1]), []):
            if id(o) in used:
                continue
            dy, dx = abs(o[3] - w[3]), abs(o[2] - w[2])
            if dy <= 4 and dx <= 12 and (best is None or dy + dx * 0.2 < best[0]):
                best = (dy + dx * 0.2, o)
        if best:
            used.add(id(best[1])); res.append((w, best[1]))
    return res


def run(exe, tag, pages):
    import line_census as L
    names, files = build_inventory()
    od = OUT / tag; od.mkdir(parents=True, exist_ok=True)
    (od / "inventory.json").write_text(json.dumps(files, ensure_ascii=False), encoding="utf-8")
    result = {}
    for key, docx in L.docs():
        pdf = L.OUT / "pdf" / f"{key}.pdf"
        if not pdf.exists():
            continue
        W = word_chars(str(pdf), pages); O = oxi_chars(exe, docx, pages)
        if O is None:
            result[key] = {"error": "no dump"}; continue
        P = pair(W, O)
        cnt = collections.Counter()
        for w, o in P:
            cnt[(canon(w[4], names), canon(o[4], names))] += 1
        result[key] = {"word_chars": len(W), "paired": len(P),
                       "pairs": [[a, b, n] for (a, b), n in cnt.most_common()]}
        mism = sum(n for (a, b), n in cnt.items() if a != b)
        print(f"  {key:44s} paired {len(P):5d}/{len(W):5d}  face-mismatch {mism}", flush=True)
    (od / "pairs.json").write_text(json.dumps(result, ensure_ascii=False), encoding="utf-8")


def report(tags):
    agg = collections.Counter(); docs = collections.defaultdict(set); tot = mis = 0
    for tag in tags:
        R = json.loads((OUT / tag / "pairs.json").read_text(encoding="utf-8"))
        for key, r in R.items():
            for a, b, n in r.get("pairs", []):
                tot += n
                if a != b:
                    mis += n; agg[(a, b)] += n; docs[(a, b)].add(key)
    print(f"paired chars {tot}, face mismatches {mis} ({100.0 * mis / max(tot, 1):.2f}%) in {len(set().union(*docs.values())) if docs else 0} docs")
    for (a, b), n in agg.most_common(60):
        print(f"  {n:6d}  word {a:32s} oxi {b:32s} docs {len(docs[(a, b)]):3d}  e.g. {sorted(docs[(a, b)])[0]}")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    if sys.argv[1] == "run":
        run(sys.argv[2], sys.argv[3], int(sys.argv[4]) if len(sys.argv) > 4 else 3)
    else:
        report(sys.argv[2:])
