# -*- coding: utf-8 -*-
"""Which documents does the break-time 約物 credit help, and which does it hurt?

The credit models Word compressing a mark to pull one more character onto the
line.  c7b923e5 draws lines up to 26pt past its right margin with it on, and its
Word PDF shows every mark at its natural advance -- no compression at all.  The
candidate discriminator is the face: a PROPORTIONAL Japanese font (MS PGothic,
MS PMincho, HGPGothicM) already ships 、。（）at about half an em, so there is no
aki left for Word to take, while a monospace face (MS Mincho, MS Gothic) carries
a full em with the blank half Word can remove.

For every corpus document that has Word's own export next to it, print the
dominant CJK face and how many of Word's lines Oxi reproduces with the credit on
and with it off.

    python _cb_font_census.py            # every document with a Word PDF
    python _cb_font_census.py c7b tokyo  # only those prefixes
"""
import collections
import os
import re
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
os.environ.setdefault("PYTHONIOENCODING", "utf-8")
sys.stdout.reconfigure(encoding="utf-8")

import _cb_budget as B  # noqa: E402

PROPORTIONAL = ("ＭＳ Ｐゴシック", "ＭＳ Ｐ明朝", "MS PGothic", "MS PMincho",
                "HGPGothicM", "HGP", "Ｐゴシック", "Ｐ明朝", "メイリオ", "Meiryo")
NO_CREDIT = "OXI_S475_SOLO=0,OXI_S475_PAIR=0,OXI_S475_OPEN=0"


def dominant_cjk_font(docx):
    """The eastAsia face most runs name, falling back to the default style."""
    z = zipfile.ZipFile(docx)
    try:
        x = z.read("word/document.xml").decode("utf-8", "replace")
    except KeyError:
        return "?"
    names = re.findall(r'<w:rFonts[^>]*w:eastAsia="([^"]+)"', x)
    if not names:
        try:
            s = z.read("word/styles.xml").decode("utf-8", "replace")
            names = re.findall(r'<w:rFonts[^>]*w:eastAsia="([^"]+)"', s)
        except KeyError:
            pass
    if not names:
        return "?"
    return collections.Counter(names).most_common(1)[0][0]


def word_mark_advance(docx):
    """Word's own median advance for 、 and 。, and the run font size it drew at.

    The point of the column is to see whether Word COMPRESSED the marks in this
    document at all: an advance at the face's natural width means Word paid for
    the line some other way and the break credit models something it never did.
    """
    import fitz
    rt = docx[:-5] + "_rt.pdf"
    if not os.path.exists(rt):
        return None
    adv = collections.defaultdict(list)
    size = []
    for pg in fitz.open(rt):
        for b in pg.get_text("rawdict")["blocks"]:
            for ln in b.get("lines", []):
                for s in ln["spans"]:
                    chars = s["chars"]
                    for i, c in enumerate(chars[:-1]):
                        if c["c"] in "、。":
                            adv[c["c"]].append(chars[i + 1]["bbox"][0] - c["bbox"][0])
                            size.append(s["size"])
    if not adv:
        return None
    med = {k: sorted(v)[len(v) // 2] for k, v in adv.items()}
    fs = sorted(size)[len(size) // 2] if size else 0.0
    return med, fs


def in_scope(docx):
    """compat15 + compressPunctuation + a justified default: the s475 arm."""
    z = zipfile.ZipFile(docx)
    try:
        st = z.read("word/settings.xml").decode("utf-8", "replace")
    except KeyError:
        return False
    if "compressPunctuation" not in st:
        return False
    m = re.search(r'w:name="compatibilityMode"[^>]*w:val="(\d+)"', st)
    if not m or int(m.group(1)) < 15:
        return False
    body = z.read("word/document.xml").decode("utf-8", "replace")
    styles = z.read("word/styles.xml").decode("utf-8", "replace")
    return 'w:val="both"' in body or 'w:val="both"' in styles


def main():
    prefixes = [a for a in sys.argv[1:] if not a.startswith("-")]
    docs = []
    for f in sorted(os.listdir(B.DOCS)):
        if not f.endswith(".docx") or f.startswith("~$"):
            continue
        path = os.path.join(B.DOCS, f)
        if not os.path.exists(path[:-5] + "_rt.pdf"):
            continue
        if prefixes and not any(f.startswith(p) for p in prefixes):
            continue
        docs.append(path)
    print("%-30s %-12s %-5s %-18s %-10s %-10s %-10s %s"
          % ("document", "eastAsia", "prop", "Word 、/。 vs em", "pre-S1167",
             "S1167", "no-credit", "delta"))
    for path in docs:
        name = os.path.basename(path)[:-5]
        face = dominant_cjk_font(path)
        prop = any(p in face for p in PROPORTIONAL)
        mk = word_mark_advance(path)
        if mk:
            med, fs = mk
            marks = "%.2f/%.2f of %.1f" % (med.get("、", 0), med.get("。", 0), fs)
        else:
            marks = "-"
        if not in_scope(path):
            print("%-30s %-12s %-5s %-18s (out of the s475 scope)"
                  % (name[:30], face[:12], prop, marks))
            continue
        try:
            pre = B.match_report(path, "OXI_S1167_DISABLE=1", "pre", quiet=True)
            now = B.match_report(path, "", "on", quiet=True)
            off = B.match_report(path, NO_CREDIT, "off", quiet=True)
        except Exception as e:  # a document the renderer or the join cannot take
            print("%-30s %-12s %-5s %-18s failed: %s"
                  % (name[:30], face[:12], prop, marks, e))
            continue
        print("%-30s %-12s %-5s %-18s %4d/%-5d %4d/%-5d %4d/%-5d %+d"
              % (name[:30], face[:12], "prop" if prop else "mono", marks,
                 pre[0], pre[1], now[0], now[1], off[0], off[1],
                 now[0] - pre[0]))


if __name__ == "__main__":
    main()
