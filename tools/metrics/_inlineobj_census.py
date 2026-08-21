# -*- coding: utf-8 -*-
"""Census: body paragraphs that carry BOTH an inline OBJECT (w:object OLE /
inline visual vector group) AND an inline PICTURE (wp:inline drawing).

That pair is the shape S854 never covered: the object is marked on its run
(inline_object_extent, S851/S974/S839) while the picture is routed to
inline_img_runs, and with n_images == 1 the picture fell to SPLIT-BLOCK = its
own line. correspondence__000407cd is the pcd witness (Word 1 page / Oxi 2).

  python tools/metrics/_inlineobj_census.py
"""
import re
import sys
import zipfile
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
ROOTS = [
    ("golden", REPO / "tools" / "golden-test" / "documents" / "docx"),
    ("corp_en", REPO / "pipeline_data" / "docx_corpus" / "en"),
    ("corp_ja", REPO / "pipeline_data" / "docx_corpus" / "ja"),
    ("real_en", REPO / "tools" / "golden-test" / "real_en"),
]

PARA_RE = re.compile(r"<w:p[ >].*?</w:p>", re.S)


def para_kinds(p: str):
    """(n_inline_pic, n_object, n_anchor) for one paragraph's XML."""
    n_pic = len(re.findall(r"<wp:inline[ >]", p))
    n_obj = len(re.findall(r"<w:object[ >]", p))
    n_anchor = len(re.findall(r"<wp:anchor[ >]", p))
    return n_pic, n_obj, n_anchor


rows = []
for label, root in ROOTS:
    if not root.is_dir():
        continue
    for f in sorted(root.rglob("*.docx")):
        try:
            with zipfile.ZipFile(f) as z:
                if "word/document.xml" not in z.namelist():
                    continue
                x = z.read("word/document.xml").decode("utf8", "replace")
        except Exception:  # noqa: BLE001
            continue
        hits = []
        for i, p in enumerate(PARA_RE.findall(x)):
            n_pic, n_obj, n_anchor = para_kinds(p)
            if n_pic >= 1 and n_obj >= 1:
                text = "".join(re.findall(r"<w:t[^>]*>(.*?)</w:t>", p, re.S))
                hits.append((i, n_pic, n_obj, n_anchor, text.strip()[:30]))
        if hits:
            rows.append((label, f, hits))

print("docs with an inline-object + inline-picture paragraph: %d" % len(rows))
for label, f, hits in rows:
    print("  [%s] %s/%s" % (label, f.parent.name, f.stem))
    for i, n_pic, n_obj, n_anchor, t in hits:
        print("      para %-4d pic=%d obj=%d anchor=%d visible_text=%r"
              % (i, n_pic, n_obj, n_anchor, t))
