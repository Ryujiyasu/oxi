"""Can each corpus document be rebuilt with the editing features Oxi has today?

Static census. For every .docx it lists the WordprocessingML features the
document uses and checks each against what the creation code can write
(`crates/oxidocs-core/src/editor.rs` DocxEdit / RunProps / ParaProps and the
`build_docx*` builders in lib.rs). A document is "rebuildable" in a mode when
every feature it uses is writable in that mode.

Two modes, because Oxi has two ways to make a document:

  template  The original is the template (build_docx_with_template): styles,
            numbering definitions, theme, settings, headers/footers and the
            FINAL sectPr are carried over. Only the body has to be expressible
            through DocxEdit.
  blank     From create_blank_docx: everything must be expressible, so page
            setup, headers/footers, footnotes, custom styles, numbering and
            document defaults all count as well.

A feature is a key such as `pPr/numPr`, `rPr/caps`, `tcPr/gridSpan`,
`drawing:anchor`, `pPr/spacing@lineRule=exact`. The census does not render;
it only says whether the editor can SAY the document. A rebuild-and-render
check is the next layer.

Usage:
    python tools/metrics/edit_coverage_census.py [corpus_dir] [out_dir]
Defaults: tools/golden-test/documents/docx  pipeline_data/edit_coverage
"""
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

import json
import re
import sys
import zipfile
from collections import Counter, defaultdict
from pathlib import Path

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"

PREFIX = {
    W: "w", MC: "mc", WP: "wp", A: "a", PIC: "pic",
    "urn:schemas-microsoft-com:vml": "v",
    "urn:schemas-microsoft-com:office:office": "o",
    "http://schemas.openxmlformats.org/officeDocument/2006/math": "m",
    "http://schemas.microsoft.com/office/word/2010/wordml": "w14",
}


def qn(tag):
    if not isinstance(tag, str):
        return None  # comments, PIs
    if tag.startswith("{"):
        ns, local = tag[1:].split("}", 1)
        return local if ns == W else f"{PREFIX.get(ns, 'x')}:{local}"
    return tag


def wattr(el, name):
    return el.get(f"{{{W}}}{name}")


# --- What the editor can write -------------------------------------------

# Content (children of body / p / r / tc ...).
CONTENT_OK = {
    "body", "p", "r", "t", "tab", "br", "cr", "tbl", "tr", "tc",
    "pPr", "rPr", "tblPr", "tblGrid", "gridCol", "tcPr",
    "commentRangeStart", "commentRangeEnd", "commentReference",
    "drawing:inline-picture",
}
# Things that carry no rendering and that an editor regenerates on its own.
IGNORABLE = {
    "proofErr", "lastRenderedPageBreak", "bookmarkStart", "bookmarkEnd",
    "permStart", "permEnd", "rPrChange", "pPrChange", "tblPrChange",
    "tcPrChange", "trPrChange", "sectPrChange", "tblGridChange",
    "w14:paraId", "x:paraId",
}
PPR_OK = {
    "pStyle", "jc", "spacing", "ind", "keepNext", "keepLines",
    "widowControl", "snapToGrid", "wordWrap", "adjustRightInd",
    "autoSpaceDE", "autoSpaceDN", "pageBreakBefore",
}
RPR_OK = {
    "b", "bCs", "i", "iCs", "u", "strike", "rFonts", "sz", "szCs", "color",
    "highlight", "spacing", "kern", "lang", "rStyle", "noProof", "vertAlign",
}
TBLPR_OK = {"tblW", "tblBorders", "tblLook"}
TCPR_OK = {"tcW"}
JC_OK = {"left", "center", "right", "both", "distribute", "start", "end"}

# The blank document's page: A4, 1 inch margins.
BLANK_PGSZ = {"w": "11906", "h": "16838"}
BLANK_PGMAR = {"top": "1440", "right": "1440", "bottom": "1440", "left": "1440"}
BLANK_STYLES = {"Normal", "Heading1", "Heading2", "Heading3", "Heading4",
                "Heading5", "Heading6"}

PROP_CONTAINERS = {"pPr", "rPr", "tblPr", "trPr", "tcPr", "sectPr",
                   "tblPrEx", "tblGrid"}


class Census:
    def __init__(self):
        self.features = Counter()   # feature -> occurrences

    def add(self, key, n=1):
        self.features[key] += n


def classify_drawing(el):
    """One key per drawing: is it the inline picture the editor writes?"""
    anchor = el.find(f"{{{WP}}}anchor")
    if anchor is not None:
        return "drawing:anchor"
    inline = el.find(f"{{{WP}}}inline")
    if inline is None:
        return "drawing:other"
    gd = inline.find(f".//{{{A}}}graphicData")
    uri = gd.get("uri", "") if gd is not None else ""
    if uri.endswith("/picture"):
        return "drawing:inline-picture"
    kind = uri.rstrip("/").rsplit("/", 1)[-1] or "unknown"
    return f"drawing:inline-{kind}"


def walk_props(c, container, el):
    """Children of a property container, with value-level splits where the
    editor writes only part of an element's range."""
    for ch in el:
        name = qn(ch.tag)
        if name is None or name in IGNORABLE:
            continue
        key = f"{container}/{name}"
        if container == "pPr" and name == "spacing":
            rule = wattr(ch, "lineRule")
            if wattr(ch, "line") is not None and rule not in (None, "auto"):
                key += f"@lineRule={rule}"
            elif any(wattr(ch, a) is not None for a in
                     ("beforeLines", "afterLines")):
                key += "@lines"
            elif any(wattr(ch, a) in ("1", "true", "on") for a in
                     ("beforeAutospacing", "afterAutospacing")):
                key += "@autospacing"
        elif container == "pPr" and name == "ind":
            if any(wattr(ch, a) is not None for a in
                   ("leftChars", "rightChars", "firstLineChars",
                    "hangingChars", "startChars", "endChars")):
                key += "@chars"
        elif container == "pPr" and name == "jc":
            if wattr(ch, "val") not in JC_OK:
                key += f"@{wattr(ch, 'val')}"
        elif container == "rPr" and name == "u":
            if wattr(ch, "val") in (None, "none", "single"):
                pass
            elif wattr(ch, "color") is not None and False:
                pass
        if container == "pPr" and name in ("rPr", "sectPr"):
            # Paragraph-mark formatting and a section break inside a paragraph.
            c.add("pPr/rPr" if name == "rPr" else "pPr/sectPr(section break)")
            continue
        c.add(key)


def walk(c, el, final_sectpr):
    name = qn(el.tag)
    if name is None or name in IGNORABLE:
        return
    if name == "drawing":
        c.add(classify_drawing(el))
        return
    if name == "mc:AlternateContent":
        choice = el.find(f"{{{MC}}}Choice")
        if choice is not None:
            for ch in choice:
                walk(c, ch, final_sectpr)
        return
    if name in ("pict", "object"):
        c.add(f"vml:{name}")
        return
    if name.startswith("m:"):
        c.add("math")
        return
    if name in PROP_CONTAINERS:
        if name == "sectPr":
            if el is final_sectpr:
                return  # handled per mode
            c.add("sectPr(section break)")
            return
        if name == "tblGrid":
            return
        walk_props(c, name, el)
        return
    if name == "sdt":
        # A content control is one feature; its body is ordinary content.
        c.add("sdt")
        content = el.find(f"{{{W}}}sdtContent")
        for ch in (content if content is not None else []):
            walk(c, ch, final_sectpr)
        return
    if name in ("del", "moveFrom"):
        c.add("tracked-change")
        return
    if name in ("ins", "moveTo"):
        c.add("tracked-change")
        for ch in el:
            walk(c, ch, final_sectpr)
        return
    if name == "trHeight":
        # Some generated files put trHeight straight under w:tr.
        c.add("trPr/trHeight")
        return
    if name == "br":
        t = wattr(el, "type")
        c.add("br" if t in (None, "textWrapping", "page", "column")
              else f"br@{t}")
        return
    if name == "fldChar" or name == "instrText" or name == "fldSimple":
        c.add("field")
        if name == "fldSimple":
            for ch in el:
                walk(c, ch, final_sectpr)
        return
    c.add(name)
    for ch in el:
        walk(c, ch, final_sectpr)


def feature_ok_template(key):
    if key in CONTENT_OK:
        return True
    if "/" in key:
        container, rest = key.split("/", 1)
        table = {"pPr": PPR_OK, "rPr": RPR_OK, "tblPr": TBLPR_OK,
                 "tcPr": TCPR_OK}.get(container)
        return table is not None and rest in table
    return False


def read(z, name):
    try:
        return etree.fromstring(z.read(name))
    except KeyError:
        return None


def has_text(root):
    return root is not None and any(
        (t.text or "").strip() for t in root.iter(f"{{{W}}}t")) or (
        root is not None and root.find(f".//{{{W}}}drawing") is not None)


def blank_only_features(z, doc_root, final_sectpr):
    """Features that only matter when starting from the blank document."""
    out = Counter()
    if final_sectpr is not None:
        for ch in final_sectpr:
            name = qn(ch.tag)
            if name is None or name in IGNORABLE:
                continue
            if name == "pgSz":
                attrs = {k: wattr(ch, k) for k in ("w", "h")}
                if attrs != BLANK_PGSZ or wattr(ch, "orient") == "landscape":
                    out["page:size"] += 1
            elif name == "pgMar":
                attrs = {k: wattr(ch, k) for k in BLANK_PGMAR}
                if attrs != BLANK_PGMAR:
                    out["page:margins"] += 1
            elif name in ("headerReference", "footerReference"):
                continue  # counted by content below
            elif name in ("cols",):
                if (wattr(ch, "num") or "1") != "1":
                    out["page:columns"] += 1
            elif name == "docGrid":
                if wattr(ch, "type") not in (None, "default"):
                    out["page:docGrid"] += 1
            elif name in ("pgNumType", "titlePg", "vAlign", "textDirection",
                          "pgBorders", "lnNumType", "bidi", "rtlGutter"):
                out[f"page:{name}"] += 1
    names = z.namelist()
    if any(has_text(read(z, n)) for n in names
           if re.match(r"word/(header|footer)\d*\.xml$", n)):
        out["headers/footers"] += 1
    fn = read(z, "word/footnotes.xml")
    if fn is not None and len([x for x in fn if wattr(x, "type") is None]):
        out["footnotes"] += 1
    en = read(z, "word/endnotes.xml")
    if en is not None and len([x for x in en if wattr(x, "type") is None]):
        out["endnotes"] += 1
    # Styles the document actually names.
    used = set()
    for tag in ("pStyle", "rStyle", "tblStyle"):
        for el in doc_root.iter(f"{{{W}}}{tag}"):
            used.add(wattr(el, "val"))
    if used - BLANK_STYLES:
        out["styles:custom"] += 1
    st = read(z, "word/styles.xml")
    if st is not None:
        dd = st.find(f"{{{W}}}docDefaults")
        if dd is not None:
            rf = dd.find(f".//{{{W}}}rFonts")
            sz = dd.find(f".//{{{W}}}sz")
            fonts = None if rf is None else (wattr(rf, "ascii"),
                                             wattr(rf, "eastAsia"),
                                             wattr(rf, "asciiTheme"))
            if fonts not in (None, ("Calibri", None, None)) or (
                    sz is not None and wattr(sz, "val") != "22"):
                out["styles:docDefaults"] += 1
        normal = st.find(f"{{{W}}}style[@{{{W}}}styleId='Normal']")
        if normal is not None and (normal.find(f"{{{W}}}pPr") is not None or
                                   normal.find(f"{{{W}}}rPr") is not None):
            out["styles:Normal"] += 1
    # Numbering reached through a style still needs numbering definitions.
    if read(z, "word/numbering.xml") is not None and (
            doc_root.find(f".//{{{W}}}numPr") is not None):
        out["numbering:definitions"] += 1
    return out


def census_doc(path):
    with zipfile.ZipFile(path) as z:
        doc = read(z, "word/document.xml")
        if doc is None:
            raise ValueError("no word/document.xml")
        body = doc.find(f"{{{W}}}body")
        final_sectpr = body.find(f"{{{W}}}sectPr") if body is not None else None
        c = Census()
        for ch in (body if body is not None else []):
            walk(c, ch, final_sectpr)
        template_blockers = sorted(k for k in c.features
                                   if not feature_ok_template(k))
        blank_extra = blank_only_features(z, doc, final_sectpr)
    return {
        "features": dict(c.features),
        "template_blockers": template_blockers,
        "blank_blockers": sorted(set(template_blockers) | set(blank_extra)),
    }


SYNTHETIC = re.compile(r"^(gen2?|repro|test|pixel|sweep|indent|prog|d7|d11)_|^probe?|^_pixel_tmp")


def is_real(name):
    """Self-authored probes and generated sweeps isolate one feature each;
    the real documents are the ones a user would ask the editor to make."""
    return not SYNTHETIC.search(Path(name).name)


def greedy_cover(per_doc, mode):
    """Order the missing features so each step frees the most documents."""
    remaining = {d: set(r[f"{mode}_blockers"]) for d, r in per_doc.items()}
    added, steps = set(), []
    done = sum(1 for s in remaining.values() if not s)
    steps.append({"add": None, "rebuildable": done})
    while True:
        gain = Counter()
        # A feature frees a doc when it is the doc's ONLY remaining blocker;
        # ties are broken by how many docs still need it at all.
        need = Counter()
        for s in remaining.values():
            left = s - added
            for f in left:
                need[f] += 1
            if len(left) == 1:
                gain[next(iter(left))] += 1
        if not need:
            break
        best = max(need, key=lambda f: (gain[f], need[f], f))
        added.add(best)
        done = sum(1 for s in remaining.values() if not (s - added))
        steps.append({"add": best, "rebuildable": done,
                      "docs_needing": need[best]})
        if len(steps) > 40:
            break
    return steps


def main():
    corpus = Path(sys.argv[1] if len(sys.argv) > 1
                  else "tools/golden-test/documents/docx")
    out = Path(sys.argv[2] if len(sys.argv) > 2 else
               "pipeline_data/edit_coverage")
    out.mkdir(parents=True, exist_ok=True)
    per_doc, errors = {}, {}
    for p in sorted(corpus.rglob("*.docx")):
        if p.name.startswith("~$"):
            continue
        try:
            per_doc[str(p.relative_to(corpus))] = census_doc(p)
        except Exception as e:  # noqa: BLE001 - census must cover every file
            errors[str(p.relative_to(corpus))] = repr(e)
    n = len(per_doc)
    summary = {"corpus": str(corpus), "docs": n, "errors": errors}
    groups = {"all": per_doc,
              "real": {d: r for d, r in per_doc.items() if is_real(d)}}
    for group, docs in groups.items():
        summary[group] = {"docs": len(docs)}
        for mode in ("template", "blank"):
            blocked = Counter()
            for r in docs.values():
                for f in r[f"{mode}_blockers"]:
                    blocked[f] += 1
            summary[group][mode] = {
                "rebuildable": sorted(d for d, r in docs.items()
                                      if not r[f"{mode}_blockers"]),
                "blockers_per_doc": dict(sorted(Counter(
                    len(r[f"{mode}_blockers"]) for r in docs.values()).items())),
                "feature_docs_blocked": dict(blocked.most_common()),
                "greedy": greedy_cover(docs, mode),
            }
    (out / "per_doc.json").write_text(
        json.dumps(per_doc, ensure_ascii=False, indent=1), encoding="utf-8")
    (out / "_summary.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=1), encoding="utf-8")
    for group, mode in (("all", "template"), ("real", "template"),
                        ("real", "blank")):
        s = summary[group][mode]
        print(f"[{group}/{mode}] rebuildable "
              f"{len(s['rebuildable'])}/{summary[group]['docs']}")
        for f, k in list(s["feature_docs_blocked"].items())[:30]:
            print(f"   {k:4d}  {f}")
        print("  greedy:", " -> ".join(
            f"{st['add']}:{st['rebuildable']}" for st in s["greedy"][:25]))
    if errors:
        print("errors:", len(errors))


if __name__ == "__main__":
    main()
