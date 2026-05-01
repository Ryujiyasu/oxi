"""Take TP3 (slope=0 baseline) and mutate ONE thing at a time. Re-measure
to isolate which difference triggers slope=0.

Mutations:
  M_baseline   : verbatim TP3 (control, expect slope=0)
  M_tblWdxa    : change <w:tblW w:type="auto" w:w="0"/> -> <w:tblW w:type="dxa" w:w="9638"/>
  M_noTblStyle : remove <w:tblStyle w:val="TableGrid"/>
  M_noStyles   : remove word/styles.xml + word/stylesWithEffects.xml
  M_noUseFE    : remove <w:useFELayout/> from settings.xml
  M_noNumbering: remove word/numbering.xml + reference

Each mutation is paired with Y50 + Y600 (the original tblpY=600tw=30pt).
The Y50 variants use tblpY=50tw=2.5pt (from TP1 spirit).
Then we have tblpY=2000tw too for amplification.

Output: tools/metrics/tp3_mutate_repro/M_*.docx
"""
import re, zipfile, io, shutil
from pathlib import Path

SRC = Path(r"C:\Users\ryuji\oxi-1\tools\metrics\tblppr_anchor_repro\TP3_anchor1_tblpY600.docx")
OUT_DIR = Path(r"C:\Users\ryuji\oxi-1\tools\metrics\tp3_mutate_repro")
OUT_DIR.mkdir(parents=True, exist_ok=True)


def load_pkg(p):
    """Load full docx as dict {filename: bytes}."""
    out = {}
    with zipfile.ZipFile(p) as z:
        for info in z.infolist():
            out[info.filename] = z.read(info.filename)
    return out


def write_pkg(p, pkg, content_types_override=None):
    with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as z:
        for fname, data in pkg.items():
            z.writestr(fname, data)


def set_tblpY(doc_xml: bytes, new_tw: int) -> bytes:
    txt = doc_xml.decode("utf-8")
    # Replace ONLY the first w:tblpY value
    pattern = re.compile(r'(w:tblpY=")(-?\d+)(")')
    replaced = {"n": 0}
    def repl(m):
        if replaced["n"] == 0:
            replaced["n"] += 1
            return f'{m.group(1)}{new_tw}{m.group(3)}'
        return m.group(0)
    txt = pattern.sub(repl, txt, count=1)
    return txt.encode("utf-8")


# Mutation functions: take pkg dict, return modified pkg dict (in place ok).
def mut_tblWdxa(pkg):
    txt = pkg["word/document.xml"].decode("utf-8")
    txt = txt.replace('<w:tblW w:type="auto" w:w="0"/>',
                      '<w:tblW w:type="dxa" w:w="9638"/>',
                      1)
    pkg["word/document.xml"] = txt.encode("utf-8")
    return pkg


def mut_noTblStyle(pkg):
    txt = pkg["word/document.xml"].decode("utf-8")
    txt = re.sub(r'<w:tblStyle w:val="[^"]*"/>', '', txt, count=1)
    pkg["word/document.xml"] = txt.encode("utf-8")
    return pkg


def mut_noStyles(pkg):
    """Remove styles.xml + stylesWithEffects.xml from package + Content_Types
    + relationship references."""
    for f in ["word/styles.xml", "word/stylesWithEffects.xml"]:
        pkg.pop(f, None)
    # Strip relationships
    rels_xml = pkg.get("word/_rels/document.xml.rels", b"").decode("utf-8")
    rels_xml = re.sub(
        r'<Relationship[^/]*Type="[^"]*styles[^"]*"[^/]*/>', "", rels_xml
    )
    rels_xml = re.sub(
        r'<Relationship[^/]*Type="[^"]*stylesWithEffects[^"]*"[^/]*/>', "", rels_xml
    )
    pkg["word/_rels/document.xml.rels"] = rels_xml.encode("utf-8")
    # Strip Content_Types overrides
    ct = pkg.get("[Content_Types].xml", b"").decode("utf-8")
    ct = re.sub(r'<Override[^/]*PartName="/word/styles\.xml"[^/]*/>', "", ct)
    ct = re.sub(r'<Override[^/]*PartName="/word/stylesWithEffects\.xml"[^/]*/>', "", ct)
    pkg["[Content_Types].xml"] = ct.encode("utf-8")
    return pkg


def mut_noUseFE(pkg):
    txt = pkg["word/settings.xml"].decode("utf-8")
    txt = txt.replace('<w:useFELayout/>', '', 1)
    pkg["word/settings.xml"] = txt.encode("utf-8")
    return pkg


def mut_noNumbering(pkg):
    pkg.pop("word/numbering.xml", None)
    rels_xml = pkg.get("word/_rels/document.xml.rels", b"").decode("utf-8")
    rels_xml = re.sub(
        r'<Relationship[^/]*Type="[^"]*numbering[^"]*"[^/]*/>', "", rels_xml
    )
    pkg["word/_rels/document.xml.rels"] = rels_xml.encode("utf-8")
    ct = pkg.get("[Content_Types].xml", b"").decode("utf-8")
    ct = re.sub(r'<Override[^/]*PartName="/word/numbering\.xml"[^/]*/>', "", ct)
    pkg["[Content_Types].xml"] = ct.encode("utf-8")
    return pkg


MUTATIONS = {
    "baseline":   lambda p: p,
    "tblWdxa":    mut_tblWdxa,
    "noTblStyle": mut_noTblStyle,
    "noStyles":   mut_noStyles,
    "noUseFE":    mut_noUseFE,
    "noNumbering":mut_noNumbering,
}

TBLPYS = [
    ("Y50",   50),    # tiny
    ("Y600",  600),   # 30pt
    ("Y2000", 2000),  # 100pt (amplifier)
]


def main():
    base_pkg = load_pkg(SRC)
    for mname, mfn in MUTATIONS.items():
        for yname, ytw in TBLPYS:
            pkg = {k: v for k, v in base_pkg.items()}  # shallow copy
            pkg = mfn(pkg)
            pkg["word/document.xml"] = set_tblpY(pkg["word/document.xml"], ytw)
            out = OUT_DIR / f"M_{mname}_{yname}.docx"
            write_pkg(out, pkg)
            print(f"Wrote {out.name}")
    print(f"\nWrote {len(MUTATIONS)*len(TBLPYS)} variants to {OUT_DIR}")


if __name__ == "__main__":
    main()
