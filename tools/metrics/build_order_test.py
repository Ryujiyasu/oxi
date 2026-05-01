"""Critical test: swap the order of <w:tblpPr> and <w:tblStyle> in TP3
without changing anything else. If slope flips from 0 to 1, the cause is
ELEMENT ORDER (tblpPr must come AFTER tblStyle per ECMA-376 CT_TblPrBase).

Output: tools/metrics/order_test_repro/O_*.docx
"""
import re, zipfile
from pathlib import Path

SRC = Path(r"C:\Users\ryuji\oxi-1\tools\metrics\tblppr_anchor_repro\TP3_anchor1_tblpY600.docx")
OUT_DIR = Path(r"C:\Users\ryuji\oxi-1\tools\metrics\order_test_repro")
OUT_DIR.mkdir(parents=True, exist_ok=True)


def load_pkg(p):
    out = {}
    with zipfile.ZipFile(p) as z:
        for info in z.infolist():
            out[info.filename] = z.read(info.filename)
    return out


def write_pkg(p, pkg):
    with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as z:
        for f, d in pkg.items():
            z.writestr(f, d)


def set_tblpY(xml: str, tw: int) -> str:
    return re.sub(r'(w:tblpY=")(-?\d+)(")', f'\\g<1>{tw}\\g<3>', xml, count=1)


def variant_baseline(xml: str) -> str:
    """Verbatim TP3 — tblpPr before tblStyle, expect slope=0."""
    return xml


def variant_swapped(xml: str) -> str:
    """Move tblStyle to BEFORE tblpPr per ECMA-376 child order."""
    # Find the <w:tblPr>...</w:tblPr> wrapping the floating table.
    # In TP3 it contains: <w:tblpPr .../><w:tblStyle .../><w:tblW .../>...
    pattern = re.compile(
        r'(<w:tblPr>)(<w:tblpPr\b[^/]*?/>)(<w:tblStyle\b[^/]*?/>)',
        re.S,
    )
    new_xml, n = pattern.subn(r'\1\3\2', xml, count=1)
    assert n == 1, f"swap pattern matched {n} times"
    return new_xml


def variant_no_tblstyle(xml: str) -> str:
    """Remove tblStyle reference (already known to give slope=1, control)."""
    return re.sub(r'<w:tblStyle\b[^/]*?/>', '', xml, count=1)


VARIANTS = [
    ("baseline", variant_baseline),
    ("swapped",  variant_swapped),
    ("noStyle",  variant_no_tblstyle),
]
TBLPYS = [("Y50", 50), ("Y600", 600), ("Y2000", 2000)]


def main():
    base_pkg = load_pkg(SRC)
    base_xml = base_pkg["word/document.xml"].decode("utf-8")
    for vname, vfn in VARIANTS:
        for yname, ytw in TBLPYS:
            txt = vfn(base_xml)
            txt = set_tblpY(txt, ytw)
            pkg = {k: v for k, v in base_pkg.items()}
            pkg["word/document.xml"] = txt.encode("utf-8")
            out = OUT_DIR / f"O_{vname}_{yname}.docx"
            write_pkg(out, pkg)
            print(f"Wrote {out.name}")
    print(f"\nWrote {len(VARIANTS)*len(TBLPYS)} variants")


if __name__ == "__main__":
    main()
