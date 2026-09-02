"""Census: docx with vertical (tbRl) sections that declare w:cols num>1.

Answers two questions the fix needs:
  (a) how many docs would change at all (vertical AND multi-col anywhere)
  (b) of those, how many need PER-SECTION bands (the num>1 sits in a
      continuous section, so page.columns misses it)
"""
import re, sys, zipfile
from pathlib import Path

SECT = re.compile(r'<w:sectPr[^>]*>.*?</w:sectPr>', re.S)
COLS = re.compile(r'<w:cols([^>]*?)/?>')
NUM = re.compile(r'w:num="(\d+)"')
TYPE = re.compile(r'<w:type w:val="(\w+)"')
TD = re.compile(r'<w:textDirection w:val="(tbRl\w*)"')

roots = [Path("pipeline_data/docx_corpus"), Path("tools/golden-test/repros")]
rows = []
for root in roots:
    if not root.exists():
        continue
    for p in sorted(root.rglob("*.docx")):
        try:
            with zipfile.ZipFile(p) as z:
                xml = z.read("word/document.xml").decode("utf-8", "replace")
        except Exception as e:
            print(f"SKIP {p}: {e}", file=sys.stderr)
            continue
        sects = []
        for m in SECT.finditer(xml):
            t = m.group(0)
            c = COLS.search(t)
            num = int(NUM.search(c.group(1)).group(1)) if (c and NUM.search(c.group(1))) else 1
            typ = TYPE.search(t)
            sects.append({
                "num": num,
                "type": typ.group(1) if typ else "nextPage",
                "vert": bool(TD.search(t)),
            })
        if not sects:
            continue
        vert = [s for s in sects if s["vert"]]
        if not vert:
            continue
        multi = [s for s in vert if s["num"] > 1]
        if not multi:
            continue
        # would page.columns alone see it? only if the FIRST section of a
        # page-group carries num>1; continuous sections merge into the first.
        first_sees = sects[0]["num"] > 1
        needs_runs = any(s["num"] > 1 and s["type"] == "continuous" for s in vert)
        rows.append((str(p), len(sects), len(vert),
                     sorted({s["num"] for s in multi}), first_sees, needs_runs))

print(f"{'doc':<62} {'sect':>4} {'vert':>4} {'nums':>10} {'first?':>6} {'runs?':>5}")
for r in rows:
    print(f"{Path(r[0]).name:<62} {r[1]:>4} {r[2]:>4} {str(r[3]):>10} {str(r[4]):>6} {str(r[5]):>5}")
print(f"\n{len(rows)} doc(s) vertical AND multi-column")
print(f"{sum(1 for r in rows if r[5])} need per-section runs (num>1 in a continuous section)")
print(f"{sum(1 for r in rows if r[4])} would already be served by page.columns")
