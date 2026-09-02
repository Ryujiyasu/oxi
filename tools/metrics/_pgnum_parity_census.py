"""Census: docs whose blank-page behaviour depends on LOGICAL page parity.

Three populations, because S1291 changes the rule for each:
  (a) evenPage/oddPage sections            -- S732's rule becomes logical
  (b) evenAndOddHeaders + pgNumType start  -- S957's rule becomes alternation
  (c) of (b), the ones whose FIRST section restarts at an EVEN number, which
      is where the old and the new rule actually disagree
"""
import re, sys, zipfile
from pathlib import Path
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

SECT = re.compile(r'<w:sectPr[^>]*>.*?</w:sectPr>', re.S)
TYPE = re.compile(r'<w:type w:val="(\w+)"')
START = re.compile(r'<w:pgNumType[^>]*w:start="(\d+)"')

roots = [Path("pipeline_data/docx_corpus"), Path("tools/golden-test/documents/docx")]
a = b = c = n = 0
rows = []
for root in roots:
    if not root.exists():
        continue
    for p in sorted(root.rglob("*.docx")):
        if p.name.startswith("~$"):
            continue
        n += 1
        try:
            z = zipfile.ZipFile(p)
            xml = z.read("word/document.xml").decode("utf-8", "replace")
            try:
                st = z.read("word/settings.xml").decode("utf-8", "replace")
            except KeyError:
                st = ""
        except Exception:
            continue
        eoh = "evenAndOddHeaders" in st
        sects = []
        for m in SECT.finditer(xml):
            t = m.group(0)
            ty = TYPE.search(t)
            sr = START.search(t)
            sects.append((ty.group(1) if ty else None,
                          int(sr.group(1)) if sr else None))
        if not sects:
            continue
        has_parity = any(t in ("oddPage", "evenPage") for t, _ in sects)
        has_restart = eoh and any(s is not None for _, s in sects)
        first_start = sects[0][1]
        even_first = first_start is not None and first_start % 2 == 0
        if has_parity:
            a += 1
        if has_restart:
            b += 1
        if has_restart and even_first:
            c += 1
        if has_parity or has_restart:
            rows.append((p.name, eoh, first_start,
                         [t for t, _ in sects if t in ("oddPage", "evenPage")],
                         [s for _, s in sects if s is not None][:4]))

print(f"{n} docx scanned")
print(f"  (a) evenPage/oddPage sections          : {a}")
print(f"  (b) evenAndOddHeaders + pgNumType start: {b}")
print(f"  (c) of (b), FIRST section start is EVEN: {c}   <- old and new rule disagree")
print()
for r in rows:
    print(f"  {r[0][:52]:<52} eoh={str(r[1]):<5} first_start={str(r[2]):<5} "
          f"parity={r[3]} starts={r[4]}")
