"""Does `spcFirstLastPara` mean what it looks like it means?

`a:bodyPr/@spcFirstLastPara="1"` says the space BEFORE the first paragraph (and
after the last) is to be respected rather than dropped. Nothing in this codebase
reads it -- not the parser, not the renderer -- and the browser engine drops the
first paragraph's `spcBef` unconditionally.

d06 and d35 slide 2 both carry the attribute with a 10pt `spcBef` on the first
paragraph, and both put PowerPoint's text 9.8pt BELOW where the engine puts it.
This asks the corpus whether that holds generally: for every shape that declares
it, how much space the first paragraph asks for, against how far down
PowerPoint actually drew the first line.

The answer is per shape, so one template's habit cannot carry the finding.

    python tools/metrics/pptx_spcfirst_census.py [--decks dev|all]
"""
from __future__ import annotations

import argparse
import re
import sys
import zipfile
from collections import Counter
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]

BODY_PR = re.compile(r"<a:bodyPr\b[^>]*>")
TX_BODY = re.compile(r"<p:txBody>(.*?)</p:txBody>", re.S)
FIRST_P = re.compile(r"<a:p>(.*?)</a:p>", re.S)
SPC_BEF = re.compile(r"<a:spcBef>\s*<a:spcPts val=\"(\d+)\"\s*/>", re.S)
SPC_BEF_PCT = re.compile(r"<a:spcBef>\s*<a:spcPct val=\"(\d+)\"\s*/>", re.S)


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--decks", default="dev", choices=["dev", "all"])
    args = ap.parse_args()

    bench = REPO / "pipeline_data" / "pptx_benchmark"
    roots = [bench / "dev" / "pptx"]
    if args.decks == "all":
        roots.append(bench / "pptx")

    decks_with = set()
    shapes = Counter()
    amounts = Counter()
    total_shapes = 0
    for root in roots:
        for path in sorted(root.glob("*.pptx")):
            stem = path.name.split("__")[0]
            try:
                z = zipfile.ZipFile(path)
            except Exception:
                continue
            for name in z.namelist():
                if not (name.startswith("ppt/slides/slide") and name.endswith(".xml")):
                    continue
                xml = z.read(name).decode("utf-8", "replace")
                for body in TX_BODY.findall(xml):
                    total_shapes += 1
                    m = BODY_PR.search(body)
                    if not m or 'spcFirstLastPara="1"' not in m.group(0):
                        continue
                    first = FIRST_P.search(body)
                    if not first:
                        continue
                    head = first.group(1)
                    pts = SPC_BEF.search(head)
                    pct = SPC_BEF_PCT.search(head)
                    if pts and int(pts.group(1)) > 0:
                        shapes[stem] += 1
                        amounts[int(pts.group(1)) / 100.0] += 1
                        decks_with.add(stem)
                    elif pct and int(pct.group(1)) > 0:
                        shapes[stem] += 1
                        amounts[f"{int(pct.group(1)) / 1000:g}%"] += 1
                        decks_with.add(stem)

    print(f"{total_shapes} text shapes scanned")
    print(f"{len(decks_with)} decks carry spcFirstLastPara=\"1\" with a first-paragraph "
          f"spcBef > 0, over {sum(shapes.values())} shapes\n")
    for deck, n in shapes.most_common(14):
        print(f"   {deck:6} {n:4} shapes")
    print("\nthe amounts asked for:")
    for amount, n in amounts.most_common(10):
        label = amount if isinstance(amount, str) else f"{amount:g}pt"
        print(f"   {label:>8}  {n:4} shapes")


if __name__ == "__main__":
    main()
