# -*- coding: utf-8 -*-
"""Re-render just the LibreOffice column of a frozen set.

An engine that shipped a new version after a set was measured leaves the table
comparing this build of Oxi against last quarter's build of somebody else. The
Word ground truth does not change and neither does any other engine, so only
one column has to be redone.

    python tools/metrics/rerender_lo_column.py en D
    python tools/metrics/rerender_lo_column.py ja D

Score it afterwards with `score_engine_column.py <lang> <letter> lo`.
"""
from __future__ import annotations

import sys
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]


def main() -> int:
    if len(sys.argv) < 3 or sys.argv[1] not in ("en", "ja"):
        print(f"usage: {Path(sys.argv[0]).name} [en|ja] <rotation letter>")
        return 2
    lang, letter = sys.argv[1], sys.argv[2].upper()
    bench = REPO / "pipeline_data" / f"{lang}_benchmark"
    module = ("_measure_ssim_blind" if lang == "en" else "_measure_ssim_jablind") + letter + "50"
    sys.path.insert(0, str(bench))
    sys.path.insert(0, str(REPO / "tools" / "metrics"))
    mod = __import__(module)

    if not mod.SOFFICE.is_file():
        print(f"LibreOffice is not where the module expects it: {mod.SOFFICE}")
        return 1
    docs = mod.selections()
    mod.LO_PDF.mkdir(parents=True, exist_ok=True)
    print(f"{lang} blind-{letter}: {len(docs)} documents through LibreOffice")
    mod.pool_run("LO", docs, mod.render_lo)
    made = len(list(mod.LO_PDF.glob("*.pdf")))
    print(f"{made} of {len(docs)} rendered into {mod.LO_PDF.name}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    raise SystemExit(main())
