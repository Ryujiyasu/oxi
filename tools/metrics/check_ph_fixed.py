import json
import sys
from pathlib import Path
_REPO = Path(__file__).resolve().parents[2]

sys.stdout.reconfigure(encoding="utf-8")

with open(str(_REPO / r"pipeline_data\pptx_probes\ph_fixed.json"), encoding="utf-8") as f:
    d = json.load(f)

for s in d.get("slides", []):
    print(f"--- slide {s.get('slide_index')} ---")
    for sh in s.get("shapes", []):
        print(
            f"  {sh.get('type')} x={sh.get('x')} y={sh.get('y')} "
            f"w={sh.get('w')} h={sh.get('h')} rot={sh.get('rotation')} "
            f"st={sh.get('shape_type')} text={str(sh.get('text',''))[:24]!r}"
        )
