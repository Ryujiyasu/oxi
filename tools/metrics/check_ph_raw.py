import json
import sys
from pathlib import Path
_REPO = Path(__file__).resolve().parents[2]

sys.stdout.reconfigure(encoding="utf-8")

with open(str(_REPO / r"pipeline_data\pptx_probes\ph_fixed.json"), encoding="utf-8") as f:
    d = json.load(f)

# dump top-level keys
print("TOP KEYS:", list(d.keys()))
for s in d.get("slides", []):
    print("--- slide keys:", list(s.keys()))
    for sh in s.get("shapes", []):
        print("  shape keys:", list(sh.keys()))
        print("  ", json.dumps(sh, ensure_ascii=False)[:300])
