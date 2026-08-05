import json
import sys

sys.stdout.reconfigure(encoding="utf-8")

with open(r"C:\Users\ryuji\oxi-main\pipeline_data\pptx_probes\ph_fixed.json", encoding="utf-8") as f:
    d = json.load(f)

# dump top-level keys
print("TOP KEYS:", list(d.keys()))
for s in d.get("slides", []):
    print("--- slide keys:", list(s.keys()))
    for sh in s.get("shapes", []):
        print("  shape keys:", list(sh.keys()))
        print("  ", json.dumps(sh, ensure_ascii=False)[:300])
