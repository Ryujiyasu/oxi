import json
import sys

sys.stdout.reconfigure(encoding="utf-8")

with open(r"C:\Users\ryuji\oxi-main\pipeline_data\pptx_probes\ph_fixed.json", encoding="utf-8") as f:
    d = json.load(f)

for s in d.get("slides", []):
    print(f"--- slide {s.get('slide_index')} ---")
    for sh in s.get("shapes", []):
        print(
            f"  {sh.get('type')} x={sh.get('x')} y={sh.get('y')} "
            f"w={sh.get('w')} h={sh.get('h')} rot={sh.get('rotation')} "
            f"st={sh.get('shape_type')} text={str(sh.get('text',''))[:24]!r}"
        )
