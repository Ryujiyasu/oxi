# -*- coding: utf-8 -*-
"""Write the measured font row-height table for the browser page to read.

The same measurements the renderer compiles into `row_defaults.rs`, as JSON
keyed "face|quarter-points" so `web/row-geometry.html` can apply the same
rule the renderer does instead of guessing at font metrics.
"""
import io
import json
import sys

SWEEP = r"pipeline_data\com_measurements\xlsx_row_height_sweep.json"
OUT = r"web\row-heights.json"


def main():
    table = {}
    for row in json.load(io.open(SWEEP, encoding="utf-8")):
        key = "%s|%d" % (row["face"], round(float(row["size"]) * 4))
        px = int(round(row["standard_height_pt"] / 0.75))
        table.setdefault(key, px)
    with io.open(OUT, "w", encoding="utf-8", newline="\n") as f:
        json.dump(table, f, ensure_ascii=False, sort_keys=True, indent=0)
    print("wrote %d font heights to %s" % (len(table), OUT))


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
