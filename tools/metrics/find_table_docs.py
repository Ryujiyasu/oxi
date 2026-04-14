"""Find docs with tables in DML cache, sorted by row count."""
import json, glob, os, sys
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

results = []
for f in sorted(glob.glob('pipeline_data/word_dml/*.json')):
    try:
        d = json.load(open(f, encoding='utf-8'))
    except Exception:
        continue
    name = os.path.basename(f).replace('.json', '')
    tables = d.get('tables', [])
    if not tables:
        continue
    # Per-table row counts
    row_counts = [t.get('rows', 0) for t in tables]
    results.append((name, len(tables), row_counts))

# Find docs with at least one 1-row table
print(f"Total docs with tables: {len(results)}")
print("\n=== Docs with a 1-row table ===")
n_1row = 0
for n, ntbl, rcs in results:
    if 1 in rcs:
        n_1row += 1
        print(f"  tables={ntbl} rows={rcs} : {n}")
        if n_1row >= 20: break
print(f"\nTotal 1-row containing: {sum(1 for _,_,rcs in results if 1 in rcs)}")
print("\n=== Docs with multi-row tables (no 1-row) ===")
multi = [(n,ntbl,rcs) for n,ntbl,rcs in results if 1 not in rcs and max(rcs) >= 3]
for n, ntbl, rcs in multi[:10]:
    print(f"  tables={ntbl} rows={rcs[:5]}... : {n}")
