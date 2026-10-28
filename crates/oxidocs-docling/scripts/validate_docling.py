"""Validate DoclingDocument JSON files: JSON Schema plus the integrity docling-core relies on
(self_ref numbering, parent/children agreement, prov pages and boxes inside the page, charspan inside orig).

    python validate_docling.py <DoclingDocument_1_10.json> out/*.json

The schema file is docling-core/docs/schemas/DoclingDocument_1_10.json. Needs the `jsonschema` package.
"""
import json
import sys

from jsonschema import Draft202012Validator

SCHEMA = json.load(open(sys.argv[1]))
V = Draft202012Validator(SCHEMA)


def resolve(d, ref):
    if ref in ("#/body", "#/furniture"):
        return d[ref[2:]]
    _, kind, i = ref.split("/")
    return d[kind][int(i)]


def check(d):
    errs = [f"schema: {e.json_path}: {e.message[:120]}" for e in V.iter_errors(d)]
    pages = d["pages"]
    for kind in ("texts", "tables", "pictures", "groups"):
        for i, it in enumerate(d[kind]):
            ref = f"#/{kind}/{i}"
            if it["self_ref"] != ref:
                errs.append(f"{ref}: self_ref {it['self_ref']}")
            par = resolve(d, it["parent"]["$ref"])
            if {"$ref": ref} not in par["children"]:
                errs.append(f"{ref}: not among its parent's children")
            for c in it["children"]:
                if resolve(d, c["$ref"])["parent"]["$ref"] != ref:
                    errs.append(f"{ref}: child {c['$ref']} points elsewhere")
            n = len(it.get("orig", ""))
            for p in it.get("prov", []):
                pg = pages.get(str(p["page_no"]))
                if pg is None:
                    errs.append(f"{ref}: page {p['page_no']} not in pages")
                    continue
                b, W, H = p["bbox"], pg["size"]["width"], pg["size"]["height"]
                if not (-0.5 <= b["l"] <= b["r"] <= W + 0.5 and -0.5 <= b["t"] <= b["b"] <= H + 0.5):
                    errs.append(f"{ref}: bbox {b} outside page {W}x{H}")
                a, z = p["charspan"]
                if not (0 <= a <= z <= n):
                    errs.append(f"{ref}: charspan {p['charspan']} vs len {n}")
    for root in ("body", "furniture"):
        for c in d[root]["children"]:
            if resolve(d, c["$ref"])["parent"]["$ref"] != f"#/{root}":
                errs.append(f"{root}: child {c['$ref']} points elsewhere")
    return errs


bad = 0
for f in sys.argv[2:]:
    e = check(json.load(open(f)))
    if e:
        bad += 1
        print(f"FAIL {f}: {len(e)} problems")
        for x in e[:6]:
            print("   ", x)
print(f"{len(sys.argv) - 2} files, {bad} with problems")
