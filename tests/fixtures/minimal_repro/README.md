# Minimal Repro Fixtures

Per-spec minimal `.docx` files used to **confirm** Word layout specifications
under the tightened Ra loop (see `CLAUDE.md` → Ra Autonomous Loop Procedure
and `memory/feedback_ra_loop_tightening.md`).

## Purpose

A specification is **hypothesis** until **3 distinct real documents** AND **a
self-authored minimal repro** all agree on the same observed value. This
directory holds the minimal repros.

A minimal repro is a `.docx` that:

- Isolates **exactly one** spec under test
- Strips all unrelated formatting (no headers, no styles beyond what is
  necessary, no images, no themes that aren't the default)
- Varies the input on the dimension being tested across **at least 3 variants**
  (e.g. 3 font sizes, 3 cell padding values, 3 PANOSE classes)
- Is small enough that the entire `.docx` package (after unzip) is human-readable

## Layout

```
minimal_repro/
├── README.md              (this file)
├── manifest.json          (one entry per spec)
└── <spec_id>/
    ├── README.md          (what this spec is, what we expect, how to measure)
    ├── variant_a.docx
    ├── variant_b.docx
    ├── variant_c.docx
    └── measure.py         (Word COM script, writes results next to itself)
```

`<spec_id>` should match the slug used in `pipeline_data/ra_manual_measurements.json`.

## Manifest schema

`manifest.json` is the index. Entries:

```json
{
  "spec_id": "charwidth_proportional_cjk",
  "status": "hypothesis",            // hypothesis | confirmed | refuted
  "owner_memo": "memory/charwidth_spec.md",
  "variants": ["variant_a.docx", "variant_b.docx", "variant_c.docx"],
  "real_docs_agreeing": [
    "tests/fixtures/test_japanese.docx",
    "tools/golden-test/documents/docx/<doc1>.docx",
    "tools/golden-test/documents/docx/<doc2>.docx"
  ],
  "measure_script": "measure.py",
  "last_measured": "2026-04-09",
  "notes": ""
}
```

`status` transitions:

- `hypothesis` — measured on 1–2 sources, or repro not yet authored
- `confirmed` — repro + 3 real docs all agree, implementation merged with
  zero regressions on `pipeline.verify`
- `refuted` — measurement disagreed; spec needs to be re-derived from a
  richer input space (do **not** stack EXCEPTIONs on a confirmed spec —
  re-derive from scratch and downgrade to hypothesis instead)

## How to use

1. Pick a spec (or carve a new `<spec_id>` directory)
2. Author the minimal `.docx` variants by hand or via `python-docx`
3. Write `measure.py` to read each variant via Word COM and dump
   the observation (Y coordinates, advance widths, whatever the spec
   needs) to `measurements.json` in the same directory
4. Cross-check the variants against ≥3 real docs from `tools/golden-test/`
5. Update `manifest.json` and the spec's memory file
6. Implement → `pipeline.verify` → must hit zero-regression gate

## What this directory is NOT

- Not a regression suite. SSIM regression lives in `pipeline_data/ssim_baseline.json`
- Not a place to dump arbitrary test docs. Each subdirectory must have a
  one-spec focus and a `README.md` explaining what it isolates
- Not generated content. These are hand-curated, intentionally minimal,
  and meant to outlive any single Ra session
