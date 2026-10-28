# oxidocs-docling

Export a `.docx` as a [DoclingDocument](https://github.com/docling-project/docling-core)
JSON (schema 1.10.0) in which every text item, table and picture carries
`prov`: the page it is on and its bounding box, computed by Oxi's layout engine.

Docling's own docx backend leaves `prov` empty because the format stores no
layout. This crate fills that gap for docx the way Docling already fills it
for PDF, so chunks built from a Word file can cite a page and a box.

```bash
cargo build --release -p oxidocs-docling
./target/release/oxi-docling report.docx report.json
```

Conventions: points, top-left origin (`coord_origin: TOPLEFT`), `pages[n].size`
in the same unit, `charspan` in code points over the item's `orig`.

Structure follows Docling's docx output: headings with levels, list groups,
`page_header` / `page_footer` furniture, `inline` groups for paragraphs with
equations, `formula` items with LaTeX text, `textbox` groups. Footnotes and
text inside shapes are not exported yet; the tool prints how many painted
fragments found no item.

Measured on Docling's public `tests/data/docx` (36 files, 2026-10-28): every
file validates against the 1.10 JSON Schema, 471 of 485 items carry
provenance, and 97% of the body texts in Docling's own ground truth appear in
the export.
