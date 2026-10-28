# oxidocs-docling

Page numbers and bounding boxes for `.docx` in [Docling](https://github.com/docling-project/docling).

Docling's DOCX backend leaves every item's `prov` empty because a Word file
stores no layout. Oxi lays the document out, so this crate can say where each
item lands: the page and the box, the way Docling already reports them for PDF.

## Annotate mode (recommended)

Run Docling as usual, then hand its JSON and the same `.docx` to `oxi-docling`.
It writes `prov` for every text, table and picture and fills `pages`. Nothing
else changes: structure, labels and text stay exactly as Docling produced them.

```bash
cargo build --release -p oxidocs-docling
docling report.docx --to json                      # or docling-rs, or the Python API
./target/release/oxi-docling annotate report.docx report.json report.prov.json
```

How items are placed:
- Text items are found in Oxi's painted text in reading order (whitespace
  ignored), one `prov` entry per page; `charspan` is the range of `orig` on
  that page. A list item's box includes its number or bullet.
- `formula` items and pictures are matched in reading order; a table through
  its cells' text (each cell also gets a `bbox`); headers and footers on every
  page they appear.
- An item without text of its own takes the union of its children; a comment
  takes the place of the text it annotates.
- An empty paragraph takes the empty line Oxi lays out for it. If Oxi paints
  nothing for it (a section-break mark), it gets a zero-height box at the end
  of the text before it.

## Standalone mode

`oxi-docling report.docx report.json` writes Oxi's own DoclingDocument, for
use where Docling is not available (the browser, Wasm). Its structure is
close to Docling's but not identical: Docling splits formatted runs into
inline groups, labels checkboxes and code blocks, and keeps rich table
cells as items, which this mode does not.

## Conventions

Points, top-left origin (`coord_origin: TOPLEFT`), `pages[n].size` in the same
unit, `charspan` in Unicode code points over the item's `orig`.

## Measured

On Docling's own public `tests/data/docx` with its ground-truth JSON
(36 files, 845 items, 2026-10-28):

| check | result |
|---|---|
| validates against `DoclingDocument_1_10.json` (refs, pages, boxes, charspans) | 36 / 36 |
| items with `prov` | 845 / 845 |
| files changed beyond `prov`, `pages`, cell `bbox` | 0 |

Reproduce:

```bash
crates/oxidocs-docling/scripts/export_all.sh \
    <docling>/tests/data/docx/sources <docling>/tests/data/docx/groundtruth out \
    <docling-core>/docs/schemas/DoclingDocument_1_10.json
```

How close the pages and boxes are to Word depends on Oxi's layout, which is
measured against Word on frozen blind sets (see the repository README). A
document using a font Oxi has no metrics for can paginate differently.
