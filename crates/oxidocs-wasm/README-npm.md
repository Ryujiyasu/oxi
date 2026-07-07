# oxidocs

A Word-compatible `.docx` layout and rendering engine, compiled to
WebAssembly — parse, lay out, edit and PDF-export Microsoft Word documents
entirely in the browser (or any JS runtime).

**Early development.** The API is 0.x and unstable; development happens in
the open at <https://gitlab.com/Ryujiyasu/oxi>. The engine core is also
published on crates.io as [`oxidocs`](https://crates.io/crates/oxidocs).

## What it does

- Parses OOXML WordprocessingML (`.docx`) into a structured document model
- Lays out documents with a line/page model measured against Microsoft
  Word's own rendering — including Japanese typography: kinsoku line
  breaking, document grid (docGrid), ruby, vertical writing
- Edits text with incremental re-layout, and exports to PDF (embedded CJK
  fonts included)

## How fidelity is measured

Layout is verified continuously against Word as the oracle: per-paragraph
pagination equality (100% on the current 87-document corpus of real-world
Japanese government/legal documents) and per-page SSIM pixel comparison.

## Usage

```js
import init, { parse_document, layout_document, docx_to_pdf } from 'oxidocs';

await init();
const bytes = new Uint8Array(await file.arrayBuffer());
const doc = parse_document(bytes);      // structured document model
const pages = layout_document(bytes);   // laid-out pages (positions in pt)
const pdf = docx_to_pdf(bytes);         // Uint8Array (PDF)
```

## License

MIT OR Apache-2.0 (bindings). The engine core is MPL-2.0.
