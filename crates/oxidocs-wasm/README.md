# @oxidocs/wasm

WebAssembly bindings for [Oxi](https://github.com/Ryujiyasu/oxi) — a document
engine that opens `.docx`, `.xlsx` and `.pptx` in the browser with no server,
and lays them out the way Microsoft Office lays them out.

The layout is scored page by page against Word's own render, on blind sets of
documents the engine has never been tuned against. The numbers, the corpora and
the method are in the [repository README](https://github.com/Ryujiyasu/oxi).

## A word about size

The `.wasm` is about 37 MB, and roughly four fifths of that is data rather than
code: font metrics, typography tables and the measured constants the layout is
derived from. Splitting the package by document format was measured and does not
help — a docs-only build is 31 MB against 35 MB for all three formats, because
the tables are shared. Load it once and cache it.

## Use

```js
import init, { layout_document, parse_spreadsheet } from '@oxidocs/wasm';

await init();

const docx = new Uint8Array(await file.arrayBuffer());
const pages = layout_document(docx);   // laid-out pages, ready to draw
```

Every export is typed; `oxidocs_wasm.d.ts` travels with the package.

The module is built for the `web` target, so it wants a real URL: serve it over
HTTP rather than opening the page off `file://`, or the browser will refuse to
instantiate it.

## What it can do

- **.docx** — parse, lay out, render, edit and export to PDF. Japanese
  typography included: kinsoku by JIS X 4051, ruby, vertical writing.
- **.xlsx** — parse, draw, a formula engine, and a VBA host that runs a
  workbook's own macros in the browser.
- **.pptx** — parse, draw, edit, present.
- **Round-trip editing** — the original ZIP is kept and only the XML that
  changed is patched, so a save that changes nothing is byte-identical.

## Licence

MIT OR Apache-2.0, the Rust convention. The engine crates behind it are
MPL-2.0.
