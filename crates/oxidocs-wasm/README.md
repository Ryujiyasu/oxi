# @oxidocs/wasm

WebAssembly bindings for [Oxi](https://github.com/Ryujiyasu/oxi) — a document
engine that opens `.docx`, `.xlsx` and `.pptx` in the browser with no server,
and lays them out the way Microsoft Office lays them out.

The layout is scored page by page against Word's own render, on blind sets of
documents the engine has never been tuned against. The numbers, the corpora and
the method are in the [repository README](https://github.com/Ryujiyasu/oxi).

## A word about size

The `.wasm` is about 62 MB (24 MB gzip-compressed, as browsers and CDNs usually
send it). It carries the metrics of the common faces — Yu Gothic, Yu Mincho,
Meiryo, the MS faces, Calibri, Cambria and the like — so ordinary documents lay
out with nothing else. The metrics of about 740 other faces that Windows ships
(HG fonts, UD Digi Kyokasho, many Latin, CJK and other-script faces) are in
[@oxidocs/wasm-fonts](https://www.npmjs.com/package/@oxidocs/wasm-fonts) and
come in through a font plugin, a face at a time, when a document uses one.
Load the module once and cache it.

## Use

```js
import init, * as oxi from '@oxidocs/wasm';
import { useFonts, withFonts } from '@oxidocs/wasm/fonts.js';
import fonts from '@oxidocs/wasm-fonts';

await init();
useFonts(oxi, fonts);   // the uncommon faces, fetched per face from jsDelivr

const docx = new Uint8Array(await file.arrayBuffer());
const pdf = await withFonts(oxi, () => oxi.docx_to_pdf(docx));
const pages = await withFonts(oxi, () => oxi.layout_document(docx));
```

`withFonts` runs the call, lets the plugins supply the faces it was missing,
and runs it again. Calling the module directly works too; a face neither built
in nor supplied by a plugin is then laid out with fallback metrics.

Font plugins in `@oxidocs/wasm-fonts`:

- the default export: online, range requests to this version's catalog on jsDelivr;
- `fontsFrom(url)`: a copy you host — an intranet server works, so does one without range support;
- `fontsFromFile(read)`: from disk, for Node.

Use the same version of both packages: the module holds the catalog's index.

Every export is typed; `oxidocs_wasm.d.ts` travels with the package.

The module is built for the `web` target, so it wants a real URL: serve it over
HTTP rather than opening the page off `file://`, or the browser will refuse to
instantiate it.

## What it can do

- **.docx** — parse, lay out, render, edit and export to PDF. Japanese
  typography included: kinsoku by JIS X 4051, ruby, vertical writing.
- **.xlsx** — parse, draw, a formula engine, export of the print pages to
  PDF (`xlsx_to_pdf`), and a VBA host that runs a workbook's own macros in
  the browser.
- **.pptx** — parse, draw, edit, present.
- **Round-trip editing** — the original ZIP is kept and only the XML that
  changed is patched, so a save that changes nothing is byte-identical.

## Licence

MIT OR Apache-2.0, the Rust convention. The engine crates behind it are
MPL-2.0.
