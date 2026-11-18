# @oxidocs/wasm-fonts

Font metrics for [@oxidocs/wasm](https://www.npmjs.com/package/@oxidocs/wasm), as font plugins.

`@oxidocs/wasm` carries the metrics of the common faces (Yu Gothic, Meiryo, the MS faces, Calibri, Cambria and the like). The metrics of about 740 other faces that Windows ships (HG fonts, UD Digi Kyokasho, BIZ UD, many Latin, CJK and other-script faces) are in `font_catalog_metrics.gz` here, one gzip member per face. A font plugin hands the engine exactly the faces a document uses.

These are measured metrics (advance widths, vertical metrics, glyph bounds), not font programs: no Microsoft font is in this package.

```js
import init, * as oxi from '@oxidocs/wasm';
import { useFonts, withFonts } from '@oxidocs/wasm/fonts.js';
import fonts from '@oxidocs/wasm-fonts';

await init();
useFonts(oxi, fonts);                                      // online: jsDelivr, range requests
const pdf = await withFonts(oxi, () => oxi.docx_to_pdf(docx));
```

- **Your own server** (an intranet without the internet included): copy `font_catalog_metrics.gz` there and use `fontsFrom('https://intranet.example/oxi/font_catalog_metrics.gz')`.
- **Node**: `fontsFromFile((offset, length) => …)` reads the ranges from disk.

Use the same version as `@oxidocs/wasm`.

MIT OR Apache-2.0.
