// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Font plugins for @oxidocs/wasm.
//
// The module carries the metrics of the common faces (Yu Gothic, Meiryo, the
// MS faces, Calibri, Cambria, …). Every other face comes from a font plugin:
// a layout records the faces it asked for and did not have, `withFonts` asks
// the plugins for them, registers what they give and lays the document out
// again. Registered faces stay, so later conversions using them cost nothing.
//
//   import init, * as oxi from '@oxidocs/wasm';
//   import { useFonts, withFonts } from '@oxidocs/wasm/fonts.js';
//   import fonts from '@oxidocs/wasm-fonts';      // online: fetches only what a document uses
//   await init();
//   useFonts(oxi, fonts);
//   const pdf = await withFonts(oxi, () => oxi.docx_to_pdf(bytes));
//
// A plugin is `{ name, load(members) }`: given `[{ key, offset, length }]`
// (byte ranges in font_catalog_metrics.gz), it resolves to the bytes of each,
// or `null` for one it does not have; the next plugin is asked for those.

const plugins = new WeakMap(); // module -> [plugin]

/** Add a font plugin to this module. Plugins are asked in the order added. */
export function useFonts(api, plugin) {
  if (!plugin || typeof plugin.load !== 'function') throw new TypeError('useFonts: a font plugin has a load(members) function');
  const list = plugins.get(api) || [];
  if (!list.includes(plugin)) list.push(plugin);
  plugins.set(api, list);
  return api;
}

/**
 * Run `convert` (any call into the module that lays a document out), let the
 * font plugins supply the faces it was missing, and run it again until
 * nothing is missing or no plugin has what is. Returns the last result.
 */
export async function withFonts(api, convert, { maxRounds = 4 } = {}) {
  let result = await convert();
  if (!api.font_catalog_missing) return result; // a build with every face inside
  for (let round = 0; round < maxRounds; round++) {
    let missing = api.font_catalog_missing();
    if (!missing.length) return result;
    let added = 0;
    for (const plugin of plugins.get(api) || []) {
      if (!missing.length) break;
      const bytes = await plugin.load(missing);
      const left = [];
      missing.forEach((m, i) => {
        if (bytes && bytes[i] && api.font_catalog_register(m.key, bytes[i])) added++;
        else left.push(m);
      });
      missing = left;
    }
    if (!added) return result; // nobody had them: the faces fall back as before
    result = await convert();
  }
  return result;
}
