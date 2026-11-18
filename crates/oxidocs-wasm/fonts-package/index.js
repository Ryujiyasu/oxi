// SPDX-License-Identifier: MIT OR Apache-2.0
//
// @oxidocs/wasm-fonts: font metrics for the faces @oxidocs/wasm does not
// carry (about 740 Windows faces), as font plugins.
//
//   import fonts, { fontsFrom, fontsFromFile } from '@oxidocs/wasm-fonts';
//
// - `fonts` (the default export) fetches the members a document needs from
//   this version's font_catalog_metrics.gz on jsDelivr, by HTTP range request.
// - `fontsFrom(url)` does the same from a copy you host (an intranet server
//   works; one that ignores Range still works, it sends the whole file once).
// - `fontsFromFile(read)` reads from disk, for Node: `read(offset, length)`
//   returns the bytes, for example with fs.readSync on an opened file.
//
// Use the same version as @oxidocs/wasm: the module holds the catalog's index.

export const CATALOG_URL = 'https://cdn.jsdelivr.net/npm/@oxidocs/wasm-fonts@__VERSION__/font_catalog_metrics.gz';

export function fontsFrom(url, { fetch: fetchFn = (...a) => globalThis.fetch(...a) } = {}) {
  let whole = null; // the whole file, once a server has ignored Range
  const one = async (m) => {
    if (whole) return (await whole).subarray(m.offset, m.offset + m.length);
    const res = await fetchFn(url, { headers: { Range: `bytes=${m.offset}-${m.offset + m.length - 1}` } });
    if (res.status === 206) return new Uint8Array(await res.arrayBuffer());
    if (res.status === 200) {
      whole = res.arrayBuffer().then((b) => new Uint8Array(b));
      return (await whole).subarray(m.offset, m.offset + m.length);
    }
    throw new Error(`@oxidocs/wasm-fonts: HTTP ${res.status} for ${url}`);
  };
  let probed = false; // ask for one range first: a server that ignores Range is then asked once, not once per face
  const load = async (members) => {
    if (!probed && members.length) {
      const first = await one(members[0]);
      probed = true;
      return [first, ...await Promise.all(members.slice(1).map(one))];
    }
    return Promise.all(members.map(one));
  };
  return { name: `fonts from ${url}`, load };
}

export function fontsFromFile(read) {
  return { name: 'fonts from a file', load: async (members) => members.map((m) => read(m.offset, m.length)) };
}

export default fontsFrom(CATALOG_URL);
