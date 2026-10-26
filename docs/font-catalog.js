// The engine module built for this site leaves its per-face font metric
// catalog (~34 MB) out. A layout records the catalog members it needed and did
// not have; this fetches exactly those members, registers them, and the caller
// lays out again. A bundled build reports nothing missing, so the same code
// runs against either build.
//
// Each member is published as its own file, font_catalog/<offset>.bin
// (tools/site/split_font_catalog.sh): GitLab Pages answers a Range request
// with the whole 34 MB file, so byte ranges of one file would not save anything.

const CATALOG_DIR = new URL('./font_catalog/', import.meta.url);

async function fetchMember(m) {
  const r = await fetch(new URL(`${m.offset}.bin`, CATALOG_DIR));
  if (!r.ok) throw new Error(`font catalog: HTTP ${r.status} for ${m.key}`);
  return new Uint8Array(await r.arrayBuffer());
}

// Fetch and register what the last layout was missing. True if anything was.
export async function loadMissingFonts(wasm) {
  if (typeof wasm.font_catalog_missing !== 'function') return false;
  const missing = wasm.font_catalog_missing();
  if (!missing.length) return false;
  const parts = await Promise.all(missing.map(fetchMember));
  missing.forEach((m, i) => {
    if (!wasm.font_catalog_register(m.key, parts[i])) {
      throw new Error(`font catalog: member ${m.key} was rejected`);
    }
  });
  return true;
}

// Run a layout-producing call until no catalog member is missing.
// Faces found missing in one pass are registered before the next; two passes
// are the usual case, four is a safety bound.
export async function withFonts(wasm, call) {
  let result = call();
  for (let pass = 0; pass < 4 && await loadMissingFonts(wasm); pass++) {
    result = call();
  }
  return result;
}
