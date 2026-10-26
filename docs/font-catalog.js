// The engine module built for this site leaves its per-face font metric
// catalog (~34 MB) out. A layout records the catalog members it needed and did
// not have; this fetches exactly those byte ranges of the catalog file,
// registers them, and the caller lays out again. A bundled build reports
// nothing missing, so the same code runs against either build.

const CATALOG_URL = new URL('./font_catalog_metrics.bin', import.meta.url);
let wholeFile = null; // only when the server ignores Range and sends the whole file

async function fetchMember(m) {
  const end = m.offset + m.length - 1;
  const r = await fetch(CATALOG_URL, { headers: { Range: `bytes=${m.offset}-${end}` } });
  if (!r.ok) throw new Error(`font catalog: HTTP ${r.status}`);
  if (r.status === 206) return new Uint8Array(await r.arrayBuffer());
  wholeFile ??= new Uint8Array(await r.arrayBuffer());
  return wholeFile.slice(m.offset, m.offset + m.length);
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
