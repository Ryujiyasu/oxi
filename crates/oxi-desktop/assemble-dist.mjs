// Assembles dist-desktop/ (the gitignored frontendDist) from the tracked web
// editor in docs/, so local and CI builds always embed the current editor +
// engine. Runs as tauri's beforeBuildCommand. analytics.js is deliberately
// not copied: the desktop app ships without telemetry.
import { cpSync, mkdirSync, rmSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const repo = join(here, '..', '..');
const docs = join(repo, 'docs');
const dist = join(repo, 'dist-desktop');
const web = join(dist, 'web');

rmSync(dist, { recursive: true, force: true });
mkdirSync(web, { recursive: true });
for (const name of [
  'docs.html',
  'oxidocs_wasm.js',
  'oxidocs_wasm_bg.wasm',
  'favicon.ico',
  'favicon-32x32.png',
]) {
  cpSync(join(docs, name), join(web, name));
}
writeFileSync(
  join(dist, 'index.html'),
  '<!DOCTYPE html><html><head><meta charset="utf-8">' +
    '<meta http-equiv="refresh" content="0;url=web/docs.html"></head><body></body></html>\n',
);
console.log('dist-desktop assembled from docs/');
