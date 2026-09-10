// Cuts every icon Oxi ships from the vector masters in crates/oxi-desktop/icons.
//
// Run it after editing icon.svg (or either hinted cut) and commit what it
// writes. The point of having it is that the sizes cannot drift apart again:
// before this existed the set had a 1280x714 banner standing in for a 256px
// icon, a 64px image named 128x128.png, and an .ico holding only 16 and 32, so
// Windows was interpolating almost every icon it drew.
//
//   node tools/make-icons.mjs
//
// Chrome does the rasterising, so the small sizes come out of the hinted cuts
// exactly as they were drawn, pixel for pixel.
import { createRequire } from 'node:module';
import { existsSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const repo = join(here, '..');
const icons = join(repo, 'crates', 'oxi-desktop', 'icons');

const require = createRequire(join(repo, 'tools', 'oxi-chromium-renderer', 'package.json'));
const puppeteer = require('puppeteer');

const BROWSERS = [
  process.env.OXI_BROWSER,
  'C:/Program Files/Google/Chrome/Application/chrome.exe',
  'C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe',
  '/usr/bin/google-chrome',
  '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
].filter(Boolean);

// At and below 32px the master's rules land between pixels — 1.875px at 32 —
// and grey both their edges, so those sizes come from cuts drawn on their own
// pixel grid instead. From 48 up the rules are wide enough that the softening
// does not read.
const CUT = { 16: 'icon-16.svg', 24: 'icon-24.svg', 32: 'icon-32.svg' };
const SIZES = [16, 24, 32, 48, 64, 128, 256];

const executablePath = BROWSERS.find((path) => existsSync(path));
if (!executablePath) {
  throw new Error(`no browser to rasterise with — set OXI_BROWSER to one of: ${BROWSERS.join(', ')}`);
}

const browser = await puppeteer.launch({ headless: true, executablePath, args: ['--no-sandbox'] });
const page = await browser.newPage();

async function cut(size) {
  const source = readFileSync(join(icons, CUT[size] ?? 'icon.svg'), 'utf8')
    .replace(/(<svg\b[^>]*?)\swidth="[^"]*"/, '$1')
    .replace(/(<svg\b[^>]*?)\sheight="[^"]*"/, '$1')
    .replace(/<svg\b/, `<svg width="${size}" height="${size}"`);
  await page.setViewport({ width: size, height: size, deviceScaleFactor: 1 });
  await page.setContent(`<body style="margin:0;line-height:0">${source}`, { waitUntil: 'load' });
  return page.screenshot({ omitBackground: true, type: 'png' });
}

const png = {};
for (const size of SIZES) {
  png[size] = await cut(size);
  console.log(`  ${String(size).padStart(3)}px  ${String(png[size].length).padStart(6)}B  ${CUT[size] ?? 'icon.svg'}`);
}
// The one size nothing else asks for: what iOS puts on a home screen.
const touch = await cut(180);
await browser.close();

/// An ICONDIR, one ICONDIRENTRY per size, then the images. Each image is kept
/// as PNG, which Windows has read since Vista and which keeps the 256 from
/// costing a quarter of a megabyte.
function ico(sizes) {
  const head = Buffer.alloc(6 + 16 * sizes.length);
  head.writeUInt16LE(0, 0);
  head.writeUInt16LE(1, 2);
  head.writeUInt16LE(sizes.length, 4);
  let offset = head.length;
  sizes.forEach((size, i) => {
    const at = 6 + 16 * i;
    head[at] = size >= 256 ? 0 : size;
    head[at + 1] = size >= 256 ? 0 : size;
    head.writeUInt16LE(1, at + 4);
    head.writeUInt16LE(32, at + 6);
    head.writeUInt32LE(png[size].length, at + 8);
    head.writeUInt32LE(offset, at + 12);
    offset += png[size].length;
  });
  return Buffer.concat([head, ...sizes.map((size) => png[size])]);
}

const written = [
  // What the bundler is pointed at, at the sizes their names promise.
  [join(icons, '32x32.png'), png[32]],
  [join(icons, '128x128.png'), png[128]],
  [join(icons, '128x128@2x.png'), png[256]],
  [join(icons, 'icon.ico'), ico(SIZES)],
  // The site, and the copy the desktop bundle carries beside the editors.
  [join(repo, 'docs', 'favicon.ico'), ico([16, 24, 32, 48])],
  [join(repo, 'docs', 'favicon-32x32.png'), png[32]],
  [join(repo, 'docs', 'icon-64.png'), png[64]],
  [join(repo, 'docs', 'apple-touch-icon.png'), touch],
  [join(repo, 'web', 'favicon.ico'), ico([16, 24, 32, 48])],
];
for (const [path, bytes] of written) {
  writeFileSync(path, bytes);
  console.log(`  ${path.slice(repo.length + 1).replace(/\\/g, '/')}  ${(bytes.length / 1024).toFixed(1)}KB`);
}
