// Turns what wasm-pack leaves in `pkg/` into the package npm actually ships.
//
// wasm-pack names the package after the crate, which on npm reads as a Rust
// crate that wandered into the wrong registry, and it lists neither the README
// nor the licences. Doing that by hand once is fine; doing it by hand every
// release is how a package ends up published under the wrong name at two in the
// morning. So it is written down.
//
//   cd crates/oxidocs-wasm
//   rm -rf pkg                       # wasm-pack chokes on its own leftovers
//   wasm-pack build --target web --release -- --features suite,external-font-catalog
//   node prepare-npm.mjs --version 0.10.0
//   node prepare-npm-fonts.mjs --version 0.10.0
//   (cd pkg && npm publish --access public) && (cd pkg-fonts && npm publish --access public)
//
// The module leaves the metric catalog of the uncommon faces out
// (external-font-catalog); @oxidocs/wasm-fonts carries it, and fonts.js,
// shipped here, is how a page plugs it in. Publish both at the same version:
// the module holds the catalog's index.
//
// The `rm -rf pkg` is not superstition: wasm-pack reads any package.json it
// finds there and its model of `repository` is a string, so a second build over
// the first one it wrote fails with "invalid type: map, expected a string" —
// after producing the files, which makes it look like a build problem.
import { copyFileSync, existsSync, openSync, readFileSync, readSync, closeSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const pkg = join(here, 'pkg');
if (!existsSync(join(pkg, 'package.json'))) {
  throw new Error('build it first: wasm-pack build --target web --release -- --features suite');
}

copyFileSync(join(here, 'README.md'), join(pkg, 'README.md'));
copyFileSync(join(here, 'fonts.js'), join(pkg, 'fonts.js'));

// A module built without external-font-catalog has the catalog compiled in;
// published as this package it would weigh 34 MB more and never use the
// plugins. The catalog's first bytes inside the .wasm give it away.
const catalog = join(here, '..', 'oxidocs-core', 'src', 'font', 'data', 'font_catalog_metrics.gz');
const head = Buffer.alloc(4096);
const fd = openSync(catalog, 'r'); readSync(fd, head, 0, head.length, 0); closeSync(fd);
if (readFileSync(join(pkg, 'oxidocs_wasm_bg.wasm')).includes(head)) {
  throw new Error('this build has the font catalog inside: build with --features suite,external-font-catalog');
}
const version = (process.argv.indexOf('--version') > 0) ? process.argv[process.argv.indexOf('--version') + 1] : null;

const manifest = JSON.parse(readFileSync(join(pkg, 'package.json'), 'utf8'));
Object.assign(manifest, {
  // The scope is the project's; `wasm` says which of its builds this is.
  name: '@oxidocs/wasm',
  description: 'Open .docx, .xlsx and .pptx in the browser and lay them out '
    + 'the way Office does — WebAssembly bindings for Oxi',
  repository: { type: 'git', url: 'git+https://github.com/Ryujiyasu/oxi.git' },
  homepage: 'https://github.com/Ryujiyasu/oxi',
  bugs: { url: 'https://github.com/Ryujiyasu/oxi/issues' },
  keywords: ['docx', 'xlsx', 'pptx', 'ooxml', 'office', 'wasm', 'word', 'layout'],
  // The licences and the README travel with it, or the npm page is bare and
  // the terms live only in a repository nobody opened.
  files: [
    'oxidocs_wasm_bg.wasm',
    'oxidocs_wasm.js',
    'oxidocs_wasm.d.ts',
    'oxidocs_wasm_bg.wasm.d.ts',
    'fonts.js',
    'README.md',
    'LICENSE-APACHE',
    'LICENSE-MIT',
  ],
});
// wasm-pack writes this for a build that has snippets; this one has none, and
// a bundler that believes it keeps dead weight alive.
delete manifest.sideEffects;
// The npm version can run ahead of the crate's (the crate's is the engine's
// measured score); --version sets it.
if (version) manifest.version = version;

for (const name of manifest.files) {
  if (!existsSync(join(pkg, name))) {
    throw new Error(`${name} is listed to ship but is not in pkg/`);
  }
}

writeFileSync(join(pkg, 'package.json'), `${JSON.stringify(manifest, null, 2)}\n`);
console.log(`pkg/ is ready to publish as ${manifest.name}@${manifest.version}`);
