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
//   wasm-pack build --target web --release -- --features suite
//   node prepare-npm.mjs
//   cd pkg && npm publish --access public
//
// The `rm -rf pkg` is not superstition: wasm-pack reads any package.json it
// finds there and its model of `repository` is a string, so a second build over
// the first one it wrote fails with "invalid type: map, expected a string" —
// after producing the files, which makes it look like a build problem.
import { copyFileSync, existsSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const pkg = join(here, 'pkg');
if (!existsSync(join(pkg, 'package.json'))) {
  throw new Error('build it first: wasm-pack build --target web --release -- --features suite');
}

copyFileSync(join(here, 'README.md'), join(pkg, 'README.md'));

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
    'README.md',
    'LICENSE-APACHE',
    'LICENSE-MIT',
  ],
});
// wasm-pack writes this for a build that has snippets; this one has none, and
// a bundler that believes it keeps dead weight alive.
delete manifest.sideEffects;

for (const name of manifest.files) {
  if (!existsSync(join(pkg, name))) {
    throw new Error(`${name} is listed to ship but is not in pkg/`);
  }
}

writeFileSync(join(pkg, 'package.json'), `${JSON.stringify(manifest, null, 2)}\n`);
console.log(`pkg/ is ready to publish as ${manifest.name}@${manifest.version}`);
