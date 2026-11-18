// Build pkg-fonts/, the @oxidocs/wasm-fonts package: the metric catalog of the
// faces @oxidocs/wasm leaves out, and the font plugins that hand it over.
//
//   node prepare-npm-fonts.mjs --version 0.10.0
//   cd pkg-fonts && npm publish --access public
//
// Publish it at the same version as @oxidocs/wasm (see prepare-npm.mjs).
import { copyFileSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const at = process.argv.indexOf('--version');
const version = at > 0 ? process.argv[at + 1] : null;
if (!version) throw new Error('give the version: node prepare-npm-fonts.mjs --version 0.10.0');

const out = join(here, 'pkg-fonts');
rmSync(out, { recursive: true, force: true });
mkdirSync(out);
copyFileSync(join(here, '..', 'oxidocs-core', 'src', 'font', 'data', 'font_catalog_metrics.gz'), join(out, 'font_catalog_metrics.gz'));
writeFileSync(join(out, 'index.js'), readFileSync(join(here, 'fonts-package', 'index.js'), 'utf8').replaceAll('__VERSION__', version));
copyFileSync(join(here, 'fonts-package', 'README.md'), join(out, 'README.md'));
for (const f of ['LICENSE-MIT', 'LICENSE-APACHE']) copyFileSync(join(here, f), join(out, f));
const manifest = {
  name: '@oxidocs/wasm-fonts',
  version,
  description: 'Font metrics for @oxidocs/wasm: about 740 Windows faces as a font plugin, fetched per face on demand',
  license: 'MIT OR Apache-2.0',
  type: 'module',
  main: 'index.js',
  exports: { '.': './index.js', './font_catalog_metrics.gz': './font_catalog_metrics.gz' },
  files: ['index.js', 'font_catalog_metrics.gz', 'README.md', 'LICENSE-MIT', 'LICENSE-APACHE'],
  repository: { type: 'git', url: 'git+https://github.com/Ryujiyasu/oxi.git' },
  homepage: 'https://github.com/Ryujiyasu/oxi',
  keywords: ['oxi', 'oxidocs', 'fonts', 'metrics', 'docx', 'wasm'],
};
writeFileSync(join(out, 'package.json'), `${JSON.stringify(manifest, null, 2)}\n`);
console.log(`pkg-fonts/ is ready to publish as ${manifest.name}@${manifest.version}`);
