// Assembles dist-desktop/ (the gitignored frontendDist) from the tracked web
// editors, so local and CI builds always embed the current editors + engine.
// Runs as tauri's beforeBuildCommand.
//
// The desktop app ships all three editors behind a launcher; the public site
// keeps its own pages (docs/sheets.html is a viewer landing page there, so the
// spreadsheet editor is taken from web/xlsx-demo.html instead). analytics.js is
// deliberately never copied: the desktop app ships without telemetry, and the
// assertion below fails the build if a reference to it survives.
import { cpSync, existsSync, mkdirSync, readdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const repo = join(here, '..', '..');
const docs = join(repo, 'docs');
const site = join(repo, 'web');
const dist = join(repo, 'dist-desktop');
const web = join(dist, 'web');

rmSync(dist, { recursive: true, force: true });
mkdirSync(join(web, 'samples'), { recursive: true });

// The engine and the assets the editors import by relative path.
for (const name of [
  'vertical-text.js',
  'oxidocs_wasm.js',
  'oxidocs_wasm_bg.wasm',
  'favicon.ico',
  'favicon-32x32.png',
]) {
  cpSync(join(docs, name), join(web, name));
}
// The spreadsheet editor's row model, its VBA worker, and the workbook it
// opens on start.
for (const name of ['row-model.js', 'vba-runner.js', 'vba-worker.js']) {
  cpSync(join(site, name), join(web, name));
}
// The app opens on a workbook of its own rather than on a sales sample, so the
// empty one is what travels with it.
cpSync(join(here, 'blank.xlsx'), join(web, 'samples', 'blank.xlsx'));
cpSync(join(here, 'blank.pptx'), join(web, 'blank.pptx'));

// The version the bundle is built at, which the pages carry so a rating can say
// what it is a rating of.
const { version } = JSON.parse(readFileSync(join(here, 'tauri.conf.json'), 'utf8'));

// Asking what someone makes of Oxi only makes sense when there is somewhere for
// the answer to go. Without OXI_FEEDBACK_ENDPOINT the card is left out of the
// build entirely, rather than shipped to fail at the moment someone bothers to
// answer it.
const asking = (process.env.OXI_FEEDBACK_ENDPOINT || '').trim();
if (asking) {
  writeFileSync(
    join(web, 'rate.js'),
    readFileSync(join(site, 'rate.js'), 'utf8').replace('{{FEEDBACK_ENDPOINT}}', asking),
  );
}

// The editors, under the names the launcher and the native menu navigate to.
for (const [from, name] of [
  [join(docs, 'docs.html'), 'docs.html'],
  [join(site, 'xlsx-demo.html'), 'sheets.html'],
  [join(docs, 'slides.html'), 'slides.html'],
]) {
  let html = readFileSync(from, 'utf8').replace(
    /^[ \t]*<script\b[^>]*\banalytics\.js\b[^>]*>\s*<\/script>[ \t]*\r?\n?/m,
    '',
  );
  if (/analytics\.js|gtag\(|dataLayer/.test(html)) {
    throw new Error(`${name}: telemetry survived the strip — the desktop app ships without it`);
  }
  // The same page serves the site, where a visitor should be shown something
  // straight away, and the app, where a person came with a file of their own.
  // The pages read this to tell which they are in.
  html = html.replace(
    /<\/head>/,
    '<script>document.documentElement.dataset.host = "desktop";</script>\n'
      + `<meta name="oxi-version" content="${version}">\n`
      + (asking ? '<script src="./rate.js" defer></script>\n' : '')
      + '</head>',
  );
  if (!/dataset\.host = "desktop"/.test(html)) {
    throw new Error(`${name}: nowhere to mark the page as the desktop app's`);
  }
  writeFileSync(join(web, name), html);
}

// The launcher, carrying the version the bundle is built at.
writeFileSync(
  join(web, 'index.html'),
  readFileSync(join(here, 'launcher.html'), 'utf8').replaceAll('{{VERSION}}', version),
);
// An editor is bundled with whichever engine build is tracked, and the two can
// disagree: a page that imports a function the engine does not export dies at
// its first line, taking the whole editor with it, and nothing about the bundle
// looks wrong from the outside. So the bundle is asked to account for itself —
// every name imported has to be exported, and every file referred to has to be
// here — while the build can still fail rather than after it ships.
const engine = readFileSync(join(web, 'oxidocs_wasm.js'), 'utf8');
const exported = new Set(
  [...engine.matchAll(/export (?:async )?function (\w+)/g)].map((match) => match[1]),
);
const faults = [];
for (const name of readdirSync(web).filter((f) => /\.(html|js)$/.test(f))) {
  const text = readFileSync(join(web, name), 'utf8');

  for (const [, names] of text.matchAll(
    /import\s+(?:\w+\s*,\s*)?\{([^}]*)\}\s*from\s*'\.\/oxidocs_wasm\.js'/g,
  )) {
    for (const raw of names.split(',')) {
      const wanted = raw.trim().split(/\s+as\s+/)[0].trim();
      if (wanted && !exported.has(wanted)) {
        faults.push(`${name} imports ${wanted}, which the bundled engine does not export`);
      }
    }
  }

  const referred = [
    ...[...text.matchAll(/from\s*'(\.\/[^']+)'/g)].map((match) => match[1]),
    ...[...text.matchAll(/(?:src|href)="(\.\/[^"#]+)"/g)].map((match) => match[1]),
    ...[...text.matchAll(/new URL\('(\.\/[^']+)'/g)].map((match) => match[1]),
  ];
  for (const path of new Set(referred)) {
    if (!existsSync(join(web, path))) {
      faults.push(`${name} refers to ${path}, which is not in the bundle`);
    }
  }
}
if (faults.length) {
  throw new Error(`the desktop bundle does not hold together:\n  ${faults.join('\n  ')}`);
}

writeFileSync(
  join(dist, 'index.html'),
  '<!DOCTYPE html><html><head><meta charset="utf-8">' +
    '<meta http-equiv="refresh" content="0;url=web/index.html"></head><body></body></html>\n',
);
console.log(`dist-desktop assembled: launcher + 3 editors (v${version})`);
