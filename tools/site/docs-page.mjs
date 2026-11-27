// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

// The site's Word editor is web/docs.html, the same page the VS Code extension
// and the desktop app use; there is no second copy of it under docs/. This
// writes it into the assembled site with what only the site needs:
//
// - its title, description, social card, canonical URL and icon;
// - the visit counter (analytics.js);
// - Oxi's own colours, where the extension follows Microsoft's;
// - the logo linking to the site's front page, and a link to the source.
//
// web/docs.html loads the site's shared modules as '../docs/name.js'. The
// site serves docs/ at its root, so those become './name.js'.
//
//   node tools/site/docs-page.mjs <site dir>   (public/ in site.yml)

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..');
const out = process.argv[2];
if (!out) throw new Error('usage: node tools/site/docs-page.mjs <site dir>');

let html = fs.readFileSync(path.join(root, 'web/docs.html'), 'utf8');

function swap(what, from, to) {
  if (!html.includes(from)) throw new Error(`web/docs.html: ${what} not found; update tools/site/docs-page.mjs`);
  html = html.replace(from, to);
}

html = html.replace(/(['"])\.\.\/docs\//g, '$1./');

swap('the <title>', '<title>Oxidocs</title>', `<title>Oxidocs — Free Online Word (.docx) Viewer &amp; Editor in Your Browser</title>
<meta name="description" content="Open, view and edit Word .docx files directly in your browser with Word-faithful layout. Free, open-source, Rust + WebAssembly — no upload, no account, nothing leaves your device.">
<meta property="og:type" content="website">
<meta property="og:title" content="Oxidocs — Free Online Word (.docx) Viewer &amp; Editor">
<meta property="og:description" content="Open, view and edit .docx files in the browser with Word-faithful layout. Free, open-source, no upload.">
<meta property="og:image" content="https://oxi-dd65f4.gitlab.io/oxi-logo.png">
<meta property="og:url" content="https://oxi-dd65f4.gitlab.io/docs.html">
<meta property="og:site_name" content="Oxi">
<meta name="twitter:card" content="summary_large_image">
<link rel="canonical" href="https://oxi-dd65f4.gitlab.io/docs.html">
<link rel="icon" type="image/x-icon" href="./favicon.ico">
<script src="./analytics.js" defer></script>`);

// After the page's own styles, so these win.
swap('the end of the page styles', '</style>', `</style>
<style>
/* The site: Oxi's colours (5b940e9f), not the Microsoft ones the extension wears. */
:root{--accent:#6A4FA3;--accent-h:#57408A;--accent-l:#EFEBF7}
.tb a.logo{color:inherit;text-decoration:none}
.tb a.logo:hover{opacity:.8}
.tb .ra a.src{color:rgba(255,255,255,.9);text-decoration:none;padding:3px 7px;border-radius:3px;font-size:12px}
.tb .ra a.src:hover{background:rgba(255,255,255,.18)}
</style>`);

swap('the logo', '<span class="logo">Oxidocs</span>',
  '<a class="logo" href="./index.html" title="Oxi">Oxidocs</a>');
swap('the title bar buttons', '<button id="btnPdf">PDF</button>',
  '<button id="btnPdf">PDF</button>\n    <a class="src" href="https://github.com/Ryujiyasu/oxi" target="_blank" rel="noopener" data-i18n="source">Source</a>');

// Every module the page loads must be there.
for (const m of html.matchAll(/(?:from|import)\s*\(?\s*['"]\.\/([\w.-]+\.js)['"]/g)) {
  if (!fs.existsSync(path.join(out, m[1]))) throw new Error(`docs.html loads ./${m[1]}, which the site does not have`);
}

fs.writeFileSync(path.join(out, 'docs.html'), html);
console.log(`docs.html <- web/docs.html (+ site head, colours, links): ${html.length} bytes`);
