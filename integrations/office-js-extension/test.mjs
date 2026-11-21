// SPDX-License-Identifier: MPL-2.0
// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

// The redirect rule's pattern against the URLs add-ins load office.js from.
// Run: node integrations/office-js-extension/test.mjs
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';

const source = readFileSync(new URL('./background.js', import.meta.url), 'utf8');
const raw = source.match(/const OFFICE_JS = '([^']+)'/)[1];
// In the source, the string literal writes each backslash twice.
const pattern = new RegExp(raw.split('\\\\').join('\\'));

for (const url of [
  'https://appsforoffice.microsoft.com/lib/1/hosted/office.js',
  'https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js',
  'https://appsforoffice.microsoft.com/lib/beta/hosted/office.js',
  'https://appsforoffice.microsoft.com/lib/1/hosted/office.debug.js',
  'https://appsforoffice.microsoft.com/lib/1/hosted/office.js?v=2',
]) assert.ok(pattern.test(url), url);

for (const url of [
  'https://appsforoffice.microsoft.com/lib/1/hosted/excel-web-16.00.js',
  'https://example.com/lib/1/hosted/office.js',
  'https://appsforoffice.microsoft.com/lib/1/hosted/office.js.map',
]) assert.ok(!pattern.test(url), url);

const manifest = JSON.parse(readFileSync(new URL('./manifest.json', import.meta.url), 'utf8'));
assert.equal(manifest.manifest_version, 3);
assert.deepEqual(manifest.web_accessible_resources[0].resources, ['oxi-office.js']);
console.log('ok - office.js URLs redirect, others do not');
