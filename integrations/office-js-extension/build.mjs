// SPDX-License-Identifier: MPL-2.0
// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

// Puts web/oxi-office.js beside the manifest, so the folder loads as an
// unpacked extension (chrome://extensions -> Load unpacked) or zips for a
// store. Run: node integrations/office-js-extension/build.mjs
import { copyFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
copyFileSync(join(here, '..', '..', 'web', 'oxi-office.js'), join(here, 'oxi-office.js'));
console.log('oxi-office.js copied');
