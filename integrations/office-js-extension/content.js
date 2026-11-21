// SPDX-License-Identifier: MPL-2.0
// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

// On an Oxi editor page: ask the background to answer office.js with
// oxi-office.js for this tab, and tell the page the extension is there (the
// editor's add-in prompt says so).
chrome.runtime.sendMessage({ oxi: 'editor-open' }, answer => {
  if (answer?.ok) document.documentElement.dataset.oxiOfficeExtension = 'on';
});

// XLL requests from the page (window.OxiXll in the editor) go to the
// background, which talks to the local oxi-xll-host; answers come back the
// same way. Only messages from this page itself are taken.
window.addEventListener('message', event => {
  const data = event.data;
  if (event.source !== window || !data || data.oxiXll !== 'request') return;
  const { id, oxiXll: _, kind = 'xll', ...request } = data;
  chrome.runtime.sendMessage({ oxi: kind === 'com' ? 'com' : 'xll', request }, answer => {
    window.postMessage({ oxiXll: 'reply', id, ...(answer ?? { error: chrome.runtime.lastError?.message ?? 'no answer' }) }, '*');
  });
});
