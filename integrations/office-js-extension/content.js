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
