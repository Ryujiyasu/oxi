// SPDX-License-Identifier: MPL-2.0
// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

// An Office add-in is a web page that loads Microsoft's office.js from
// appsforoffice.microsoft.com. Opened in the Oxi editor's add-in pane, it
// should talk to the workbook on the page instead, so for that tab -- and
// only that tab -- the request for office.js is answered with oxi-office.js.
// Everywhere else (Excel on the web, other sites) office.js loads as usual.

const OFFICE_JS = '^https://appsforoffice\\.microsoft\\.com/lib/[^/]+/hosted/office(\\.debug)?\\.js(\\?.*)?$';

// One session rule per editor tab; its id is the tab's.
async function redirectFor(tabId) {
  await chrome.declarativeNetRequest.updateSessionRules({
    removeRuleIds: [tabId],
    addRules: [{
      id: tabId,
      priority: 1,
      action: { type: 'redirect', redirect: { extensionPath: '/oxi-office.js' } },
      condition: { regexFilter: OFFICE_JS, tabIds: [tabId], resourceTypes: ['script'] },
    }],
  });
}

async function forget(tabId) {
  await chrome.declarativeNetRequest.updateSessionRules({ removeRuleIds: [tabId] });
}

chrome.runtime.onMessage.addListener((message, sender, reply) => {
  if (message?.oxi === 'editor-open' && sender.tab?.id != null) {
    redirectFor(sender.tab.id).then(() => reply({ ok: true }), error => reply({ ok: false, error: String(error) }));
    return true;
  }
  return false;
});

chrome.tabs.onRemoved.addListener(tabId => { forget(tabId); });
