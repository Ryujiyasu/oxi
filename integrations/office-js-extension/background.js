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

// XLL and COM add-ins: the editor page asks (through the content script)
// for a call; the local oxi-xll-host (XLL) or oxi-com-host (COM Automation
// add-ins) answers over native messaging. One host process of each kind,
// kept while requests come.
const NATIVE_HOSTS = { xll: 'com.oxi.xll_host', com: 'com.oxi.com_host' };
const ports = {};
const waiting = new Map();
let nextId = 1;
function nativePort(kind) {
  if (ports[kind]) return ports[kind];
  const port = chrome.runtime.connectNative(NATIVE_HOSTS[kind]);
  port.onMessage.addListener(answer => {
    const reply = waiting.get(answer.id);
    if (reply) { waiting.delete(answer.id); reply(answer); }
  });
  port.onDisconnect.addListener(() => {
    const why = chrome.runtime.lastError?.message ?? `the ${kind} host went away`;
    for (const [id, reply] of waiting) if (reply.kind === kind) { waiting.delete(id); reply({ error: why }); }
    delete ports[kind];
  });
  ports[kind] = port;
  return port;
}
chrome.runtime.onMessage.addListener((message, sender, reply) => {
  if (message?.oxi !== 'xll' && message?.oxi !== 'com') return false;
  const kind = message.oxi;
  const id = nextId++;
  const done = answer => { const { id: _, ...rest } = answer; reply(rest); };
  done.kind = kind;
  waiting.set(id, done);
  try {
    nativePort(kind).postMessage({ ...message.request, id });
  } catch (error) {
    waiting.delete(id);
    reply({ error: String(error) });
  }
  return true;
});
