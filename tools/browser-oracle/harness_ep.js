// Browser-oracle harness: drives @eigenpal/docx-editor-react's DocxEditor
// headlessly, the way a host app ships it (readOnly viewing mode).
// Unlike BetterOffice (canvas pages), eigenpal renders DOM pages
// (.layout-page, ProseMirror + its own layout engine), so the PNG capture
// happens on the Playwright side via element screenshots at
// deviceScaleFactor = dpi/96. This file only mounts the editor, waits for
// the page count to settle, and exposes per-page scroll/locate helpers.
import React from 'react';
import { createRoot } from 'react-dom/client';
import { DocxEditor } from '@eigenpal/docx-editor-react';
import '@eigenpal/docx-editor-react/styles.css';

let root = null;
let editorRef = null;

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

window.oracleInit = async (docUrl) => {
  const buf = await (await fetch(docUrl)).arrayBuffer();
  if (root) { try { root.unmount(); } catch (e) { /* ignore */ } }
  const host = document.getElementById('host');
  host.innerHTML = '';
  window.__epError = null;
  editorRef = React.createRef();
  root = createRoot(host);
  root.render(React.createElement(DocxEditor, {
    ref: editorRef,
    documentBuffer: buf,
    readOnly: true,
    showToolbar: false,
    showZoomControl: false,
    showOutlineButton: false,
    initialZoom: 1,
    colorMode: 'light',
    onError: (e) => { window.__epError = String(e && e.message || e); },
  }));

  // Layout runs async; wait until the page count is stable.
  const t0 = Date.now();
  let last = -1;
  let stableSince = 0;
  while (Date.now() - t0 < 120000) {
    if (window.__epError) throw new Error('editor error: ' + window.__epError);
    let n = 0;
    try { n = editorRef.current ? editorRef.current.getTotalPages() : 0; } catch (e) { n = 0; }
    if (n > 0 && n === last) {
      if (!stableSince) stableSince = Date.now();
      if (Date.now() - stableSince >= 3000) return n;
    } else {
      stableSince = 0;
      last = n;
    }
    await sleep(250);
  }
  throw new Error('layout did not settle; pages=' + last);
};

// Scroll page i (0-based) into view so a (possibly virtualized) page exists
// in the DOM, and confirm the .layout-page element is present and laid out.
// The actual PNG capture is a Playwright element screenshot on the caller's
// side: page.locator('.layout-page').nth(i).screenshot().
window.oracleGoto = async (i) => {
  if (!editorRef || !editorRef.current) throw new Error('no document');
  const total = editorRef.current.getTotalPages();
  if (i >= total) throw new Error('page out of range');
  try { editorRef.current.scrollToPage(i + 1); } catch (e) { /* ignore */ }
  const t0 = Date.now();
  while (Date.now() - t0 < 30000) {
    const pages = document.querySelectorAll('.layout-page');
    const el = pages[i];
    if (el) {
      const r = el.getBoundingClientRect();
      if (r.width > 0 && r.height > 0) {
        await sleep(150); // let the painter settle
        return { count: pages.length, w: r.width, h: r.height };
      }
    }
    await sleep(200);
  }
  throw new Error('page element not found: ' + i + '/' + total);
};

// Debug helper: dump page-element structure for the smoke test.
window.oracleDom = () => {
  const pages = [...document.querySelectorAll('.layout-page')].map((p) => {
    const r = p.getBoundingClientRect();
    return { w: r.width, h: r.height, cls: p.className, dataset: { ...p.dataset } };
  });
  return {
    pages,
    total: editorRef && editorRef.current ? editorRef.current.getTotalPages() : -1,
    err: window.__epError,
  };
};

window.oracleReady = true;
