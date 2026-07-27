// Browser-oracle harness: drives @betteroffice/docx-react's DocxEditor
// headlessly, the way a host app ships it (readOnly viewing mode).
// Playwright calls window.oracleInit(docUrl) then window.oraclePage(i)
// and reads back PNG data URLs of each page canvas. Page bitmaps are
// backing-store resolution, so the driver sets deviceScaleFactor = dpi/96.
import React from 'react';
import { createRoot } from 'react-dom/client';
import { DocxEditor } from '@betteroffice/docx-react';
import '@betteroffice/docx-react/styles.css';

let root = null;
let editorRef = null;

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

// The editor stamps each page canvas with data-page-index (0-based).
function pageCanvas(i) {
  return document.querySelector(`canvas[data-page-index="${i}"]`);
}

// The canvas holds glyphs over a transparent background; the page fill
// lives on the .canvas-page wrapper's CSS. Composite before exporting so
// the PNG matches what the user sees.
function exportCanvas(canvas) {
  const target = document.createElement('canvas');
  target.width = canvas.width;
  target.height = canvas.height;
  const ctx = target.getContext('2d');
  let fill = '#ffffff';
  const wrap = canvas.closest('.canvas-page');
  if (wrap) {
    const bg = getComputedStyle(wrap).backgroundColor;
    if (bg && bg !== 'transparent' && bg !== 'rgba(0, 0, 0, 0)') fill = bg;
  }
  ctx.fillStyle = fill;
  ctx.fillRect(0, 0, target.width, target.height);
  ctx.drawImage(canvas, 0, 0);
  return target.toDataURL('image/png');
}

window.oracleInit = async (docUrl) => {
  const buf = await (await fetch(docUrl)).arrayBuffer();
  if (root) { try { root.unmount(); } catch (e) { /* ignore */ } }
  const host = document.getElementById('host');
  host.innerHTML = '';
  window.__boError = null;
  editorRef = React.createRef();
  root = createRoot(host);
  root.render(React.createElement(DocxEditor, {
    ref: editorRef,
    documentBuffer: buf,
    readOnly: true,
    showZoomControl: false,
    showOutlineButton: false,
    initialZoom: 1,
    colorMode: 'light',
    onError: (e) => { window.__boError = String(e && e.message || e); },
  }));

  // Layout runs async in wasm; wait until the page count is stable.
  const t0 = Date.now();
  let last = -1;
  let stableSince = 0;
  while (Date.now() - t0 < 120000) {
    if (window.__boError) throw new Error('editor error: ' + window.__boError);
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

// Render page i (0-based) and return a PNG data URL of its canvas backing
// store. Scrolls the page into view first in case the view virtualizes.
window.oraclePage = async (i) => {
  if (!editorRef || !editorRef.current) throw new Error('no document');
  const total = editorRef.current.getTotalPages();
  if (i >= total) throw new Error('page out of range');
  try { editorRef.current.scrollToPage(i + 1); } catch (e) { /* ignore */ }
  // Give the painter a beat, then wait for the canvas to exist and settle.
  const t0 = Date.now();
  while (Date.now() - t0 < 30000) {
    const canvas = pageCanvas(i);
    if (canvas && canvas.width > 0) {
      await sleep(150);
      return exportCanvas(canvas);
    }
    await sleep(200);
  }
  throw new Error('page canvas not found: ' + i + '/' + total);
};

// Debug helper for the smoke test: dump canvas + container structure.
window.oracleDom = () => {
  const canvases = [...document.querySelectorAll('canvas')].map((c) => ({
    w: c.width, h: c.height, cw: c.clientWidth, ch: c.clientHeight,
    cls: c.className, parentCls: c.parentElement ? c.parentElement.className : '',
    grandCls: c.parentElement && c.parentElement.parentElement
      ? c.parentElement.parentElement.className : '',
    dataset: { ...c.dataset },
    parentDataset: c.parentElement ? { ...c.parentElement.dataset } : {},
  }));
  return {
    canvases,
    total: editorRef && editorRef.current ? editorRef.current.getTotalPages() : -1,
    err: window.__boError,
  };
};

window.oracleReady = true;
