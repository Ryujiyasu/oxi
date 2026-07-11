// Browser-oracle harness: drives @silurus/ooxml's DocxViewer headlessly.
// Playwright calls window.oracleInit(docUrl) then window.oraclePage(i, dpi)
// and reads back PNG data URLs. The viewer renders on a hidden canvas.
import { DocxDocument } from '@silurus/ooxml/docx';

let doc = null;

window.oracleInit = async (docUrl) => {
  const buf = await (await fetch(docUrl)).arrayBuffer();
  doc = await DocxDocument.load(buf);
  return doc.pageCount;
};

// Render page i at the given DPI (Word-point page width * dpi/72) and return
// a PNG data URL. A fresh canvas per call keeps state independent.
window.oraclePage = async (i, dpi) => {
  const sz = doc.pageSize(i); // { widthPt, heightPt }
  const px = Math.round((sz.widthPt / 72) * dpi);
  if (doc.renderPageToBitmap) {
    const bmp = await doc.renderPageToBitmap(i, { width: px, dpr: 1 });
    const target = document.createElement('canvas');
    target.width = bmp.width;
    target.height = bmp.height;
    target.getContext('2d').drawImage(bmp, 0, 0);
    return target.toDataURL('image/png');
  }
  const target = document.createElement('canvas');
  await doc.renderPage(target, i, { width: px, dpr: 1 });
  return target.toDataURL('image/png');
};

window.oracleProbe = async () => {
  const r = {};
  try { r.pageSize = doc.pageSize(0); } catch (e) { r.pageSize = 'ERR ' + e.message; }
  try {
    const bmp = await doc.renderPageToBitmap(0, { width: 800, dpr: 1 });
    r.bmp = { w: bmp.width, h: bmp.height };
  } catch (e) { r.bmp = 'ERR ' + e.message; }
  try {
    const bmp2 = await doc.renderPageToBitmap(0, {});
    r.bmpDefault = { w: bmp2.width, h: bmp2.height };
  } catch (e) { r.bmpDefault = 'ERR ' + e.message; }
  return r;
};

window.oracleDebug = () => ({
  proto: Object.getOwnPropertyNames(Object.getPrototypeOf(doc)),
  hasBitmap: typeof doc.renderPageToBitmap,
  hasRender: typeof doc.renderPage,
});

window.oracleReady = true;
