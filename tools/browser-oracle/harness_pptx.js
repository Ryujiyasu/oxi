// Browser-oracle harness: drives @silurus/ooxml's PptxPresentation headlessly.
// Playwright calls window.oracleInit(deckUrl) then window.oraclePage(i, dpi)
// and reads back PNG data URLs — the pptx sibling of harness.js (docx).
import { PptxPresentation } from '@silurus/ooxml/pptx';

let pres = null;

window.oracleInit = async (deckUrl) => {
  const buf = await (await fetch(deckUrl)).arrayBuffer();
  pres = await PptxPresentation.load(buf);
  return pres.slideCount;
};

// Render slide i (0-based) at the given DPI (slide EMU width -> inches * dpi)
// and return a PNG data URL. A fresh canvas per call keeps state independent.
window.oraclePage = async (i, dpi) => {
  const px = Math.round((pres.slideWidth / 914400) * dpi);
  if (pres.renderSlideToBitmap) {
    const bmp = await pres.renderSlideToBitmap(i, { width: px, dpr: 1 });
    const target = document.createElement('canvas');
    target.width = bmp.width;
    target.height = bmp.height;
    target.getContext('2d').drawImage(bmp, 0, 0);
    bmp.close();
    return target.toDataURL('image/png');
  }
  const target = document.createElement('canvas');
  await pres.renderSlide(target, i, { width: px, dpr: 1 });
  return target.toDataURL('image/png');
};

window.oracleDebug = () => ({
  proto: Object.getOwnPropertyNames(Object.getPrototypeOf(pres)),
  slides: pres ? pres.slideCount : -1,
  emu: pres ? [pres.slideWidth, pres.slideHeight] : null,
});

window.oracleReady = true;
