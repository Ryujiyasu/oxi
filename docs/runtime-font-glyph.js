// SPDX-License-Identifier: MIT OR Apache-2.0

// The supplied outline stays in memory. Its provider must validate the actual
// family, face, style and selected glyph against the registered font program.
export function drawRuntimeFontGlyph(ctx, element, baseline, painting) {
  if (!element.font_glyph) return false;
  if (!painting) throw new Error('The selected glyph needs its font program');
  const selected = element.font_glyph;
  if (painting.index !== selected.index) throw new Error('Selected glyph mismatch');
  if (!Array.isArray(painting.bounds_em) || painting.bounds_em.length !== 4 ||
      !Array.isArray(selected.bounds_em) || selected.bounds_em.length !== 4 ||
      painting.bounds_em.some((v, i) => !Number.isFinite(v) ||
        !Number.isFinite(selected.bounds_em[i]) || Math.abs(v - selected.bounds_em[i]) > 1e-6)) {
    throw new Error('Registered font geometry differs from the layout');
  }
  const size = element.font_size;
  if (!Number.isFinite(size) || size <= 0 || !Number.isFinite(element.x) || !Number.isFinite(baseline)) {
    throw new Error('Invalid glyph origin or size');
  }
  if (!Array.isArray(painting.commands)) throw new Error('Missing outline commands');
  const arities = { Move: 2, Line: 2, Quad: 4, Curve: 6, Close: 0 };
  for (const command of painting.commands) {
    const arity = arities[command.op];
    const points = command.points ?? [];
    if (arity == null || !Array.isArray(points) || points.length !== arity ||
        points.some(v => !Number.isFinite(v))) throw new Error('Invalid outline command');
  }
  ctx.save();
  try {
    ctx.translate(element.x, baseline);
    ctx.scale(size, -size);
    if (element.color) ctx.fillStyle = element.color;
    ctx.beginPath();
    for (const command of painting.commands) {
      const points = command.points ?? [];
      switch (command.op) {
        case 'Move': ctx.moveTo(...points); break;
        case 'Line': ctx.lineTo(...points); break;
        case 'Quad': ctx.quadraticCurveTo(...points); break;
        case 'Curve': ctx.bezierCurveTo(...points); break;
        case 'Close': ctx.closePath(); break;
      }
    }
    ctx.fill();
  } finally {
    ctx.restore();
  }
  return true;
}

// `fit` (RuntimeFontLoader.standInFit): true when the element is painted in a
// stand-in face, whose advances are not the ones the layout measured; text
// that would run past the element's width is then narrowed to it, so it does
// not overprint the next element. An array of [text, x, scale] pieces sets
// each piece at x from the element's left edge, scaled horizontally:
// ＭＳ Ｐゴシック's Latin, kana and 「（・」 each in the cell the layout
// measured for it.
export function drawRegisteredText(ctx, element, baseline, getOutline, fit = false) {
  if (element.font_glyph) {
    const glyph = element.font_glyph;
    const painting = getOutline(element.font_family, !!element.bold, !!element.italic,
      glyph.index, new Float32Array(glyph.bounds_em));
    drawRuntimeFontGlyph(ctx, element, baseline, painting);
    return;
  }
  const text = element.text || '';
  if (Array.isArray(fit)) {
    for (const [piece, x, scale] of fit) {
      ctx.save();
      try {
        ctx.translate(element.x + x, baseline);
        ctx.scale(scale, 1);
        ctx.fillText(piece, 0, 0);
      } finally {
        ctx.restore();
      }
    }
    return;
  }
  if (fit && element.width > 0) {
    const drawn = ctx.measureText(text.replace(/\s+$/u, '')).width;
    if (drawn > element.width * 1.01) {
      ctx.save();
      try {
        ctx.translate(element.x, baseline);
        ctx.scale(element.width / drawn, 1);
        ctx.fillText(text, 0, 0);
      } finally {
        ctx.restore();
      }
      return;
    }
  }
  ctx.fillText(text, element.x, baseline);
}
