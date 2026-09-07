// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

/** Draw the core engine's upright vertical text columns in page coordinates.
 * Returns false for horizontal text, leaving the caller's canvas untouched.
 */
export function renderVerticalText(ctx, elem) {
  if (!elem.is_vertical) return false;

  const fs = elem.font_size || 11;
  const x = elem.x + Math.max(0, (elem.width - fs) / 2);
  const advance = fs + (elem.character_spacing || 0);
  ctx.save();
  ctx.font = `${elem.italic ? 'italic ' : ''}${elem.bold ? 'bold ' : ''}${fs}px ${elem.font_family || 'Calibri, sans-serif'}`;
  ctx.textAlign = 'left';
  ctx.textBaseline = 'alphabetic';
  ctx.letterSpacing = '0px';

  if (elem.highlight && elem.highlight !== 'none') {
    ctx.fillStyle = elem.highlight;
    ctx.fillRect(elem.x, elem.y, elem.width, elem.height);
  }
  ctx.fillStyle = elem.color || '#000000';
  let baseline = elem.y + fs * 0.85;
  // Iterate Unicode code points; surrogate pairs occupy one vertical cell.
  for (const ch of elem.text || '') {
    if (!/\s/u.test(ch)) ctx.fillText(ch, x, baseline);
    baseline += advance;
  }

  ctx.strokeStyle = elem.color || '#000000';
  ctx.lineWidth = fs * 0.05;
  const line = (lineX) => {
    ctx.moveTo(lineX, elem.y);
    ctx.lineTo(lineX, elem.y + elem.height);
  };
  if (elem.underline || elem.strikethrough) {
    ctx.beginPath();
    if (elem.underline) {
      line(x + fs);
      if (elem.underline_style === 'double') line(x + fs + Math.max(1, fs * 0.08));
    }
    if (elem.strikethrough) line(x + fs * 0.5);
    ctx.stroke();
  }
  ctx.restore();
  return true;
}
