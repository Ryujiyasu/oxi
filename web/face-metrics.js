// What this browser can measure, for the layout engine to use.
//
// The engine's compiled tables hold the faces the build machine could measure.
// A browser can measure any face it is able to DRAW, which is a much larger set
// and the only one that helps somebody else's deck -- so the page measures the
// characters a shape needs and hands them to the engine, whose rules (the
// 1/8pt master unit, the per-run measure, the indent geometry) then run on real
// numbers instead of declining.
//
// ★The trap this module exists for: a browser asked for a font it does not
// have does NOT fail. It silently substitutes, and `measureText` then returns
// the advances of a DIFFERENT face -- so measuring without checking would feed
// the engine confident, wrong numbers, which is worse than declining. Every
// family is therefore proved present before it is measured.

const probeCanvas = document.createElement('canvas');
const probe = probeCanvas.getContext('2d');

const PROBE_TEXT = 'mmmwwwiiillMWQ0123!@#';
const GENERICS = ['monospace', 'serif', 'sans-serif'];

const presentCache = new Map();
const faceCache = new Map();

function widthIn(font) {
  probe.font = font;
  return probe.measureText(PROBE_TEXT).width;
}

/**
 * Whether this browser actually has `family`.
 *
 * Asked by comparing the family against three generic fallbacks: a family the
 * browser does not have falls back to the generic and measures IDENTICALLY to
 * it. Matching all three means absent.
 *
 * A family that genuinely matches every generic reads as absent too, and that
 * is the safe direction -- the engine falls back to its tables rather than
 * trusting a substitute.
 */
export function familyPresent(family) {
  if (!family) return false;
  if (presentCache.has(family)) return presentCache.get(family);
  const quoted = JSON.stringify(String(family));
  let present = false;
  for (const generic of GENERICS) {
    const base = widthIn(`100px ${generic}`);
    const asked = widthIn(`100px ${quoted}, ${generic}`);
    if (Math.abs(base - asked) > 0.5) { present = true; break; }
  }
  presentCache.set(family, present);
  return present;
}

// A family no browser can have, used to see what the DEFAULT font measures --
// which is what a missing glyph falls through to.
const GHOST = 'Zzyzx No Such Family 42';

/**
 * Whether a per-glyph check can tell this family's own glyphs from the
 * default font's.
 *
 * ★A glyph the face LACKS is not an error either: the browser substitutes for
 * that ONE character and `measureText` returns the substitute's advance. The
 * engine's whole reason for `has_all_glyphs` is to refuse those, so measuring
 * without checking would quietly undo it.
 *
 * The check is "does this character measure exactly like the default font
 * does" -- which is only meaningful when the family itself measures
 * differently from the default. For a family that does not (it IS the default,
 * or is metrically identical to it) no per-glyph check is possible, and the
 * face is taken as-is.
 */
function glyphCheckable(family) {
  const quoted = JSON.stringify(String(family));
  return Math.abs(widthIn(`100px ${quoted}`) - widthIn(`100px ${JSON.stringify(GHOST)}`)) > 0.5;
}

// The ink of one character, hashed, for asking WHICH font drew it.
//
// ★Comparing advances is not enough: Arial's 'W' and Times New Roman's are
// both 0.944 em, so an advance test calls a letter Arial plainly has
// "substituted". The shapes are nothing alike, so the ink answers cleanly.
const inkCanvas = document.createElement('canvas');
inkCanvas.width = 64;
inkCanvas.height = 64;
const inkCtx = inkCanvas.getContext('2d', { willReadFrequently: true });

function inkHash(font, ch) {
  inkCtx.clearRect(0, 0, 64, 64);
  inkCtx.font = font;
  inkCtx.textBaseline = 'alphabetic';
  inkCtx.fillStyle = '#000';
  inkCtx.fillText(ch, 6, 52);
  const d = inkCtx.getImageData(0, 0, 64, 64).data;
  let h = 2166136261, on = 0;
  for (let i = 3; i < d.length; i += 4) {
    if (d[i]) on++;
    h ^= d[i];
    h = Math.imul(h, 16777619);
  }
  return { hash: h >>> 0, blank: on === 0 };
}

/**
 * The advance of each character of `chars` in EM units.
 *
 * Null for the whole answer when the browser does not have the family, and
 * null in one slot for a character the family itself does not have.
 */
export function measureFace(family, bold, italic, chars) {
  if (!familyPresent(family)) return null;
  const key = `${family}|${bold ? 1 : 0}|${italic ? 1 : 0}`;
  let face = faceCache.get(key);
  if (!face) {
    // A large size so the returned float carries enough digits: the engine
    // quantises to 1/8pt, and at 1000px one unit of rounding is far below it.
    const style = `${italic ? 'italic ' : ''}${bold ? 'bold ' : ''}`;
    face = {
      font: `${style}1000px ${JSON.stringify(String(family))}`,
      ghost: `${style}1000px ${JSON.stringify(GHOST)}`,
      // The ink test wants a size that fits the drawing canvas.
      inkFont: `${style}40px ${JSON.stringify(String(family))}`,
      inkGhost: `${style}40px ${JSON.stringify(GHOST)}`,
      checkable: glyphCheckable(family),
      em: new Map(),
    };
    faceCache.set(key, face);
  }
  const out = [];
  for (const ch of chars) {
    let em = face.em.get(ch);
    if (em === undefined) {
      // Measured one character at a time, so the answer is the isolated
      // advance the break law sums -- not a kerned pair width.
      probe.font = face.font;
      em = probe.measureText(ch).width / 1000;
      if (face.checkable) {
        const mine = inkHash(face.inkFont, ch);
        const theirs = inkHash(face.inkGhost, ch);
        // ★What this does NOT catch: a character the face lacks whose
        // substitute is a THIRD font -- one the default font would not have
        // used, because the default font has the character itself. Then the
        // two renderings differ and the substitute's advance is accepted. The
        // case that matters in practice (CJK in a Latin face, where the
        // default font has no such glyph either) does land on the same
        // substitute and is caught.
        if (mine.hash === theirs.hash) {
          // Identical ink means the default font drew it -- unless the
          // character has no ink at all (a space), where the advance is the
          // only thing that distinguishes the two.
          probe.font = face.ghost;
          const ghostEm = probe.measureText(ch).width / 1000;
          if (!mine.blank || Math.abs(ghostEm - em) < 1e-9) em = null;
        }
      }
      face.em.set(ch, em);
    }
    // A refused character stays null in the answer rather than sinking the
    // whole face: the caller leaves it out of what it hands over, and the
    // engine then declines exactly the runs that need it -- which is what
    // `has_all_glyphs` does with a face that lacks a glyph.
    out.push(em);
  }
  return out;
}

/**
 * Everything the page can measure for a set of runs, in the shape
 * `layout_slide_shape` takes.
 *
 * `runs` is an iterable of `{text, font_family, bold, italic}`; families left
 * blank take `fallbackFamily`.
 */
export function collectAdvances(runs, fallbackFamily) {
  const need = new Map();   // key -> {family, bold, italic, chars:Set}
  for (const r of runs) {
    const family = r.font_family || fallbackFamily;
    if (!family || !r.text) continue;
    const bold = !!r.bold, italic = !!r.italic;
    const key = `${family}|${bold ? 1 : 0}|${italic ? 1 : 0}`;
    let e = need.get(key);
    if (!e) { e = { family, bold, italic, chars: new Set() }; need.set(key, e); }
    for (const ch of r.text) e.chars.add(ch);
  }
  const out = [];
  for (const e of need.values()) {
    // Indexed by CODE POINT, not by code unit: `measureFace` walks the string
    // with for..of, so a character outside the basic plane occupies one slot
    // in the answer and two in the string.
    const cps = [...e.chars];
    const measured = measureFace(e.family, e.bold, e.italic, cps.join(''));
    if (!measured) continue;
    // Only the characters this face actually has are handed over. The rest are
    // left unanswered on purpose, so the engine declines the run instead of
    // laying it out on a substitute's advances.
    const chars = [], em = [];
    for (let i = 0; i < cps.length; i++) {
      if (measured[i] === null) continue;
      chars.push(cps[i]);
      em.push(measured[i]);
    }
    if (chars.length) {
      out.push({ family: e.family, bold: e.bold, italic: e.italic,
                 chars: chars.join(''), em });
    }
  }
  return out;
}

/** Which of these families this browser can measure, for an honest report. */
export function measurableFamilies(families) {
  const out = new Set();
  for (const f of families) if (familyPresent(f)) out.add(f);
  return out;
}
