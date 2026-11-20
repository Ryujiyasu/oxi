// The machine's own fonts, for the Word editor page. The page asks for a face
// by family, bold and italic; this answers with the bytes of the installed
// font program whose name table says it is that face, or null.
//
// It is the VS Code counterpart of the desktop app's font provider
// (crates/oxi-desktop/font-provider): the same name-table matching (family
// names 1 and 16, case and spacing folded; bold and italic from OS/2
// fsSelection), so a face resolves to the same file in both. Only the name
// and OS/2 tables are read to build the index; a program's bytes are read
// when the page asks for that face, and they go to the page in this window
// and nowhere else.

const fs = require('fs');
const os = require('os');
const path = require('path');

const FONT_FILE = /\.(ttf|otf|ttc)$/i;

/** Font folders in priority order. Earlier folders win a family/style. */
function fontRoots(platform = process.platform, env = process.env, home = os.homedir()) {
  if (platform === 'win32') {
    const roots = [];
    if (env.LOCALAPPDATA) {
      // Office's cloud fonts, then the per-user install, as the desktop app.
      roots.push(path.join(env.LOCALAPPDATA, 'Microsoft', 'FontCache', '4', 'CloudFonts'));
      roots.push(path.join(env.LOCALAPPDATA, 'Microsoft', 'Windows', 'Fonts'));
    }
    roots.push(path.join(env.SystemRoot || env.windir || 'C:\\Windows', 'Fonts'));
    return roots;
  }
  if (platform === 'darwin') {
    return [
      path.join(home, 'Library', 'Fonts'),
      '/Library/Fonts',
      '/System/Library/Fonts',
      '/System/Library/Fonts/Supplemental',
      // Fonts Microsoft Office for Mac installs for its own use.
      '/Applications/Microsoft Word.app/Contents/Resources/DFonts',
    ];
  }
  return [
    path.join(home, '.local', 'share', 'fonts'),
    path.join(home, '.fonts'),
    '/usr/local/share/fonts',
    '/usr/share/fonts',
  ];
}

function key(name) {
  return name.split(/\s+/).filter(Boolean).join(' ').toLowerCase();
}

async function listFontFiles(root, depth = 4) {
  let entries;
  try {
    entries = await fs.promises.readdir(root, { withFileTypes: true });
  } catch {
    return [];
  }
  const files = [];
  const dirs = [];
  for (const entry of entries) {
    const full = path.join(root, entry.name);
    if (entry.isDirectory()) dirs.push(full);
    else if (FONT_FILE.test(entry.name)) files.push(full);
  }
  files.sort();
  dirs.sort();
  if (depth > 0) {
    for (const dir of dirs) files.push(...await listFontFiles(dir, depth - 1));
  }
  return files;
}

async function readAt(handle, offset, length) {
  const buffer = Buffer.alloc(length);
  const { bytesRead } = await handle.read(buffer, 0, length, offset);
  return buffer.subarray(0, bytesRead);
}

/** A name record's text, for the encodings ttf-parser decodes. */
function decodeName(platform, encoding, bytes) {
  if (platform === 0 || (platform === 3 && (encoding === 0 || encoding === 1 || encoding === 10))) {
    let s = '';
    for (let i = 0; i + 1 < bytes.length; i += 2) s += String.fromCharCode(bytes.readUInt16BE(i));
    return s;
  }
  if (platform === 1 && encoding === 0) return bytes.toString('latin1');
  return null;
}

/** Family names (IDs 1 and 16), bold and italic of the face at `offset`. */
async function readFace(handle, offset) {
  const header = await readAt(handle, offset, 12);
  if (header.length < 12) return null;
  const numTables = header.readUInt16BE(4);
  const dir = await readAt(handle, offset + 12, numTables * 16);
  const tables = {};
  for (let i = 0; i + 16 <= dir.length; i += 16) {
    tables[dir.toString('latin1', i, i + 4)] = { offset: dir.readUInt32BE(i + 8), length: dir.readUInt32BE(i + 12) };
  }
  if (!tables.name) return null;
  const name = await readAt(handle, tables.name.offset, tables.name.length);
  if (name.length < 6) return null;
  const count = name.readUInt16BE(2);
  const strings = name.readUInt16BE(4);
  // Each family name with whether the face gives it as its own (ID 1) or
  // only as the typographic family it belongs to (ID 16, as "Yu Gothic" is
  // for Yu Gothic Light).
  const families = new Map();
  for (let i = 0; i < count; i++) {
    const at = 6 + i * 12;
    if (at + 12 > name.length) break;
    const id = name.readUInt16BE(at + 6);
    if (id !== 1 && id !== 16) continue;
    const length = name.readUInt16BE(at + 8);
    const start = strings + name.readUInt16BE(at + 10);
    if (start + length > name.length) continue;
    const text = decodeName(name.readUInt16BE(at), name.readUInt16BE(at + 2), name.subarray(start, start + length));
    if (text && !(families.get(key(text)) === 1)) families.set(key(text), id);
  }
  let bold = false;
  let italic = false;
  let weight = 400;
  if (tables['OS/2'] && tables['OS/2'].length >= 64) {
    weight = (await readAt(handle, tables['OS/2'].offset + 4, 2)).readUInt16BE(0) || 400;
    const fsSelection = (await readAt(handle, tables['OS/2'].offset + 62, 2)).readUInt16BE(0);
    bold = (fsSelection & 0x20) !== 0;
    // Oblique (bit 9) is a style of its own, not italic, as in ttf-parser.
    italic = (fsSelection & 0x200) === 0 && (fsSelection & 0x1) !== 0;
  }
  return { families, bold, italic, weight };
}

async function indexFile(file, faces) {
  let handle;
  try {
    handle = await fs.promises.open(file, 'r');
    const head = await readAt(handle, 0, 12);
    if (head.length < 12) return;
    let offsets = [0];
    if (head.toString('latin1', 0, 4) === 'ttcf') {
      const n = head.readUInt32BE(8);
      const table = await readAt(handle, 12, n * 4);
      offsets = [];
      for (let i = 0; i + 4 <= table.length; i += 4) offsets.push(table.readUInt32BE(i));
    }
    for (const [index, offset] of offsets.entries()) {
      const face = await readFace(handle, offset);
      if (!face) continue;
      for (const [family, nameId] of face.families) {
        const k = `${family}\u0000${face.bold}\u0000${face.italic}`;
        // Several faces can claim one family/style: Yu Gothic Light and
        // Yu Gothic Regular both say "Yu Gothic", not bold, not italic. The
        // face whose own name it is wins, then the weight nearest the
        // regular (or bold) one; otherwise the earlier folder and file.
        const rank = (nameId === 1 ? 0 : 10000) + Math.abs(face.weight - (face.bold ? 700 : 400));
        const held = faces.get(k);
        if (!held || rank < held.rank) faces.set(k, { file, index, rank });
      }
    }
  } catch {
    // An unreadable or malformed file is skipped, never guessed at.
  } finally {
    if (handle) await handle.close();
  }
}

class InstalledFonts {
  constructor(roots = fontRoots()) {
    this.roots = roots;
    this.index = null;
  }

  /** Map of "family\0bold\0italic" → {file, index}, built once on first use. */
  faces() {
    if (!this.index) {
      this.index = (async () => {
        const faces = new Map();
        for (const root of this.roots) {
          for (const file of await listFontFiles(root)) await indexFile(file, faces);
        }
        return faces;
      })();
    }
    return this.index;
  }

  /** Where this face is: {file, index} (its member in a collection), or null. */
  async locate(family, bold, italic) {
    const faces = await this.faces();
    const found = faces.get(`${key(family)}\u0000${!!bold}\u0000${!!italic}`);
    return found ? { file: found.file, index: found.index } : null;
  }

  /** The program holding this face, {bytes, index, file}, or null. */
  async resolve(family, bold, italic) {
    const found = await this.locate(family, bold, italic);
    if (!found) return null;
    try {
      return { bytes: await fs.promises.readFile(found.file), index: found.index, file: found.file };
    } catch {
      return null;
    }
  }
}

module.exports = { InstalledFonts, fontRoots };
