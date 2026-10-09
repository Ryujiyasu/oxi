// SPDX-License-Identifier: MIT OR Apache-2.0
// Font bytes and CSS faces stay in memory. No outline/font data files or
// family translation lists are produced. Program metadata resolves aliases.

export class LocalFontAccessRequired extends Error {}

export class RuntimeFontLoader {
  constructor(api, host = window) {
    this.api = api;
    this.host = host;
    this.pending = new Map();
    this.cssFaces = new Map();
    this.cssFaceRevisions = new Map();
    this.localFonts = null;
  }

  key(family, bold, italic) { return JSON.stringify([family, !!bold, !!italic]); }

  async enableLocalFonts() {
    if (typeof this.host.queryLocalFonts !== 'function') {
      throw new LocalFontAccessRequired('This browser needs font files to reproduce the original glyphs.');
    }
    // Called only by the explicit font-access button. Enumeration follows the
    // browser's own permission UI; opening a document never grants permission.
    this.localFonts = await this.host.queryLocalFonts();
    this.pending.clear();
  }

  async localFontList() {
    if (this.localFonts) return this.localFonts;
    if (typeof this.host.queryLocalFonts !== 'function') {
      throw new LocalFontAccessRequired('The document needs its fonts. Choose font files to continue.');
    }
    let permission;
    try { permission = await this.host.navigator.permissions.query({ name: 'local-fonts' }); }
    catch { throw new LocalFontAccessRequired('Allow this document to use fonts on this device.'); }
    if (permission.state !== 'granted') {
      throw new LocalFontAccessRequired('Allow this document to use fonts on this device.');
    }
    this.localFonts = await this.host.queryLocalFonts();
    return this.localFonts;
  }

  async addCssFace(family, bold, italic) {
    const key = this.key(family, bold, italic);
    const revision = this.api.font_program_revision();
    if (this.cssFaces.has(key) && this.cssFaceRevisions.get(key) === revision) return;
    const oldFace = this.cssFaces.get(key);
    if (oldFace) this.host.document.fonts.delete(oldFace);
    // The selected member is lifted from TTC into a complete SFNT container.
    // CFF remains wrapped as OTF here; PDF's bare-CFF stream is not a CSS font.
    const bytes = this.api.get_registered_font_sfnt(family, !!bold, !!italic);
    const face = new this.host.FontFace(family, bytes, {
      weight: bold ? '700' : '400', style: italic ? 'italic' : 'normal',
    });
    await face.load();
    this.host.document.fonts.add(face);
    this.cssFaces.set(key, face);
    this.cssFaceRevisions.set(key, revision);
  }

  async loadFace(family, bold, italic) {
    if (this.api.has_font_program(family, !!bold, !!italic)) {
      await this.addCssFace(family, bold, italic);
      return;
    }
    const invoke = this.host.__TAURI__?.core?.invoke;
    if (invoke) {
      const bytes = await invoke('load_font_program', { family, bold: !!bold, italic: !!italic });
      this.api.register_font_program_family(family, !!bold, !!italic, new Uint8Array(bytes));
      await this.addCssFace(family, bold, italic);
      return;
    }
    const list = await this.localFontList();
    // The browser's family strings are hints only: localized name-table aliases
    // and TTC members are matched by the real font parser, never a name map.
    const canonical = name => name.split(/\s+/).filter(Boolean).join(' ').toLocaleLowerCase();
    const ordered = [...list].sort((a, b) =>
      Number(canonical(b.family) === canonical(family)) - Number(canonical(a.family) === canonical(family)));
    for (const font of ordered) {
      const bytes = new Uint8Array(await (await font.blob()).arrayBuffer());
      if (!this.api.try_register_font_program_family(family, !!bold, !!italic, bytes)) continue;
      await this.addCssFace(family, bold, italic);
      return;
    }
    throw new Error(`Font unavailable: ${family}${bold ? ' bold' : ''}${italic ? ' italic' : ''}`);
  }

  async ensure(layout) {
    const requests = new Map();
    for (const page of layout.pages || []) {
      for (const element of page.elements || []) {
        if (element.kind !== 'text' || !element.text || !element.font_family) continue;
        const key = this.key(element.font_family, element.bold, element.italic);
        requests.set(key, [element.font_family, !!element.bold, !!element.italic]);
      }
    }
    // Sequential requests bound transient font-program memory and native IPC.
    for (const [key, args] of requests) {
      if (!this.pending.has(key)) {
        this.pending.set(key, this.loadFace(...args).catch(error => {
          const failure = error instanceof Error ? error : new Error(String(error));
          failure.fontAccessError = true;
          throw failure;
        }));
      }
      const pending = this.pending.get(key);
      try { await pending; }
      finally {
        if (this.pending.get(key) === pending) this.pending.delete(key);
      }
      await this.addCssFace(...args);
    }
  }

  async resolveLayout(createLayout, initial, isCurrent = () => true) {
    let layout = initial || createLayout();
    while (isCurrent()) {
      const revision = this.api.font_program_revision();
      try { await this.ensure(layout); }
      catch (error) { error.fontLayout = layout; throw error; }
      if (!isCurrent() || revision === this.api.font_program_revision()) return layout;
      layout = createLayout();
    }
    return layout;
  }

  async importFiles(files, layout) {
    // Manual input works where the browser's system-font API is unavailable.
    // Only requested family/style combinations are registered and rendered.
    for (const file of files) {
      const bytes = new Uint8Array(await file.arrayBuffer());
      const requested = new Map();
      for (const page of layout.pages || []) for (const element of page.elements || []) {
        if (element.kind === 'text' && element.text && element.font_family) {
          requested.set(this.key(element.font_family, element.bold, element.italic),
            [element.font_family, !!element.bold, !!element.italic]);
        }
      }
      for (const [key, args] of requested) {
        if (this.api.try_register_font_program_family(...args, bytes)) {
          const oldFace = this.cssFaces.get(key);
          if (oldFace) this.host.document.fonts.delete(oldFace);
          this.cssFaces.delete(key);
          this.pending.delete(key);
        }
      }
    }
    this.pending.clear();
    await this.ensure(layout);
  }
}

export function showFontAccess(container, loader, layout, retry, error) {
  container.textContent = error.message || String(error);
  const device = document.createElement('button');
  device.textContent = 'Use device fonts';
  device.onclick = async () => {
    try { await loader.enableLocalFonts(); await loader.ensure(layout); await retry(); }
    catch (next) { showFontAccess(container, loader, layout, retry, next); }
  };
  const file = document.createElement('input');
  file.type = 'file'; file.accept = '.ttf,.otf,.ttc'; file.multiple = true;
  file.onchange = async () => {
    try { await loader.importFiles(file.files, layout); await retry(); }
    catch (next) { showFontAccess(container, loader, layout, retry, next); }
  };
  container.append(device, file);
}
