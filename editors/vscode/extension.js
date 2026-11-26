// Oxi Office editor — opens .docx / .xlsx / .pptx in a webview that runs the
// same WebAssembly engine as the browser demo, and saves edits back to the
// file. The document bytes are read and written by VS Code; nothing leaves the
// machine.

const vscode = require('vscode');
const { InstalledFonts } = require('./fonts');

// The id from the first (view-only) release is kept so that editor associations
// users already set keep working.
const VIEW_TYPE = 'oxi.officeViewer';

/** The machine's fonts, indexed once for every Word editor that asks. */
const installedFonts = new InstalledFonts();

/**
 * Output › Oxi lists each face a Word document asked for and where it came
 * from, once per face: the file on this machine, or the open stand-in Oxi
 * draws when the machine does not have it.
 */
let output = null;
const reported = new Set();
function reportFace(family, bold, italic, file) {
  const face = `${family}${bold ? ' Bold' : ''}${italic ? ' Italic' : ''}`;
  if (!output || reported.has(face)) return;
  reported.add(face);
  output.appendLine(file ? `font  ${face}  ←  ${file}` : `font  ${face}  ←  not on this machine; drawn with Oxi's stand-in`);
}

/**
 * The page each format opens in: Word documents in the Word editor, which
 * edits through the document model (typing, Enter, Backspace across
 * paragraphs, tables, notes, comments); workbooks in the spreadsheet editor
 * the desktop app uses (formatting, merges, number formats, rows and columns,
 * undo, macros); decks in the suite page.
 */
function pageFor(uri) {
  if (/\.docx$/i.test(uri.path)) return 'docs.html';
  if (/\.xlsx$/i.test(uri.path)) return 'xlsx-demo.html';
  return 'index.html';
}

/** A save that waits longer than this is reported as failed. */
const SAVE_TIMEOUT_MS = 120000;

/**
 * The file as the webview page last received it. The page applies its edits
 * to these bytes, so a save with no edits writes them back unchanged.
 */
class OxiDocument {
  constructor(uri, bytes) {
    this.uri = uri;
    this.bytes = bytes;
    this.panel = null;
    this.pending = new Map();
    this.nextRequest = 1;
    // Edits made since the file was opened, saved or reverted.
    this.dirty = false;
    this.disposables = [];
  }

  dispose() {
    for (const d of this.disposables) d.dispose();
    this.disposables = [];
    for (const { reject } of this.pending.values()) {
      reject(new Error('the editor was closed'));
    }
    this.pending.clear();
  }

  /**
   * Send the current bytes to the page, which lays the document out again.
   * `reload` keeps the page where it was (scroll, slide, sheet): the same
   * file, changed on disk.
   */
  show(reload = false) {
    if (!this.panel) return;
    this.panel.webview.postMessage({
      type: 'oxi:open',
      reload,
      name: basename(this.uri),
      // Webview messages are JSON, so the bytes travel as base64 rather than
      // as a Uint8Array turned inside out into an object of indices.
      data: Buffer.from(this.bytes).toString('base64'),
    });
  }

  /** Ask the page for the file with its edits applied. */
  requestBytes() {
    if (!this.panel) return Promise.resolve(this.bytes);
    const id = this.nextRequest++;
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        this.pending.delete(id);
        reject(new Error('the editor did not produce the file in time'));
      }, SAVE_TIMEOUT_MS);
      this.pending.set(id, {
        resolve: (value) => { clearTimeout(timer); resolve(value); },
        reject: (err) => { clearTimeout(timer); reject(err); },
      });
      this.panel.webview.postMessage({ type: 'oxi:save', id });
    });
  }

  /** The page's answer to requestBytes. */
  settle(msg) {
    const waiter = this.pending.get(msg.id);
    if (!waiter) return;
    this.pending.delete(msg.id);
    if (msg.error) {
      waiter.reject(new Error(msg.error));
    } else if (msg.data) {
      waiter.resolve(new Uint8Array(Buffer.from(msg.data, 'base64')));
    } else {
      // No edits: the file is the one the page was given.
      waiter.resolve(this.bytes);
    }
  }
}

class OxiEditorProvider {
  constructor(context) {
    this.context = context;
    this.changes = new vscode.EventEmitter();
    this.onDidChangeCustomDocument = this.changes.event;
  }

  async openCustomDocument(uri, openContext) {
    const source = openContext.backupId ? vscode.Uri.parse(openContext.backupId) : uri;
    const bytes = await vscode.workspace.fs.readFile(source);
    const document = new OxiDocument(uri, bytes);
    this.watch(document);
    return document;
  }

  /**
   * Follow the file on disk: when another program writes it (Claude Code, a
   * script, git), the tab shows the new file. Edits not yet saved are never
   * thrown away without asking.
   */
  watch(document) {
    if (document.uri.scheme !== 'file') return;
    const folder = vscode.Uri.joinPath(document.uri, '..');
    const watcher = vscode.workspace.createFileSystemWatcher(
      new vscode.RelativePattern(folder, basename(document.uri))
    );
    let timer = null;
    // A writer may write the file in pieces, or replace it by renaming a new
    // file over it; act once it has been quiet for a moment.
    const changed = () => {
      clearTimeout(timer);
      timer = setTimeout(() => this.reloadFromDisk(document), 400);
    };
    watcher.onDidChange(changed);
    watcher.onDidCreate(changed);
    document.disposables.push(watcher, { dispose: () => clearTimeout(timer) });
  }

  async reloadFromDisk(document) {
    let bytes;
    try {
      bytes = await vscode.workspace.fs.readFile(document.uri);
    } catch {
      return; // Deleted or unreadable for now; a later change brings it back.
    }
    // Our own save, or a write of the same content.
    if (Buffer.from(bytes).equals(Buffer.from(document.bytes))) return;
    if (!document.dirty) {
      document.bytes = bytes;
      document.show(true);
      return;
    }
    if (document.asking) return;
    document.asking = true;
    const ja = vscode.env.language.startsWith('ja');
    const reload = ja ? '読み直す' : 'Reload';
    const keep = ja ? 'このまま編集を続ける' : 'Keep my edits';
    const choice = await vscode.window.showWarningMessage(
      ja
        ? `${basename(document.uri)} がほかのプログラムで書き換えられました。読み直すと、保存していない編集は失われます。`
        : `${basename(document.uri)} was changed on disk. Reloading discards your unsaved edits.`,
      reload, keep
    );
    document.asking = false;
    if (choice === reload && document.panel) {
      document.panel.reveal();
      await vscode.commands.executeCommand('workbench.action.files.revert');
    }
  }

  async resolveCustomEditor(document, panel) {
    const mediaRoot = vscode.Uri.joinPath(this.context.extensionUri, 'media');
    panel.webview.options = {
      enableScripts: true,
      localResourceRoots: [mediaRoot],
    };

    const page = pageFor(document.uri);
    let html;
    try {
      const raw = await vscode.workspace.fs.readFile(
        vscode.Uri.joinPath(mediaRoot, page)
      );
      html = new TextDecoder('utf-8').decode(raw);
    } catch (err) {
      panel.webview.html = missingMediaHtml();
      return;
    }

    document.panel = panel;
    panel.onDidDispose(() => {
      if (document.panel === panel) document.panel = null;
    });
    panel.webview.html = withWebviewHead(html, panel.webview, mediaRoot, page);

    panel.webview.onDidReceiveMessage((msg) => {
      if (!msg) return;
      switch (msg.type) {
        // The page tells us when its WebAssembly module is up; only then are
        // the bytes worth sending.
        case 'oxi:ready':
          document.show();
          break;
        case 'oxi:edit':
          document.dirty = true;
          if (msg.step === undefined) {
            // The suite page has no undo stack: the file is just modified.
            this.changes.fire({ document });
          } else if (msg.step) {
            // The Word page began an undo step: VS Code keeps the step, and its
            // undo and redo are carried out by the page.
            this.changes.fire({
              document,
              label: 'Edit',
              undo: () => document.panel?.webview.postMessage({ type: 'oxi:undo' }),
              redo: () => document.panel?.webview.postMessage({ type: 'oxi:redo' }),
            });
          }
          break;
        case 'oxi:requestUndo':
          vscode.commands.executeCommand('undo');
          break;
        case 'oxi:requestRedo':
          vscode.commands.executeCommand('redo');
          break;
        case 'oxi:saved':
          document.settle(msg);
          break;
        case 'oxi:requestSave':
          vscode.commands.executeCommand('workbench.action.files.save');
          break;
        case 'oxi:error':
          vscode.window.showErrorMessage(`Oxi: ${msg.message}`);
          break;
        // The Word editor draws with the font programs themselves, which a
        // webview cannot read; the bytes are sent to this page only.
        case 'oxi:font': {
          const family = String(msg.family || '');
          installedFonts.resolve(family, !!msg.bold, !!msg.italic)
            .catch(() => null)
            .then((font) => {
              reportFace(family, !!msg.bold, !!msg.italic, font && font.file);
              return panel.webview.postMessage({
                type: 'oxi:font',
                id: msg.id,
                data: font ? Buffer.from(font.bytes).toString('base64') : null,
                index: font ? font.index : 0,
              });
            });
          break;
        }
      }
    });
  }

  async saveCustomDocument(document) {
    const bytes = await this.writeTo(document, document.uri);
    // What is on disk now: the watcher sees this write as no change.
    document.bytes = bytes;
    document.dirty = false;
  }

  async saveCustomDocumentAs(document, destination) {
    await this.writeTo(document, destination);
  }

  async writeTo(document, target) {
    const bytes = await document.requestBytes();
    await vscode.workspace.fs.writeFile(target, bytes);
    return bytes;
  }

  async revertCustomDocument(document) {
    document.bytes = await vscode.workspace.fs.readFile(document.uri);
    document.dirty = false;
    document.show(true);
  }

  async backupCustomDocument(document, context) {
    let bytes;
    try {
      bytes = await document.requestBytes();
    } catch {
      bytes = document.bytes;
    }
    await vscode.workspace.fs.writeFile(context.destination, bytes);
    return {
      id: context.destination.toString(),
      delete: async () => {
        try {
          await vscode.workspace.fs.delete(context.destination);
        } catch {
          // Already gone.
        }
      },
    };
  }
}

function basename(uri) {
  const parts = uri.path.split('/');
  return parts[parts.length - 1] || 'document';
}

/**
 * A webview resolves relative URLs against its own origin, so the page needs a
 * <base> pointing at the extension's media directory, and a policy that lets
 * WebAssembly compile.
 */
function withWebviewHead(html, webview, mediaRoot, page) {
  const base = webview.asWebviewUri(mediaRoot).toString();
  const csp = [
    "default-src 'none'",
    `img-src ${webview.cspSource} data: blob:`,
    `font-src ${webview.cspSource} data:`,
    `style-src ${webview.cspSource} 'unsafe-inline'`,
    `script-src ${webview.cspSource} 'unsafe-inline' 'unsafe-eval' 'wasm-unsafe-eval' blob:`,
    `worker-src ${webview.cspSource} blob:`,
    `connect-src ${webview.cspSource} data: blob:`,
  ].join('; ');

  // VS Code owns the file, so the browser page's ways of picking, leaving and
  // downloading a document are hidden from the first paint, before the engine
  // has loaded. Saving is VS Code's own (Ctrl+S, the title bar's save button).
  const hidden = (page === 'docs.html'
    ? [
      // New, open, rename, and the exports a webview cannot download or print.
      '#btnNew', '#btnOpen', '#btnPdf',
      '#fmNew', '#fmOpen', '#fmCopy', '#fmExport', '#fmPrint', '#fmRename',
    ]
    : page === 'xlsx-demo.html'
      // The site's name and lede (a link away from the editor), the sample,
      // new and open buttons, and the drop sign.
      ? ['header > .name', 'header > .what', '#reset', '#fresh', 'header label.file', '#drop']
      : ['#landingHeader', '#dropZone', '#sampleButtons', '.side-door', '#toolbar']
  ).join(', ');

  const head =
    `<base href="${base}/">\n` +
    `<meta http-equiv="Content-Security-Policy" content="${csp}">\n` +
    `<style>${hidden} { display: none !important; }</style>\n`;

  if (html.includes('<head>')) {
    return html.replace('<head>', `<head>\n${head}`);
  }
  return head + html;
}

function missingMediaHtml() {
  return `<!DOCTYPE html><html><body style="font-family: sans-serif; padding: 2rem">
    <h2>Oxi editor is missing its engine</h2>
    <p>The extension was packaged without <code>media/</code>. Run
    <code>npm run sync-web</code> in <code>editors/vscode</code> to copy the
    WebAssembly build in, then reload.</p>
  </body></html>`;
}

function activate(context) {
  output = vscode.window.createOutputChannel('Oxi');
  context.subscriptions.push(output);
  context.subscriptions.push(
    vscode.window.registerCustomEditorProvider(
      VIEW_TYPE,
      new OxiEditorProvider(context),
      { supportsMultipleEditorsPerDocument: false, webviewOptions: { retainContextWhenHidden: true } }
    )
  );

  context.subscriptions.push(
    vscode.commands.registerCommand('oxi.openWithDefaultEditor', async () => {
      const uri = vscode.window.activeTextEditor?.document.uri
        ?? vscode.window.tabGroups.activeTabGroup.activeTab?.input?.uri;
      if (uri) {
        await vscode.commands.executeCommand('vscode.openWith', uri, 'default');
      }
    })
  );
}

function deactivate() {}

module.exports = { activate, deactivate };
