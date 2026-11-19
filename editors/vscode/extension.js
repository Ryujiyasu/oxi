// Oxi Office editor — opens .docx / .xlsx / .pptx in a webview that runs the
// same WebAssembly engine as the browser demo, and saves edits back to the
// file. The document bytes are read and written by VS Code; nothing leaves the
// machine.

const vscode = require('vscode');

// The id from the first (view-only) release is kept so that editor associations
// users already set keep working.
const VIEW_TYPE = 'oxi.officeViewer';

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
  }

  dispose() {
    for (const { reject } of this.pending.values()) {
      reject(new Error('the editor was closed'));
    }
    this.pending.clear();
  }

  /** Send the current bytes to the page, which lays the document out again. */
  show() {
    if (!this.panel) return;
    this.panel.webview.postMessage({
      type: 'oxi:open',
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
    return new OxiDocument(uri, bytes);
  }

  async resolveCustomEditor(document, panel) {
    const mediaRoot = vscode.Uri.joinPath(this.context.extensionUri, 'media');
    panel.webview.options = {
      enableScripts: true,
      localResourceRoots: [mediaRoot],
    };

    let html;
    try {
      const raw = await vscode.workspace.fs.readFile(
        vscode.Uri.joinPath(mediaRoot, 'index.html')
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
    panel.webview.html = withWebviewHead(html, panel.webview, mediaRoot);

    panel.webview.onDidReceiveMessage((msg) => {
      if (!msg) return;
      switch (msg.type) {
        // The page tells us when its WebAssembly module is up; only then are
        // the bytes worth sending.
        case 'oxi:ready':
          document.show();
          break;
        case 'oxi:edit':
          this.changes.fire({ document });
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
      }
    });
  }

  async saveCustomDocument(document) {
    await this.writeTo(document, document.uri);
  }

  async saveCustomDocumentAs(document, destination) {
    await this.writeTo(document, destination);
  }

  async writeTo(document, target) {
    const bytes = await document.requestBytes();
    await vscode.workspace.fs.writeFile(target, bytes);
  }

  async revertCustomDocument(document) {
    document.bytes = await vscode.workspace.fs.readFile(document.uri);
    document.show();
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
function withWebviewHead(html, webview, mediaRoot) {
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
  const hidden = [
    '#landingHeader', '#dropZone', '#sampleButtons', '.side-door', '#toolbar',
  ].join(', ');

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
