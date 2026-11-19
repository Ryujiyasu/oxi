// Oxi Office editor — opens .docx / .xlsx / .pptx in a webview that runs the
// same WebAssembly engine as the browser editor, and saves the edits back.
//
// The document bytes are read by VS Code and handed to the webview. On save,
// the page builds the edited file the way it does for its download button
// (the changed text written into the original package for .docx and .pptx,
// the changed cells for .xlsx), hands the bytes back, and VS Code writes them.
// Nothing leaves the machine.

const vscode = require('vscode');

const VIEW_TYPE = 'oxi.officeViewer';

/** A document is its bytes on disk; the page holds the edits until a save. */
class OxiDocument {
  constructor(uri, bytes) {
    this.uri = uri;
    this.bytes = bytes;
    this.panel = null;
    this.pending = new Map();
    this.nextId = 1;
  }

  /** Show `bytes` as the document in its webview (on open, after a save or a revert). */
  show(bytes) {
    this.panel?.webview.postMessage({
      type: 'oxi:open',
      name: basename(this.uri),
      data: Buffer.from(bytes).toString('base64'),
    });
  }

  /** The edited file from the page, or null when nothing has changed. */
  edited() {
    if (!this.panel) return Promise.resolve(null);
    const id = this.nextId++;
    return new Promise((resolve, reject) => {
      this.pending.set(id, { resolve, reject });
      this.panel.webview.postMessage({ type: 'oxi:save', id });
    });
  }

  answer(msg) {
    const waiting = this.pending.get(msg.id);
    if (!waiting) return;
    this.pending.delete(msg.id);
    if (msg.error) waiting.reject(new Error(msg.error));
    else waiting.resolve(msg.unchanged ? null : { bytes: Buffer.from(msg.data, 'base64'), warning: msg.warning });
  }

  dispose() {
    for (const { reject } of this.pending.values()) reject(new Error('the editor was closed'));
    this.pending.clear();
  }
}

class OxiEditorProvider {
  constructor(context) {
    this.context = context;
    this.changes = new vscode.EventEmitter();
    // VS Code marks the tab as changed when this fires, and asks to save on close.
    this.onDidChangeCustomDocument = this.changes.event;
  }

  async openCustomDocument(uri, openContext) {
    // After a crash or a restart with unsaved changes, VS Code hands back the backup.
    const from = openContext.backupId ? vscode.Uri.parse(openContext.backupId) : uri;
    const bytes = await vscode.workspace.fs.readFile(from);
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
    panel.onDidDispose(() => { if (document.panel === panel) document.panel = null; });
    panel.webview.html = withWebviewHead(html, panel.webview, mediaRoot);

    panel.webview.onDidReceiveMessage((msg) => {
      if (!msg) return;
      // The page tells us when its WebAssembly module is up; only then are the
      // bytes worth sending.
      if (msg.type === 'oxi:ready') document.show(document.bytes);
      if (msg.type === 'oxi:edit') this.changes.fire({ document });
      if (msg.type === 'oxi:saved') document.answer(msg);
      if (msg.type === 'oxi:error') vscode.window.showErrorMessage(`Oxi: ${msg.message}`);
    });
  }

  async saveCustomDocument(document) {
    const edited = await document.edited();
    if (!edited) return;
    await vscode.workspace.fs.writeFile(document.uri, edited.bytes);
    document.bytes = edited.bytes;
    if (edited.warning) vscode.window.showWarningMessage(`Oxi: ${edited.warning}`);
    // The saved file is the new starting point for the next edits.
    document.show(document.bytes);
  }

  async saveCustomDocumentAs(document, destination) {
    const edited = await document.edited();
    await vscode.workspace.fs.writeFile(destination, edited ? edited.bytes : document.bytes);
  }

  async revertCustomDocument(document) {
    document.bytes = await vscode.workspace.fs.readFile(document.uri);
    document.show(document.bytes);
  }

  async backupCustomDocument(document, context) {
    const edited = await document.edited().catch(() => null);
    await vscode.workspace.fs.writeFile(context.destination, edited ? edited.bytes : document.bytes);
    return {
      id: context.destination.toString(),
      delete: async () => {
        try { await vscode.workspace.fs.delete(context.destination); } catch { /* already gone */ }
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

  const head =
    `<base href="${base}/">\n` +
    `<meta http-equiv="Content-Security-Policy" content="${csp}">\n`;

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
