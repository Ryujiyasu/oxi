// Oxi Office viewer — opens .docx / .xlsx / .pptx in a webview that runs the
// same WebAssembly engine as the browser demo. The document bytes are read by
// VS Code and handed to the webview; nothing is written back (this editor is
// read-only) and nothing leaves the machine.

const vscode = require('vscode');

const VIEW_TYPE = 'oxi.officeViewer';

/** A document that is just its bytes — the engine does the rest. */
class OxiDocument {
  constructor(uri, bytes) {
    this.uri = uri;
    this.bytes = bytes;
  }
  dispose() {}
}

class OxiEditorProvider {
  constructor(context) {
    this.context = context;
  }

  async openCustomDocument(uri) {
    const bytes = await vscode.workspace.fs.readFile(uri);
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

    panel.webview.html = withWebviewHead(html, panel.webview, mediaRoot);

    // The page tells us when its WebAssembly module is up; only then are the
    // bytes worth sending.
    panel.webview.onDidReceiveMessage((msg) => {
      if (msg && msg.type === 'oxi:ready') {
        // Webview messages are JSON, so the bytes travel as base64 rather
        // than as a Uint8Array turned inside out into an object of indices.
        panel.webview.postMessage({
          type: 'oxi:open',
          name: basename(document.uri),
          data: Buffer.from(document.bytes).toString('base64'),
        });
      }
      if (msg && msg.type === 'oxi:error') {
        vscode.window.showErrorMessage(`Oxi: ${msg.message}`);
      }
    });
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
    <h2>Oxi viewer is missing its engine</h2>
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
