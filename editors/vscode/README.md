# Oxi — Word / Excel / PowerPoint viewer for VS Code

Open `.docx`, `.xlsx` and `.pptx` files directly in VS Code. The rendering is
done by [Oxi](https://gitlab.com/Ryujiyasu/oxi), a Rust + WebAssembly Office
engine whose page layout is measured, document by document, against Microsoft
Office's own output.

Everything runs locally inside the editor's webview. No server, no upload, no
telemetry — the file is read by VS Code and handed to the WebAssembly module in
the same window.

## How faithful is it?

Oxi is scored against Microsoft Office on blind document sets that are frozen
before measurement and never used as fix targets:

| Blind set | Oxi mean SSIM vs Office | page/slide count matches |
|---|---|---|
| English, 50 documents | 0.875 | 48 / 50 |
| Japanese, 50 documents | 0.842 | 43 / 50 |
| PowerPoint, 48 decks | 0.953 | 48 / 48 |

Japanese typography is a first-class target: JIS X 4051 kinsoku line breaking,
document grid, ruby, vertical writing, warichu, emphasis marks.

## What it does today

- Renders `.docx`, `.xlsx` and `.pptx` in a VS Code tab
- **Read-only.** Edits made in the embedded UI are not written back to the file
- Large documents take a moment on first open — the engine and its font metric
  tables are a single WebAssembly module

To open a file with VS Code's own editor instead, use
**Oxi: Reopen this file with VS Code's default editor**, or right-click the file
and choose *Open With…*

## Building from source

```bash
cd editors/vscode
npm run sync-web     # copies the WebAssembly build out of web/ into media/
npx @vscode/vsce package
```

`media/` is generated, never committed: it is a copy of the repository's `web/`
build, so the extension and the browser demo can never drift apart.

## License

MIT OR Apache-2.0, matching Oxi's binding layer. The engine itself is MPL-2.0.
