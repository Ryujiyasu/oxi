# Oxi — Word / Excel / PowerPoint editor for VS Code

Open, edit and save `.docx`, `.xlsx` and `.pptx` files directly in VS Code. The rendering is
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
- Edit the text of Word documents and slides, and the cells of workbooks, in
  place; **Ctrl+S** saves back to the file, **Save As** writes a copy, and
  **Revert File** goes back to what is on disk. A tab with unsaved edits is
  marked like any other, and the edits survive a restart (hot exit)
- **Only what you change is written.** A save puts the changed text into the
  original package and leaves the rest of the file as it was: only the part
  holding the edited text (the document body, that slide, that sheet) is
  rewritten, and every other part — styles, images, charts, parts Oxi does
  not draw — is kept byte for byte, so the file still opens the same in Word,
  Excel and PowerPoint
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
