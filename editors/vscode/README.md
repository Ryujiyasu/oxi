# Oxi — .docx / .xlsx / .pptx editor for VS Code

Open, edit and save `.docx`, `.xlsx` and `.pptx` files directly in VS Code. The rendering is
done by [Oxi](https://gitlab.com/Ryujiyasu/oxi), a Rust + WebAssembly Office
engine whose page layout is measured, document by document, against Microsoft
Office's own output.

Everything runs locally inside the editor's webview. No server, no upload, no
telemetry — the file is read and written by VS Code and handed to the
WebAssembly module in the same window.

## How faithful is it?

Oxi is scored against Microsoft Office on blind document sets that are frozen
before measurement and never used as fix targets:

| Blind set | Oxi mean SSIM vs Office | page/slide count matches |
|---|---|---|
| Word, English, 50 documents (2026-11-20) | 0.926 | 49 / 50 |
| Word, Japanese, 50 documents (2026-11-20) | 0.884 | 43 / 50 |
| PowerPoint, 48 decks (2026-08-20) | 0.953 | 48 / 48 |

The method, the corpora and the other engines measured on the same sets are in
the [Oxi README](https://gitlab.com/Ryujiyasu/oxi#readme).

Japanese typography is a first-class target: JIS X 4051 kinsoku line breaking,
document grid, ruby, vertical writing, warichu, emphasis marks.

## What it does today

- Renders `.docx`, `.xlsx` and `.pptx` in a VS Code tab
- Edit text in place: Word paragraphs, Excel cells, PowerPoint text boxes
- Word documents: Enter and Backspace split and join paragraphs; formatting,
  tables, links, footnotes, comments, headers, footers and fields from the
  ribbon are saved into the file
- Word text is drawn with your machine's fonts; a face you do not have (Calibri
  on a Mac, for instance) is drawn in an open look-alike Oxi ships, while the
  layout keeps Word's own measured widths
- Save with **Ctrl+S** (or *Save As…*); the tab shows the file as modified while
  there are unsaved edits, and *Revert File* reloads it from disk
- Saving changes only what you edited: Oxi rewrites the edited runs and cells
  inside the original package and keeps everything else as it was
- In spreadsheets and decks, the formatting buttons change the view only; they
  are not saved yet
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
