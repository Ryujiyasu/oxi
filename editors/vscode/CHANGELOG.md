# Changelog

## 0.9.0

Saves. The viewer is now an editor: edits to the text of Word documents and
slides, and to the cells of workbooks, are written back to the file with
Ctrl+S (also Save As, Revert File and hot exit). Only the changed text or
cells are written; the rest of the package is kept as it was.

## 0.8.0

First release. Opens `.docx`, `.xlsx` and `.pptx` in a VS Code tab, rendered by
Oxi 0.8.0 — the same WebAssembly build that powers the browser demo.

Read-only: the document is rendered, not saved back.
