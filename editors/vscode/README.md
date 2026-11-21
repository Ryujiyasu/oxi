# Oxi — Word, Excel and PowerPoint in VS Code, next to your AI agent

Ask Claude Code, or any other coding agent, to write or fix a `.docx`, `.xlsx`
or `.pptx`, and watch the result appear in the tab beside it. Oxi draws the
pages the way Microsoft Office does, and redraws them each time the agent saves
the file. You stay on the same page, slide or sheet.

You can also open, edit and save these files yourself, with a Word-style
ribbon. Everything runs locally in the editor's webview. Oxi uploads nothing
and sends no telemetry.

## Watch your agent work on Office files

1. Open the document in VS Code. It opens in an Oxi tab.
2. Ask your agent for the change. For example:
   - *"Add a row for consulting work to the price table in quote.docx and
     update the total."*
   - *"Turn notes.md into a 6-slide deck, deck.pptx."*
   - *"In budget.xlsx, add a sheet that sums each department by month."*
3. The agent writes the file with whatever it likes to use, such as
   python-docx, openpyxl, python-pptx or a document skill. The Oxi tab notices
   the new file on disk and redraws it in place.

If you have edits in the tab that are not saved yet, Oxi asks before it
reloads. It never silently throws your work away. When you save in the tab,
Oxi writes back only what you edited and keeps the rest of the file as it was.
Your changes and the agent's can therefore take turns on the same file.

### Let the agent check its own work

An agent can't see your VS Code tab. It can see an image of the page. The
[`oxidocs`](https://pypi.org/project/oxidocs/) Python package bundles the same
engine as a command:

```bash
pip install oxidocs
oxidocs docx-to-png --quiet --page 1 quote.docx page1.png
oxidocs pptx-to-png --quiet --page 3 deck.pptx slide3.png
oxidocs xlsx-to-png --quiet budget.xlsx sheet.png
oxidocs docx-to-pdf --quiet quote.docx quote.pdf
```

To make this a habit, add a note to your project's `CLAUDE.md` or
`AGENTS.md`:

```markdown
After changing a .docx, .pptx or .xlsx, render the pages you changed with
`oxidocs docx-to-png --quiet --page N file.docx out.png` (or pptx-to-png /
xlsx-to-png) and look at the image before saying the change is done.
```

Many AI document generators check their output by converting it with
LibreOffice. Oxi is scored against Microsoft Office itself, so what the agent
sees is close to what the person who opens the file in Word will see.

## How faithful is it?

Oxi is scored against Microsoft Office on blind sets of documents. The sets are
frozen before measurement, and their documents are never used as targets for
fixes.

| Blind set | Oxi mean SSIM vs Office | page/slide count matches |
|---|---|---|
| Word, English, 50 documents (2026-11-20) | 0.926 | 49 / 50 |
| Word, Japanese, 50 documents (2026-11-20) | 0.884 | 43 / 50 |
| PowerPoint, 48 decks (2026-08-20) | 0.953 | 48 / 48 |

The [Oxi README](https://gitlab.com/Ryujiyasu/oxi#readme) describes the method
and the corpora, and lists the other engines measured on the same sets.

Japanese typography is a first-class target. Oxi supports JIS X 4051 kinsoku
line breaking, the document grid, ruby, vertical writing, warichu and emphasis
marks.

## What it does today

- **Follows the file on disk.** When another program writes the file (an
  agent, a script, `git checkout`), the tab redraws it and keeps your place. If
  the tab has unsaved edits, Oxi asks first.
- **Renders** `.docx`, `.xlsx` and `.pptx` in a VS Code tab.
- **Edits text in place:** Word paragraphs, Excel cells and PowerPoint text
  boxes.
- **Word documents:**
  - Enter and Backspace split and join paragraphs.
  - Selecting works as in Word: a double click selects a word, a triple click a
    paragraph, and Shift+click or Shift+arrow extends the selection. Deleting a
    selection works even when it runs into a table.
  - The ribbon shows the font, size and effects of the selected text.
  - The ribbon's formatting, tables, links, footnotes, comments, headers,
    footers, fields, and page size and margins are saved into the file.
  - Undo and redo work with Ctrl+Z / Ctrl+Y.
- **Fonts:** Word text is drawn with your machine's fonts. If you don't have a
  face (Calibri on a Mac, for instance), Oxi draws it with an open look-alike
  that it ships, and the layout keeps Word's own measured widths.
- **Saving:** save with **Ctrl+S** or *Save As…*. While there are unsaved
  edits, the tab shows the file as modified, and *Revert File* reloads it from
  disk. Saving changes only what you edited. Oxi rewrites the edited runs and
  cells inside the original package and keeps everything else as it was.
- **Spreadsheets and decks:** the formatting buttons change only the view. They
  are not saved yet.
- **First open:** large documents take a moment the first time. The engine and
  its font metric tables are a single WebAssembly module.

To open a file with VS Code's own editor instead, use
**Oxi: Reopen this file with VS Code's default editor**, or right-click the file
and choose *Open With…*

## Building from source

```bash
cd editors/vscode
npm run sync-web     # copies the WebAssembly build out of web/ into media/
npx @vscode/vsce package
```

`media/` is generated and never committed. It is a copy of the repository's
`web/` build, so the extension and the browser demo can't drift apart.

## License

The extension is MIT OR Apache-2.0, matching Oxi's binding layer. The engine
itself is MPL-2.0.
