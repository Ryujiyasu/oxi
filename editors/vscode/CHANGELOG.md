# Changelog

## 0.9.4

In a narrow pane the slide list and the slide shrink to fit, instead of the
slide being cut off. Word layout fixes: no extra space above the second line
of a paragraph, text in vertical headers and footers is drawn, letter spacing
on a document grid is spread per letter in Latin text, and private-use
characters take one em in Japanese fonts.

## 0.9.3

Table cells in Word documents can be edited and saved; the cell's other runs
keep their formatting. Only text that the save can put back exactly where it
was typed takes editing — it shows on hover — and the rest (headers, text
boxes) is read-only, where before an edit could land in the wrong paragraph.
A Word document with a tab inside a run no longer loses text on save. In a
narrow pane the ribbon scrolls sideways instead of squeezing, and a wide page
starts at the left edge instead of being cut off.

## 0.9.2

The spreadsheet selection and the slide list use Oxi's own colours instead of
the Excel and PowerPoint brand colours.

## 0.9.1

Edits are saved. Text edited in Word paragraphs, Excel cells and PowerPoint
text boxes is written back to the file with Ctrl+S or Save As; the tab shows
unsaved edits, and Revert File reloads the file from disk. The extension is now
named an editor.

## 0.9.0

First release. Opens `.docx`, `.xlsx` and `.pptx` in a VS Code tab, rendered by
Oxi 0.9.0, the same WebAssembly build that powers the browser demo.

Read-only: the document is rendered, not saved back.
