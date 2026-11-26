# Changelog

## 0.9.11

Japanese text in a font the machine does not have (drawn in Oxi's stand-in
faces) now keeps its letters in place. In ＭＳ ゴシック / ＭＳ 明朝, Latin
letters sit inside their half-width cells instead of touching. In 游ゴシック,
メイリオ, 游明朝, ＭＳ Ｐゴシック and ＭＳ Ｐ明朝, Latin text is proportional, no
longer typewriter-like with a gap after each word («Wi ndows»). In ＭＳ Ｐゴシック
and ＭＳ Ｐ明朝, kana and marks such as 「（」「・」 are as narrow as those fonts
set them, no longer running into the next character. Word documents also lay
out Japanese closer to Word: punctuation squeezed in table cells, lines
justified across every gap, and vertical text. PowerPoint line spacing is
kept in whole percent, as PowerPoint does.

Text in a Latin font Oxidocs has no measurements for (Inter, for one) is now
drawn in Carlito, the face whose widths it was laid out with, so words keep
their spaces and closing marks no longer land on the letter before them.
Excel sheet tabs stay on one line: a long sheet name is cut with an ellipsis
and shown in full as the tooltip.

The extension is now called Oxidocs.

## 0.9.10

Word documents: Up and Down move the caret a line at a time, keeping its
column, and Shift+Up and Shift+Down select line by line (before, only Left
and Right moved the caret). Pages on a character grid break lines closer to
Word. Excel: compact pivot tables draw their expand buttons.

## 0.9.9

PowerPoint decks: tables are drawn in the editor (before, a slide's table
was missing there while exports showed it), and oval and rounded callouts
are drawn with their shape instead of as rectangles.

## 0.9.8

The tab follows the file on disk. When another program writes it (Claude
Code or another AI agent, a script, `git checkout`), the document is drawn
again in place, on the same page, slide or sheet. If the tab has edits that
are not saved, Oxi asks before it reloads. The README now starts with this
way of working, including how an agent can render pages to check its own
work.

## 0.9.7

Word documents: a selection that runs from body text into a table, or from
one cell to another, is deleted with Backspace or Delete. The text it covers
goes and the table keeps its rows and cells (before, nothing was deleted).

The font and size boxes and the Bold, Italic, Underline and Strikethrough
buttons follow the selected text, as in Word: a button is down when all of
it has that effect, and a box is blank when the text mixes fonts or sizes.
A size that is not on the list (26 pt, say) is shown as it is.

## 0.9.6

Word documents: the Layout tab's Size (A4, Letter) and Margins (Normal,
Narrow) buttons work. They change the page setup of the section the cursor
is in, as Word's galleries do, and Undo takes them back. In Japanese,
Normal is Word's Japanese default (35 mm top, 30 mm elsewhere).

Spreadsheets and decks take the colours people know these formats by (green
for workbooks, orange-red for decks) on the same neutral grays as the Word
editor, and the PowerPoint ribbon's tabs, groups and button tips are in
Japanese when the editor is. An empty cell that still has formatting no
longer reads "Empty".

## 0.9.5

Word documents open in Oxi's Word editor, which edits through the document
model: Enter splits a paragraph, Backspace at the start of one joins it to the
one before, and bold, italic, font, size,
colour, alignment, tables, links, footnotes, comments, headers, footers and
fields are saved into the file. Text is drawn with the machine's own fonts,
read by the extension and handed to the page in the same window; a face the
machine does not have is drawn in the open face Oxi ships for it (Carlito for
Calibri, Liberation for Arial and Times New Roman, OxiMincho and OxiGothic
for Japanese faces), and the layout keeps Word's measured widths either way.
Output › Oxi lists which file each face came from, or that it was drawn with
a stand-in.

Backspace at the start of a numbered paragraph takes the number off first and
leaves the text where it stood, as Word does. Undo and redo work in Word
documents: Ctrl+Z, Ctrl+Y, the Edit menu and the buttons in the title bar.
Typing is undone a run at a time, and undoing back to what was saved shows
the file as unmodified again.

Selecting text works as in Word: a double click selects a word (Japanese is
cut into words too), a triple click the paragraph, and shift-click or
Shift+Left/Right extends the selection; Backspace or Delete removes it, and a
paragraph selected whole goes with its mark. A click puts the caret at the
nearer edge of the character instead of always after it.

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
