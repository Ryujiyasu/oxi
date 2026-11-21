# Oxi Office Add-ins (browser extension)

Office add-ins are web pages that load Microsoft's `office.js` from
`appsforoffice.microsoft.com` and talk to Excel through it. The Oxi spreadsheet
editor opens an add-in page in its side pane; with this extension installed,
that page's request for `office.js` is answered with Oxi's own
[`oxi-office.js`](../../web/oxi-office.js), which speaks the same Excel
JavaScript API to the workbook open in the editor. The add-in itself is not
changed.

Only tabs with the Oxi editor open are redirected. Excel on the web and every
other site keep loading Microsoft's `office.js`.

## Install (unpacked)

```
node integrations/office-js-extension/build.mjs
```

Then open `chrome://extensions` (or `edge://extensions`), turn on developer
mode, choose **Load unpacked** and pick `integrations/office-js-extension/`.

## What an add-in can do

What `oxi-office.js` implements: ranges (values, formulas, number formats,
text, fill, font, alignment, wrapping, column width, row height), worksheets
(list, get, add, activate, used range), names, tables (add, parts, rows),
`onChanged` and `onSelectionChanged`, and the `load()` / `context.sync()`
bookkeeping, including `PropertyNotLoaded` for reading before a sync. The
editor applies values as undoable edits and the rest through its macro engine.
