# oxidocs

Convert `.docx`, `.pptx` and `.xlsx` to PDF in Python without LibreOffice or Microsoft Office.

```sh
pip install oxidocs
```

```python
import oxidocs

result = oxidocs.docx_to_pdf("report.docx", "report.pdf")
print(result.pages)

oxidocs.pptx_to_pdf("deck.pptx", "deck.pdf")   # one page per slide
oxidocs.xlsx_to_pdf("report.xlsx", "report.pdf")  # visible sheets, in print order
```

The package bundles the [Oxi](https://github.com/Ryujiyasu/oxi) engine as a single binary. Oxi is an open-source Rust engine that lays out Word documents the way Word does. There is no office suite to install and no server to run. The document is converted on the machine that holds it.

## Many documents

A `Converter` keeps one engine process running and can be shared between threads:

```python
with oxidocs.Converter(threads=4) as converter:
    for name in names:
        converter.convert(f"{name}.docx", f"{name}.pdf")
```

## Knowing when to fall back

With `strict=True`, a document that holds content Oxi does not draw raises `UnsupportedContent` instead of producing an incomplete PDF. Examples are a chart type it does not support, or SmartArt saved without its drawing. The same applies when Oxi's pages would show less than 90% of the document's own text. You can then hand that document to another converter:

```python
try:
    oxidocs.docx_to_pdf(src, dst, strict=True)
except oxidocs.UnsupportedContent as e:
    print("falling back:", e.kinds)   # e.g. ['chart']
    convert_with_libreoffice(src, dst)
```

`ConversionError` (the base class) covers files that cannot be read at all.

## Presentations

`pptx_to_pdf` writes one PDF page per slide, with text laid out by rules measured from PowerPoint. Slides are newer in Oxi than Word documents, and some content is not drawn as PowerPoint draws it yet: charts, gradient fills (drawn in one colour), dashed outlines (drawn solid), rotated text and pictures, and text warps. `strict=True` raises `UnsupportedContent` for a deck that holds any of these, as it does for documents:

```python
try:
    oxidocs.pptx_to_pdf(src, dst, strict=True)
except oxidocs.UnsupportedContent as e:
    print("falling back:", e.kinds)   # e.g. ['chart']
```

A `Converter` takes documents, presentations and spreadsheets alike.

## Spreadsheets

`xlsx_to_pdf` exports visible worksheets in workbook order. It supports cell text
(including Japanese), saved formula results, number formats, solid backgrounds,
simple borders, alignment, wrapping, shrink-to-fit, merged cells and hidden rows
and columns. Printing uses each sheet's cell-range print areas, Letter/Legal/A3/A4/A5
paper size, orientation, margins, scale, fit-to-page settings, centering and manual
page breaks. Separate print areas start on separate pages. Empty sheets are skipped.

```python
oxidocs.xlsx_to_pdf("report.xlsx", "report.pdf", strict=True)
```

The engine does not recalculate formulas or run VBA; save updated calculation
results before conversion. Macro-enabled workbooks (`.xlsm`) are accepted too.
This first spreadsheet PDF path does not establish Excel print-layout parity:
row heights use the saved values and unwrapped text is clipped to its cell.
Drawings/charts, table styles, conditional formatting, rich or rotated text,
complex borders, merges crossing print boundaries, repeating print titles, headers/footers and print headings are
reported as unsupported. `strict=True` refuses these before writing the output;
without strict the engine warns and writes the supported content.

CLI: `oxidocs xlsx-to-pdf [--strict] [--quiet] input.xlsx output.pdf`.

### Excel の PDF 出力

`oxidocs.xlsx_to_pdf("帳票.xlsx", "帳票.pdf")` で、表示中のシートを順に
PDFへ出力できます。印刷範囲、用紙サイズ・向き、余白、倍率、ページへの
縮小、手動改ページ、結合セル、数値書式、背景色、単純な罫線に対応します。
数式には保存済みの計算結果を使い、再計算やマクロ実行は行いません。
Excelとの完全な印刷一致を保証する段階ではありません。グラフ・図形、
条件付き書式、印刷タイトル、ヘッダー／フッターなどは未対応として報告し、
`strict=True` を指定すると出力前に `UnsupportedContent` で停止します。

## Pages without drawing: which page is each paragraph on?

A .docx stores no pages: Word works them out when it lays the document out. Tools that read .docx therefore cannot say on which page a heading or a paragraph sits, and a search or RAG pipeline cites every passage of a long Word file as "page 1". The page break markers Word leaves in a file are only as fresh as its last save in Word, and files written by programs have none.

`docx_to_docling` lays the document out as Word would and returns its structure and text as a [DoclingDocument](https://github.com/docling-project/docling-core), nothing drawn. Every heading, paragraph, list item, table and picture carries the page it is on and its box there:

```python
import oxidocs

doc = oxidocs.docx_to_docling("contract.docx")
for item in doc["texts"]:
    for p in item["prov"]:
        print(p["page_no"], p["bbox"], item["text"][:40])
```

With [docling-core](https://pypi.org/project/docling-core/) installed, the same dict loads as a `DoclingDocument`, so docling's own exports work page by page:

```python
from docling_core.types.doc import DoclingDocument

doc = DoclingDocument.model_validate(oxidocs.docx_to_docling("contract.docx"))
page_3 = doc.export_to_markdown(page_no=3)
```

Boxes are in points from the page's top left (`coord_origin: "TOPLEFT"`). An item that runs onto the next page has one entry per page. How headings, lists, tables, text boxes and equations are mapped is described in [oxidocs-docling](https://github.com/Ryujiyasu/oxi/tree/main/crates/oxidocs-docling#readme). Pages come from the same layout as the PDF output, so they are as close to Word's as the figures under [How faithful is it?](#how-faithful-is-it) show.

## Command line

Installing the package also puts the `oxidocs` command on your PATH:

```sh
oxidocs docx-to-pdf report.docx report.pdf
oxidocs docx-to-pdf --strict report.docx report.pdf   # exit status 3 on content Oxi does not draw
oxidocs pptx-to-pdf deck.pptx deck.pdf
oxidocs docx-to-docling report.docx report.json        # stdout without the second path
```

`python -m oxidocs ...` runs the same command. The `OXIDOCS_BINARY` environment variable points the package at another oxidocs binary.

## How faithful is it?

The figures below are for Word documents. Slides have no such measurement for this PDF output yet.

We froze 50 English and 50 Japanese documents from the public [docx-corpus](https://docxcorp.us) before any engine rendered them. Each was converted by this engine and by LibreOffice 26.8.2.2 on the same Linux machine, and both outputs were scored against Microsoft Word's own PDF of the same file.

| | English SSIM | page count matches Word | Japanese SSIM | page count matches Word |
|---|---|---|---|---|
| Oxi | 0.884 | 47 / 50 | 0.797 | 45 / 50 |
| LibreOffice | 0.841 | 39 / 50 | 0.769 | 26 / 50 |

Over 4,988 documents, the median conversion took 0.08 s, against 0.89 s for LibreOffice. The method and the full tables are in [integrations/gotenberg](https://github.com/Ryujiyasu/oxi/tree/main/integrations/gotenberg#measured).

Fonts: Oxi does not bundle Microsoft fonts. It lays text out with metrics measured from Word's fonts and draws it with metric-compatible open fonts. Line and page breaks therefore do not depend on which fonts the machine has installed, but the glyph shapes differ from Microsoft's.

## License

The bundled `oxidocs` binary is the Oxi engine under MPL-2.0. This package's Python code is MIT OR Apache-2.0. The equation converter used by `docx_to_docling` is a port of docling's and dwml's OMML converters (MIT and Apache-2.0); their notices ship in the package as `oxidocs/THIRD_PARTY_NOTICES.md`.
