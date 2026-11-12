# oxidocs

Convert `.docx` to PDF in Python without LibreOffice or Microsoft Word.

```sh
pip install oxidocs
```

```python
import oxidocs

result = oxidocs.docx_to_pdf("report.docx", "report.pdf")
print(result.pages)
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

## Command line

Installing the package also puts the `oxidocs` command on your PATH:

```sh
oxidocs docx-to-pdf report.docx report.pdf
oxidocs docx-to-pdf --strict report.docx report.pdf   # exit status 3 on content Oxi does not draw
```

`python -m oxidocs ...` runs the same command. The `OXIDOCS_BINARY` environment variable points the package at another oxidocs binary.

## How faithful is it?

We froze 50 English and 50 Japanese documents from the public [docx-corpus](https://docxcorp.us) before any engine rendered them. Each was converted by this engine and by LibreOffice 26.8.2.2 on the same Linux machine, and both outputs were scored against Microsoft Word's own PDF of the same file.

| | English SSIM | page count matches Word | Japanese SSIM | page count matches Word |
|---|---|---|---|---|
| Oxi | 0.884 | 47 / 50 | 0.797 | 45 / 50 |
| LibreOffice | 0.841 | 39 / 50 | 0.769 | 26 / 50 |

Over 4,988 documents, the median conversion took 0.08 s, against 0.89 s for LibreOffice. The method and the full tables are in [integrations/gotenberg](https://github.com/Ryujiyasu/oxi/tree/main/integrations/gotenberg#measured).

Fonts: Oxi does not bundle Microsoft fonts. It lays text out with metrics measured from Word's fonts and draws it with metric-compatible open fonts. Line and page breaks therefore do not depend on which fonts the machine has installed, but the glyph shapes differ from Microsoft's.

## License

The bundled `oxidocs` binary is the Oxi engine under MPL-2.0. This package's Python code is MIT OR Apache-2.0.
