# Oxidocs

A Word-compatible `.docx` layout and rendering engine in Rust, built for
browser-native (WASM) and headless use.

**Early development.** The API is 0.x and unstable; this release exists to
anchor the crate name while the engine is developed in the open at
<https://gitlab.com/Ryujiyasu/oxi>.

## What it does

- Parses OOXML WordprocessingML (`.docx`) into a language-agnostic IR
- Lays out documents with a line/page model measured against Microsoft
  Word's own rendering — including Japanese typography: kinsoku line
  breaking, document grid (docGrid), ruby, vertical writing
- Renders pages (Canvas/GDI/DirectWrite backends in the repository)

## How fidelity is measured

Layout is verified continuously against Word as the oracle:

- **Pagination equality** — per-paragraph page placement matches Word on a
  corpus of real-world Japanese government/legal documents (100% on the
  current 87-document corpus)
- **Pixel comparison** — per-page SSIM against Word-rendered references

## License

MPL-2.0
