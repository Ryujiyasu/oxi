<p align="center">
  <img src="docs/oxi-logo.png" alt="Oxi" width="400">
</p>

<p align="center">
  <b>OSS document processing suite built with Rust + WebAssembly.</b><br>
  View &amp; edit .docx / .xlsx / .pptx natively in the browser — no server required.
</p>

<p align="center">
  <a href="https://ryujiyasu.github.io/oxi/"><strong>Live Demo</strong></a>
</p>

---

## Features

- **Parse** .docx, .xlsx, .pptx into a language-agnostic IR
- **Render** documents with layout engine (paragraphs, tables, images, kinsoku)
- **Edit** text in all three formats with round-trip fidelity (original XML preserved)
- **Download** edited files — changes are patched into the original ZIP, not rebuilt
- **100% client-side** — all processing runs in WebAssembly, nothing leaves your browser

## Architecture

```
crates/
  oxi-common/         # Shared OOXML utilities (archive, relationships, XML helpers)
  oxidocs-core/       # .docx engine — parser, IR, layout, font metrics, editor
  oxicells-core/      # .xlsx engine — parser, IR, editor
  oxislides-core/     # .pptx engine — parser, IR, editor
  oxi-wasm/           # WebAssembly bindings (wasm-bindgen)
web/                   # Web demo (vanilla JS + Canvas)
tests/fixtures/        # Test .docx / .xlsx / .pptx files
```

### Editing approach

Round-trip editing preserves the original ZIP archive. Only the specific XML text nodes that changed are patched:

| Format | Coordinate system | Patched element |
|--------|------------------|----------------|
| .docx | (paragraph, run) | `<w:t>` text nodes |
| .xlsx | (sheet, row, col) | `<c>` cell values (inline string) |
| .pptx | (slide, shape, paragraph, run) | `<a:t>` text nodes |

## Quick Start

```bash
# Build & test
cargo build
cargo test

# Build Wasm
cd crates/oxi-wasm
wasm-pack build --target web

# Serve web demo
cd web
python3 -m http.server 8080
```

## Roadmap

### Done
- [x] .docx / .xlsx / .pptx parsers
- [x] Language-agnostic IR for all three formats
- [x] Layout engine (paragraph, table, image)
- [x] Japanese typography (kinsoku shori)
- [x] Wasm build + web demo
- [x] Round-trip editing for .docx / .xlsx / .pptx

### Next
- [ ] Justify / advanced font metrics (.docx)
- [ ] Formula evaluation / cell merge / charts (.xlsx)
- [ ] Slide masters / transitions / animation (.pptx)
- [ ] Vertical writing (tate-gaki)
- [ ] Ruby (furigana)

## Why Rust + Wasm?

- **Rust**: memory safety, performance, modern tooling
- **Wasm**: runs natively in the browser, no install required

## Contributing

All contributions welcome. See CONTRIBUTING.md.

## License

MIT
