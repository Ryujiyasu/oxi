<p align="center">
  <img src="docs/oxi-logo.png" alt="Oxi" width="400">
</p>

<p align="center">
  <b>OSS document processing suite built with Rust + WebAssembly.</b><br>
  View &amp; edit .docx / .xlsx / .pptx / PDF natively in the browser — no server required.
</p>

<p align="center">
  <a href="https://ryujiyasu.github.io/oxi/"><strong>Live Demo</strong></a> ·
  <a href="docs/roadmap.md"><strong>Roadmap</strong></a> ·
  <a href="#contributing"><strong>Contributing</strong></a>
</p>

<p align="center">
  <a href="LICENSE"><img src="https://img.shields.io/badge/license-MIT-blue.svg" alt="MIT License"></a>
  <img src="https://img.shields.io/badge/rust-1.93+-orange.svg" alt="Rust 1.93+">
  <img src="https://img.shields.io/badge/wasm--pack-0.14-purple.svg" alt="wasm-pack 0.14">
</p>

---

## Features

- **Parse** .docx, .xlsx, .pptx, PDF into a language-agnostic IR
- **Render** documents with a layout engine (paragraphs, tables, images, headers/footers, page borders)
- **Edit** text in .docx / .xlsx / .pptx with round-trip fidelity — original XML is preserved
- **Download** edited files — changes are patched into the original ZIP, not rebuilt from scratch
- **PDF** text extraction, structure parsing, and PDF generation from scratch
- **Japanese typography** — kinsoku shori (JIS X 4051), MS Gothic / MS Mincho / Yu Gothic font metrics
- **Hanko / Inkan** — Japanese digital stamp generation and PAdES PDF signatures
- **100% client-side** — all processing runs in WebAssembly, nothing leaves your browser

## Architecture

```
crates/
  oxi-common/         Shared OOXML utilities (ZIP, XML, relationships)
  oxidocs-core/       .docx engine — parser, IR, layout, font metrics, editor
  oxicells-core/      .xlsx engine — parser, IR, editor
  oxislides-core/     .pptx engine — parser, IR, editor
  oxipdf-core/        PDF 1.7 engine — parser, text extraction, generator
  oxihanko/           Japanese digital stamp (hanko) generator + PAdES signer
  oxi-wasm/           WebAssembly bindings (wasm-bindgen)
web/                  Web demo (vanilla JS + Canvas)
tools/
  font-metrics-gen/   Standalone tool to extract font metrics from system fonts
  metrics/            Line-height analysis scripts and data
tests/fixtures/       Test .docx / .xlsx / .pptx files
```

### IR design

The Intermediate Representation is language-agnostic and does not depend on Word/Excel/PowerPoint internals:

```
Document → Page → Block (Paragraph | Table | Image) → Run
```

### Round-trip editing

Original ZIP archives are preserved. Only the specific XML text nodes that changed are patched:

| Format | Coordinate system | Patched element |
|--------|------------------|----------------|
| .docx | (paragraph, run) | `<w:t>` text nodes |
| .xlsx | (sheet, row, col) | `<c>` cell values (inline string) |
| .pptx | (slide, shape, paragraph, run) | `<a:t>` text nodes |

### WASM API

All processing is exposed via `wasm-bindgen` and can be called directly from JavaScript:

```js
import init, {
  parse_document,       // .docx → IR (JSON)
  parse_spreadsheet,    // .xlsx → IR (JSON)
  parse_presentation,   // .pptx → IR (JSON)
  parse_pdf,            // PDF → structure (JSON)
  pdf_extract_text,     // PDF → plain text
  layout_document,      // .docx → positioned layout with coordinates
  edit_docx,            // apply text edits → new .docx bytes
  edit_xlsx,            // apply cell edits → new .xlsx bytes
  edit_pptx,            // apply slide edits → new .pptx bytes
  create_blank_docx,    // generate empty .docx
  create_pdf,           // generate PDF from scratch
} from "./oxi_wasm.js";

await init();

const response = await fetch("sample.docx");
const bytes = new Uint8Array(await response.arrayBuffer());
const ir = parse_document(bytes);     // language-agnostic IR as JSON
const layout = layout_document(bytes); // positioned elements for canvas rendering
```

## Quick Start

### Prerequisites

- [Rust](https://rustup.rs/) 1.93+
- [wasm-pack](https://rustwasm.github.io/wasm-pack/installer/) 0.14+

### Build & test

```bash
cargo build                          # Build all crates
cargo test                           # Run tests
cargo clippy                         # Lint
```

### Build Wasm & run demo

```bash
cd crates/oxi-wasm
wasm-pack build --target web         # Build .wasm + JS bindings

cd ../../web
python3 -m http.server 8080          # Serve at http://localhost:8080
```

## Tech Stack

| Layer | Technology |
|-------|-----------|
| Core engines | Rust (memory-safe, zero-cost abstractions) |
| XML parsing | `quick-xml` |
| ZIP handling | `zip` crate |
| Serialization | `serde` / `serde_json` |
| Browser bindings | `wasm-bindgen` + `wasm-pack` |
| Font metrics | Pre-computed from Windows system fonts (13 fonts, ~55 KB JSON) |
| Web demo | Vanilla JS + Canvas (no framework dependencies) |

## Japanese Typography

Oxi implements Japanese line-breaking rules (禁則処理) per **JIS X 4051**:

- **Line-start prohibited** — closing brackets, periods, commas, small kana (。、）〕ぁぃぅ…)
- **Line-end prohibited** — opening brackets (（〔［｛…)
- **Font metrics** — MS Gothic, MS Mincho, Yu Gothic, Yu Mincho with correct win ascent/descent
- **Line height** — `max(winAsc + winDes, hheaAsc + hheaDes + hheaGap) / UPM × fontSize`
- **DocGrid snapping** — `ceil(height / pitch) × pitch` for grid-aligned layouts

## Roadmap

### v1 — Foundation (current)
- [x] .docx / .xlsx / .pptx parsers & language-agnostic IR
- [x] Layout engine (paragraphs, tables, images, headers/footers, page borders)
- [x] Japanese typography (kinsoku shori)
- [x] Round-trip editing for all three formats
- [x] PDF parsing, text extraction, generation
- [x] Hanko stamp generator + PAdES signatures
- [x] Wasm build + web demo
- [ ] Advanced font metrics & justification (.docx)
- [ ] Formula evaluation / cell merge / charts (.xlsx)
- [ ] Slide masters / transitions / animation (.pptx)
- [ ] Vertical writing (tate-gaki) & ruby (furigana)

### v2 — Collaboration
- Real-time co-editing via CRDT (yrs)
- AI-assisted document processing
- End-to-end encryption
- PWA / offline support

### v3 — Platform
- Plugin system
- Desktop & mobile apps (Tauri)
- Workflow automation
- SaaS offering

### v4 — Enterprise
- Compliance & audit trails
- Industry verticals (legal, healthcare, government)
- Developer ecosystem & marketplace

See [docs/roadmap.md](docs/roadmap.md) for the detailed roadmap.

## Why Rust + Wasm?

- **Performance** — native-speed document parsing and layout in the browser
- **Memory safety** — no buffer overflows, no use-after-free, no data races
- **Small binary** — the compiled `.wasm` is ~900 KB for the entire suite
- **Zero server cost** — all processing runs client-side, no backend infrastructure needed
- **Privacy** — documents never leave the user's device

## Contributing

All contributions are welcome! Please:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Run tests and linting (`cargo test && cargo clippy`)
4. Submit a pull request

## License

[MIT](LICENSE)
