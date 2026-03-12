<p align="center">
  <img src="docs/oxi-logo.png" alt="Oxi" width="400">
</p>

<p align="center">
  <b>Open-source document processing suite built with Rust + WebAssembly</b><br>
  View, render, and edit .docx / .xlsx / .pptx / PDF natively in the browser — no server required
</p>

<p align="center">
  <a href="https://ryujiyasu.gitlab.io/oxi/"><strong>Live Demo</strong></a> ·
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

- **Parse** .docx, .xlsx, .pptx, PDF into a language-agnostic Intermediate Representation (IR)
- **Render** documents with a layout engine (paragraphs, tables, images, headers/footers, page borders)
- **Edit** text in .docx / .xlsx / .pptx with round-trip fidelity — original XML is preserved
- **Download** edited files — changes are patched into the original ZIP, not rebuilt from scratch
- **PDF** text extraction, structure parsing, and PDF generation
- **Japanese typography** — kinsoku shori (JIS X 4051), CJK font metrics
- **Hanko / Inkan** — Japanese digital stamp generation (round, square, oval) + PAdES PDF signatures
- **100% client-side** — all processing runs in WebAssembly, nothing leaves your browser

> Try it now: **[Live Demo](https://ryujiyasu.gitlab.io/oxi/)**

## 100% Clean-Room Implementation

Oxi's rendering engine was built without any disassembly, decompilation, or binary analysis of proprietary software.

All layout specifications are derived exclusively from two sources:

1. **Published standards** — OOXML (ISO/IEC 29500 / ECMA-376), PDF (ISO 32000)
2. **Black-box testing** — Observing output values via the Microsoft Office COM API

AI (Claude) was used throughout the specification derivation process — including root-cause analysis, COM API measurement, pattern confirmation, and fix implementation. All specification decisions are grounded exclusively in values measured via the COM API. Human review confirmed correctness at each stage.

Under Microsoft's [Open Specification Promise](https://learn.microsoft.com/en-us/openspecs/dev_center/ms-devcentlp/1c24c7c8-28b0-4ce1-a47d-95fe1ff504bc), no patents are asserted against implementations of the OOXML specification.

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

### IR Design

The Intermediate Representation is language-agnostic and does not depend on Word/Excel/PowerPoint internals:

```
Document → Page → Block (Paragraph | Table | Image) → Run
```

### Round-Trip Editing

Original ZIP archives are preserved. Only the specific XML text nodes that changed are patched:

| Format | Coordinate System | Patched Element |
|--------|------------------|----------------|
| .docx | (paragraph, run) | `<w:t>` text nodes |
| .xlsx | (sheet, row, col) | `<c>` cell values (inline string) |
| .pptx | (slide, shape, paragraph, run) | `<a:t>` text nodes |

### WASM API

All processing is exposed via `wasm-bindgen` and can be called directly from JavaScript:

```js
import init, {
  parse_document,        // .docx → IR (JSON)
  parse_spreadsheet,     // .xlsx → IR (JSON)
  parse_presentation,    // .pptx → IR (JSON)
  layout_document,       // .docx → positioned layout with coordinates
  edit_docx,             // apply text edits → new .docx bytes
  edit_xlsx,             // apply cell edits → new .xlsx bytes
  edit_pptx,             // apply slide edits → new .pptx bytes
  create_blank_docx,     // generate empty .docx
  parse_pdf,             // PDF → structure (JSON)
  pdf_extract_text,      // PDF → plain text
  create_pdf,            // generate PDF from scratch
  pdf_verify_signatures, // verify PDF signatures
  generate_hanko_svg,    // generate stamp SVG with custom config
  preview_hanko,         // quick stamp preview by name
} from "./oxi_wasm.js";

await init();
const bytes = new Uint8Array(await (await fetch("sample.docx")).arrayBuffer());
const layout = layout_document(bytes); // positioned elements for canvas rendering
```

## Mission: Democratizing Document Software

Billions of people depend on proprietary document formats (.docx, .xlsx, .pptx, .pdf) for work, education, and government — yet no truly compatible open-source rendering engine exists. LibreOffice breaks layouts. Google Docs requires a server. The world deserves a document engine that is:

- **Free forever** — MIT license, no vendor lock-in
- **Runs anywhere** — browser, desktop, mobile, server, embedded
- **High-fidelity rendering** — based on published OOXML standards and COM API measurements
- **Private by design** — your data never leaves your device
- **Accessible** — works on low-end hardware, no installation required

Oxi is built to close this gap.

## Quick Start

### Prerequisites

- [Rust](https://rustup.rs/) 1.93+
- [wasm-pack](https://rustwasm.github.io/wasm-pack/installer/) 0.14+

### Build & Test

```bash
cargo build                          # Build all crates
cargo test                           # Run tests
cargo clippy                         # Lint
```

### Build Wasm & Run Demo

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
| Font metrics | Generated at build time from user's local system fonts via `tools/font-metrics-gen` |
| Web demo | Vanilla JS + Canvas (no framework dependencies) |

## Golden Tests — 504 Files, 100% Parse Success

Tested against 504 real-world government documents (Japanese ministries) + generated files:

| | Oxi | LibreOffice |
|---|---|---|
| **Overall** | **100.0%** | 99.2% |
| DOCX | 100.0% | 100.0% |
| XLSX | **100.0%** | 98.6% |
| PPTX | 100.0% | 100.0% |

> LibreOffice timed out (>45s) on 4 large government xlsx files. Oxi parsed all instantly.

## Roadmap

### v1 — Foundation (current)
- [x] .docx / .xlsx / .pptx parser & language-agnostic IR
- [x] Layout engine (paragraphs, tables, images, headers/footers, page borders)
- [x] Japanese typography (kinsoku shori)
- [x] Round-trip editing for all 3 formats
- [x] PDF parse, text extraction, generation
- [x] Hanko generation + PAdES digital signatures
- [x] Wasm build + web demo
- [ ] Advanced font metrics & CJK justification (.docx)
- [ ] Formula evaluation / cell merging / charts (.xlsx)
- [ ] Slide masters / transitions / animations (.pptx)
- [ ] Vertical writing & ruby (furigana)

### v2 — Collaboration
- CRDT (yrs) real-time co-editing
- AI assist
- End-to-end encryption
- PWA / offline support

### v3 — Platform
- Plugin system
- Desktop & mobile apps (Tauri)
- Workflow automation

### v4 — Enterprise
- Compliance & audit trails
- Industry-specific (legal, healthcare, government)
- Developer ecosystem & marketplace

## Why Rust + Wasm?

- **Performance** — native-speed document parsing and layout in the browser
- **Memory safety** — no buffer overflows, no use-after-free, no data races
- **Small binary** — the compiled `.wasm` is ~1.4 MB for the entire suite
- **Zero server cost** — all processing runs client-side, no backend needed
- **Privacy** — documents never leave the user's device

## Contributing

Contributions welcome!

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Run tests and lint (`cargo test && cargo clippy`)
4. Submit a pull request

## License

[MIT](LICENSE)

---

<details>
<summary><strong>日本語</strong></summary>

## 特徴

- **パース** — .docx, .xlsx, .pptx, PDF を言語非依存の中間表現 (IR) に変換
- **レンダリング** — レイアウトエンジンで段落・表・画像・ヘッダー/フッター・ページ罫線を描画
- **編集** — .docx / .xlsx / .pptx のテキストをラウンドトリップ編集（元の XML を保持）
- **100% クライアントサイド** — すべて WebAssembly で処理、データはブラウザの外に出ない

> **[Live Demo](https://ryujiyasu.gitlab.io/oxi/)**

## ミッション: ドキュメントソフトウェアの民主化

世界中の数十億人がプロプライエタリなドキュメント形式に依存しています。しかし本当に互換性のあるOSSレンダリングエンジンは存在しません。Oxi はこのギャップを埋めるために作られました。

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
  // Documents
  parse_document,        // .docx → IR (JSON)
  parse_spreadsheet,     // .xlsx → IR (JSON)
  parse_presentation,    // .pptx → IR (JSON)
  layout_document,       // .docx → positioned layout with coordinates
  // Editing
  edit_docx,             // apply text edits → new .docx bytes
  edit_xlsx,             // apply cell edits → new .xlsx bytes
  edit_pptx,             // apply slide edits → new .pptx bytes
  create_blank_docx,     // generate empty .docx
  // PDF
  parse_pdf,             // PDF → structure (JSON)
  pdf_extract_text,      // PDF → plain text
  create_pdf,            // generate PDF from scratch
  pdf_verify_signatures, // verify PDF signatures
  // Hanko (Japanese stamps)
  generate_hanko_svg,    // generate stamp SVG with custom config
  preview_hanko,         // quick stamp preview by name
} from "./oxi_wasm.js";

await init();

// Document processing
const response = await fetch("sample.docx");
const bytes = new Uint8Array(await response.arrayBuffer());
const ir = parse_document(bytes);     // language-agnostic IR as JSON
const layout = layout_document(bytes); // positioned elements for canvas rendering

// Hanko stamp preview
const stamp = preview_hanko("山田");   // SVG string

// PDF generation
const pdf = create_pdf("Report", "Hello, World!");
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
| Font metrics | Generated at build time from user's local system fonts via `tools/font-metrics-gen` |
| Web demo | Vanilla JS + Canvas (no framework dependencies) |

## Hanko — Digital Japanese Stamps

The `oxihanko` crate digitizes Japan's essential business seals:

| Style | Use case | Layout |
|-------|----------|--------|
| **Round** (丸印) | Personal seal | 1-2 chars → horizontal, 3+ chars → vertical |
| **Square** (角印) | Company seal | ≤4 chars in 2×2 grid (right-to-left traditional order) |
| **Oval** (小判型) | Bank registration | Ellipse style |
| **Approval** | Date stamp | Name + date + divider lines |

- SVG output — scalable and high-resolution
- Ink colors: vermilion, red, black, or custom RGB
- Integrates with PDF signing — visible stamp embedded in PAdES signature
- Available via WASM — generate stamps directly in the browser

```js
const svg = preview_hanko("山田");                    // quick preview
const svg = generate_hanko_svg({                      // custom config
  name: "株式会社", style: "Square",
  color: { r: 227, g: 66, b: 52 }, size: 120
});
```

## PDF Engine

The `oxipdf-core` crate provides full PDF 1.7 support:

| Feature | Description |
|---------|-------------|
| **Parse** | PDF structure analysis (xref, objects, content streams, CMap) |
| **Text extraction** | Per-page and whole-document plain text |
| **Generation** | Create new PDFs with text content and metadata |
| **Digital signatures** | PAdES / PKCS#7 signing and verification |
| **Hanko signing** | Visible stamp appearance via `oxihanko` integration |

Signature providers use a plugin design (`SignatureProvider` trait), extensible to support hardware tokens like Japan's My Number Card (JPKI) in the future.

## Why Rust + Wasm?

- **Performance** — native-speed document parsing and layout in the browser
- **Memory safety** — no buffer overflows, no use-after-free, no data races
- **Small binary** — the compiled `.wasm` is ~1.4 MB for the entire suite
- **Zero server cost** — all processing runs client-side, no backend infrastructure needed
- **Privacy** — documents never leave the user's device

</details>
