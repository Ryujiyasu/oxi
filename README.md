# Oxi

Open-source document processing suite built with Rust + WebAssembly
View, render, and edit .docx / .xlsx / .pptx / PDF natively in the browser — no server required

[Live Demo](https://ryujiyasu.gitlab.io/oxi/) · [Roadmap](#roadmap) · [Contributing](#contributing)

![MIT License](https://img.shields.io/badge/license-MIT-blue) ![Rust 1.93+](https://img.shields.io/badge/rust-1.93%2B-orange) ![wasm-pack 0.14](https://img.shields.io/badge/wasm--pack-0.14-green)

---

## Features

- Parse .docx, .xlsx, .pptx, PDF into a language-agnostic Intermediate Representation (IR)
- Render documents with a layout engine (paragraphs, tables, images, headers/footers, page borders)
- Edit text in .docx / .xlsx / .pptx with round-trip fidelity — original XML is preserved
- Download edited files — changes are patched into the original ZIP, not rebuilt from scratch
- PDF text extraction, structure parsing, and PDF generation
- Japanese typography — kinsoku shori (JIS X 4051), CJK font metrics
- Hanko / Inkan — Japanese digital stamp generation (round, square, oval) + PAdES PDF signatures
- 100% client-side — all processing runs in WebAssembly, nothing leaves your browser

[Try it now: Live Demo](https://ryujiyasu.gitlab.io/oxi/)

---

## Mission

Billions of people depend on proprietary document formats (.docx, .xlsx, .pptx, .pdf) for work, education, and government — yet no truly compatible open-source rendering engine exists. LibreOffice breaks layouts. Google Docs requires a server. The world deserves a document engine that is:

- **Free forever** — MIT license, no vendor lock-in
- **Runs anywhere** — browser, desktop, mobile, server, embedded
- **High-fidelity rendering** — based on published OOXML standards and COM API measurements
- **Private by design** — your data never leaves your device
- **Accessible** — works on low-end hardware, no installation required

Oxi is built to close this gap.

---

## Vision: The Oxi Ecosystem

Oxi's core mission is one thing: **pixel-perfect document rendering**. Everything else is an extension or a fork.

### Extensions
Extensions add functionality on top of Oxi's rendering core without compromising pixel accuracy.

Examples:
- **oxi-hyde** — Hardware-backed encryption and PQC signatures (TPM 2.0 + ML-KEM-768) for document integrity
- **oxi-mcp** — Model Context Protocol integration for AI agent workflows
- **oxi-tauri** — Desktop application wrapper

### Forks
Forks are purpose-built derivatives of Oxi targeting specific use cases and user communities.

**Example: University Oxi**
AI-generated documents are a growing challenge in academic settings. University Oxi envisions binding the *process* of document creation — the dialogue with AI, the edit history, the human revisions — to the final submission. Using oxi + hyde + argo (ZKP), a student can prove not just *what* was submitted, but *how* it was created, without exposing the full content. This is not about banning AI — it's about making AI use transparent and verifiable.

We welcome forks that serve communities Oxi's core cannot.

---

## 100% Clean-Room Implementation

Oxi's rendering engine was built without any disassembly, decompilation, or binary analysis of proprietary software.

All layout specifications are derived exclusively from two sources:

1. **Published standards** — OOXML (ISO/IEC 29500 / ECMA-376), PDF (ISO 32000)
2. **Black-box testing** — Observing output values via the Microsoft Office COM API

AI (Claude) was used throughout the specification derivation process — including root-cause analysis, COM API measurement, pattern confirmation, and fix implementation. All specification decisions are grounded exclusively in values measured via the COM API. Human review confirmed correctness at each stage.

Under Microsoft's [Open Specification Promise](https://learn.microsoft.com/en-us/openspecs/dev_center/ms-devcentlp/1c24c7c8-28b0-4ce1-a47d-95fe1ff504bc), no patents are asserted against implementations of the OOXML specification.

---

## Font Rendering Strategy

Oxi targets pixel-perfect rendering using open-licensed fonts only — no proprietary fonts are bundled or required.

**Two-tier approach:**

1. **OpenFont baseline** — All metrics (advance width, line height, kerning) are matched exactly to their Microsoft Font counterparts. Within this baseline, 100% pixel-identical rendering is guaranteed and verified by automated tests.

2. **Real-world documents** — Documents authored with Microsoft Fonts (Calibri, MS Gothic, etc.) are assessed using a **font divergence score**: a per-glyph pixel-diff table (generated once on a system with Microsoft Fonts installed) combined with character frequency in the target document. This produces a machine-calculable visual fidelity estimate without requiring proprietary fonts at runtime.

This approach keeps Oxi fully open-source and CI-friendly while honestly quantifying rendering fidelity for real-world documents.

---

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
  font-glyph-diff-gen/ Per-glyph pixel-diff table for Microsoft Font divergence scoring
  metrics/            Line-height analysis scripts and data
tests/fixtures/       Test .docx / .xlsx / .pptx files
```

---

## IR Design

The Intermediate Representation is language-agnostic and does not depend on Word/Excel/PowerPoint internals:

```
Document → Page → Block (Paragraph | Table | Image) → Run
```

---

## Round-Trip Editing

Original ZIP archives are preserved. Only the specific XML text nodes that changed are patched:

| Format | Coordinate System | Patched Element |
|--------|-------------------|-----------------|
| .docx  | (paragraph, run)  | `<w:t>` text nodes |
| .xlsx  | (sheet, row, col) | `<c>` cell values (inline string) |
| .pptx  | (slide, shape, paragraph, run) | `<a:t>` text nodes |

---

## WASM API

All processing is exposed via wasm-bindgen and can be called directly from JavaScript:

```javascript
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

---

## Quick Start

### Prerequisites
- Rust 1.93+
- wasm-pack 0.14+

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

---

## Tech Stack

| Layer | Technology |
|-------|------------|
| Core engines | Rust (memory-safe, zero-cost abstractions) |
| XML parsing | quick-xml |
| ZIP handling | zip crate |
| Serialization | serde / serde_json |
| Browser bindings | wasm-bindgen + wasm-pack |
| Font metrics | Generated at build time from user's local system fonts via tools/font-metrics-gen |
| Web demo | Vanilla JS + Canvas (no framework dependencies) |

---

## Golden Tests — 504 Files, 100% Parse Success

Tested against 504 real-world government documents (Japanese ministries) + generated files:

|        | Oxi    | LibreOffice |
|--------|--------|-------------|
| Overall | 100.0% | 99.2% |
| DOCX   | 100.0% | 100.0% |
| XLSX   | 100.0% | 98.6% |
| PPTX   | 100.0% | 100.0% |

LibreOffice timed out (>45s) on 4 large government xlsx files. Oxi parsed all instantly.

---

## Roadmap

### v1 — Foundation (current)
- .docx / .xlsx / .pptx parser & language-agnostic IR
- .docx layout engine (paragraphs, tables, images, headers/footers, page borders)
- Japanese typography (kinsoku shori, CJK punctuation compression)
- Round-trip editing (.docx structural editing, .xlsx/.pptx basic text editing)
- PDF parse, text extraction, generation, PAdES signatures
- Hanko (Japanese digital stamp) SVG generation
- WASM build + web demo
- Basic formula evaluation (.xlsx: SUM, AVERAGE, IF, etc.)

### Next
- .docx layout accuracy improvements (font metrics, CJK justification)
- .xlsx layout engine (cell rendering, charts)
- .pptx layout engine (slide rendering, masters)
- Vertical writing & ruby (furigana)
- oxi-hyde integration (standard Extension)

---

## Contributing

Contributions are welcome. Oxi has a simple acceptance criterion:

**Every merged PR must improve the pixel accuracy of at least one document.**

### What belongs in core
1. Pixel accuracy improvements to existing layout engine
2. New test documents with low pixel accuracy (must use OpenFont, improvement tracked via Issue)
3. New OpenFont additions (Microsoft Font metric parity verification required)
4. Format engine additions: .xlsx layout, .pptx layout, vertical writing, etc.

### What belongs in an Extension or Fork
Features that go beyond pixel-accurate rendering — collaboration, AI integration, desktop apps, purpose-specific workflows — belong in an Extension or Fork, not in Oxi core.

See [Vision: The Oxi Ecosystem](#vision-the-oxi-ecosystem) for the distinction between Extensions and Forks.

### How to contribute
1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Run tests and lint (`cargo test && cargo clippy`)
4. Submit a pull request with pixel accuracy results

---

## Why Rust + Wasm?

- **Performance** — native-speed document parsing and layout in the browser
- **Memory safety** — no buffer overflows, no use-after-free, no data races
- **Small binary** — the compiled .wasm is ~1.4 MB for the entire suite
- **Zero server cost** — all processing runs client-side, no backend needed
- **Privacy** — documents never leave the user's device

---

## License

MIT
