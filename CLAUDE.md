# Oxidocs Development Guide

## Project Overview

Oxidocs is a Word-compatible OSS document processing engine built with Rust + WebAssembly.
The goal is to parse, render, and eventually edit .docx files natively in the browser.

## Architecture

- **oxidocs-core**: Core engine — parser, IR, layout, font metrics
- **oxidocs-wasm**: WebAssembly bindings via wasm-bindgen
- **web/**: React + Canvas demo app

## IR Design Principles

The Intermediate Representation (IR) must be language-agnostic and NOT depend on Word-specific internals.
Structure: Document → Page → Block (Paragraph | Table | Image) → Run

## Font Metrics

Font files are NEVER committed to the repository. Only pre-computed metrics tables are included.
Metrics are measured on GitHub Actions Windows runners and stored as data tables.

## Japanese Typography (Kinsoku)

Priority order:
1. Kinsoku processing (line-start/line-end prohibited characters)
2. Character spacing (justification)
3. Ruby (furigana)
4. Vertical writing (basics only)

Reference: JIS X 4051

## Testing

- Golden tests: render .docx with Oxidocs, compare pixel-by-pixel against Word screenshots
- Test fixtures go in tests/fixtures/
- CI: `cargo test`, `cargo clippy`, `wasm-pack build`

## Build Commands

```bash
cargo build                          # Build all
cargo test                           # Run tests
cargo clippy                         # Lint
cd crates/oxidocs-wasm && wasm-pack build --target web  # Wasm build
```

## License

MIT. All third-party crate licenses must be MIT-compatible.
