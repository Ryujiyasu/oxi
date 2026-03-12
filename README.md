# Oxi

OSS document processing suite built with Rust + WebAssembly.

**[Live Demo](https://ryujiyasu.github.io/oxi/)**

## Vision

Microsoft Office (.docx / .xlsx / .pptx) compatibility for the browser — without Microsoft.
Built as a base by one developer, grown by the community.

## Status

Early development

## Goals (Phase 1)

- [ ] .docx parser
- [ ] .xlsx parser
- [ ] .pptx parser
- [ ] Language-agnostic IR
- [ ] Basic layout engine (paragraph, table, image)
- [ ] Japanese typography (kinsoku)
- [ ] Wasm build
- [ ] Web demo

## Project Structure

```
crates/
  oxi-common/         # Shared OOXML utilities
  oxidocs-core/       # .docx engine (Rust)
    src/
      parser/          # .docx parser (OOXML)
      ir/              # Intermediate Representation
      layout/          # Layout engine
      font/            # Font metrics
  oxicells-core/      # .xlsx engine (Rust)
  oxislides-core/     # .pptx engine (Rust)
  oxi-wasm/           # Wasm bindings
web/                   # Web demo (React + Canvas)
tests/
  fixtures/            # Test .docx files
```

## Why Rust + Wasm?

- **Rust**: memory safety, performance, modern tooling
- **Wasm**: runs natively in the browser, no install required

## Building

```bash
# Build core
cargo build

# Build Wasm
cd crates/oxi-wasm
wasm-pack build --target web
```

## Contributing

All contributions welcome. See CONTRIBUTING.md.

## License

MIT
