# Oxidocs

Word-compatible OSS document processing engine built with Rust + WebAssembly.

## Vision

Microsoft Word (.docx) compatibility for the browser — without Microsoft.
Built as a base by one developer, grown by the community.

## Status

Early development

## Goals (Phase 1)

- [ ] .docx parser
- [ ] Language-agnostic IR
- [ ] Basic layout engine (paragraph, table, image)
- [ ] Japanese typography (kinsoku)
- [ ] Wasm build
- [ ] Web demo

## Project Structure

```
crates/
  oxidocs-core/       # Core engine (Rust)
    src/
      parser/          # .docx parser (OOXML)
      ir/              # Intermediate Representation
      layout/          # Layout engine
      font/            # Font metrics
  oxidocs-wasm/        # Wasm bindings
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
cd crates/oxidocs-wasm
wasm-pack build --target web
```

## Contributing

All contributions welcome. See CONTRIBUTING.md.

## License

MIT
