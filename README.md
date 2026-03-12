# oxidocs

Word-compatible OSS document engine (Rust + Wasm)

## Overview

oxidocs is an open-source document engine that generates Word-compatible `.docx` files with pixel-accurate layout. Built in Rust with WebAssembly support for browser-based rendering.

## Project Structure

```
crates/
  oxidocs-core/   # Core document model and OOXML generation
  oxidocs-wasm/   # WebAssembly bindings for browser usage
```

## Key Design Decisions

- **Font metrics only**: Font files are NOT included in the repository. Only pre-computed metrics tables are bundled.
- **License**: MIT. All third-party crate licenses are verified to be MIT-compatible.

## Building

```bash
cargo build
```

## License

MIT
