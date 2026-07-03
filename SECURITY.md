# Security Policy

## Reporting a Vulnerability

Please report suspected vulnerabilities privately via
**[GitHub Security Advisories](../../security/advisories/new)** (preferred) or by
email to the maintainer. Do not open a public issue for security reports.

You can expect an acknowledgement within 7 days. Please include a minimal
reproducing document where possible — a crafted `.docx` / `.xlsx` / `.pptx` /
`.pdf` that triggers the behavior.

## Supported Versions

Oxi is pre-1.0; only the latest release line receives security fixes.

## Threat Model

Oxi's job is to parse and render untrusted Office documents. The engine is
designed so that a hostile document cannot do more than render incorrectly:

| Property | Design |
|---|---|
| **Memory safety** | The four core crates (`oxi-common`, `oxidocs-core`, `oxicells-core`, `oxislides-core`) contain **zero `unsafe` blocks**. Platform `unsafe` is confined to the optional native renderers (GDI / DirectWrite bindings), which are not part of the WASM build. |
| **No active content** | Macros (`vbaProject.bin`), embedded scripts, and OLE objects are **never executed** — they are treated as opaque bytes and preserved for round-trip. Field instructions (`w:instrText`) are never evaluated; only `PAGE` / `NUMPAGES` are substituted from the engine's own layout state, and other fields display their cached result text. |
| **No XXE / no entity expansion** | XML is parsed with [`quick-xml`](https://crates.io/crates/quick-xml) in streaming mode. External entities and DTDs are never fetched or expanded. |
| **No network I/O** | Parsing and layout perform no network requests. In the browser build, every byte stays in the WASM sandbox on the client — documents never leave the machine. |
| **Constrained decompression** | OOXML containers are opened with the `zip` crate built with `default-features = false` (deflate only — no executable-adjacent codecs). |
| **Fonts** | Document-embedded fonts are **not loaded or rasterized**. Text is measured against pre-computed metric tables shipped with the engine, so a malicious font file in a document is inert. |

## Hardening Roadmap

- Structured fuzzing of the OOXML/PDF parsers (cargo-fuzz) beyond the current
  504-file real-world parse corpus.
- Decompression-ratio limits for pathological (zip-bomb-shaped) containers.
- `cargo audit` / `cargo deny` in CI.
