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

Oxi parses and renders untrusted Office documents. The browser engine runs
in WebAssembly; spreadsheet VBA execution is provided by a separate host API.
The relevant boundaries and current hardening limits are listed below:

| Property | Design |
|---|---|
| **Memory safety / platform bindings** | Browser processing uses Rust and WebAssembly. Native Windows font and renderer integrations use platform FFI behind target-specific configuration; those Windows bindings are not part of the browser runtime. |
| **VBA / active content** | Parsing preserves VBA projects for round-trip. Supported macros can execute through the spreadsheet VBA host, which interprets VBA using its own APIs rather than Microsoft Office or native Windows COM. The document renderer does not execute OLE objects or embedded scripts. |
| **No XXE / no entity expansion** | XML is parsed with [`quick-xml`](https://crates.io/crates/quick-xml) in streaming mode. External entities and DTDs are never fetched or expanded. |
| **Network boundary** | Parsing and layout do not make network requests. Embedding applications control document loading, sharing and any other network activity. |
| **Constrained decompression** | OOXML containers are opened with the `zip` crate built with `default-features = false` (deflate only — no executable-adjacent codecs). |
| **Fonts** | Layout uses pre-computed metric tables. Browser font delivery and native font access depend on the host integration. |

## Hardening Roadmap

- Structured fuzzing of the OOXML/PDF parsers (cargo-fuzz) beyond the current
  504-file real-world parse corpus.
- Decompression-ratio limits for pathological (zip-bomb-shaped) containers.
- `cargo audit` / `cargo deny` in CI.
