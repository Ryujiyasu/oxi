# Oxi Roadmap

> **Note**: v2+ are long-term goals, not commitments. Priorities may change based on community feedback and contributor interest.

---

## v1 — Foundation (current)

### Implemented
- .docx / .xlsx / .pptx parser & language-agnostic IR
- .docx layout engine (paragraphs, tables, images, headers/footers, page borders, floating elements)
- Japanese typography (kinsoku shori, CJK punctuation compression)
- Round-trip editing (.docx structural editing, .xlsx/.pptx basic text editing)
- PDF 1.7 parse, text extraction, generation
- PAdES / PKCS#7 digital signatures
- Hanko (Japanese digital stamp) SVG generation + PDF signing integration
- WASM build + unified Canvas editor (click-to-edit, instant re-layout at 9ms)
- Basic formula evaluation (.xlsx: SUM, AVERAGE, IF, MIN, MAX, COUNT, COUNTIF)
- Ra autonomous specification loop (COM-measured Word compatibility)
- FontMetricsRegistry caching (43x layout speedup)

### In progress
- .docx layout accuracy → SSIM 0.95+ (Ra loop continues)
- IME (Japanese/CJK input) support in Canvas editor
- Text selection & formatting toolbar integration
- .xlsx cell rendering and chart support
- .pptx slide layout engine and masters
- Vertical writing & ruby (furigana)

---

## v2 — oxidocs Native Format

**Architectural Guarantee: oxidocs core layer can always be losslessly converted to .docx.**

### oxidocs format

```
oxidocs
├── core layer   Word-compatible fields. Owned by Oxi core. Forks cannot modify.
│                Always exportable to .docx.
└── ext layer    Fork/Extension additions. On .docx export: customXml/ or discard.
```

### Output patterns

| Pattern | Format | Use case | Compatibility |
|---------|--------|----------|---------------|
| A | .docx + oxi extensions in customXml/ | External sharing, submission, archive | Opens in all Word clients |
| B | .oxidocs (native, optimized) | Internal storage, Fork sharing, waterdocs base | Smaller than .docx. Always convertible to Pattern A |

### Dual font engine

| Format | Engine | Reason |
|--------|--------|--------|
| .docx | GDI | Word uses GDI metrics. Pixel-identical layout requires matching GDI behavior. |
| .oxidocs | DirectWrite | Platform-independent floating-point precision. Variable fonts, high-DPI. |
| .pdf | DirectWrite | PDF spec uses floating-point coordinates. |

### Deliverables
- [ ] oxidocs specification (core layer / ext layer definition)
- [ ] oxidocs ↔ .docx bidirectional conversion (Pattern A / B)
- [ ] oxidocs-to-docx Generator (full generation, no original file required)
- [ ] Dual font engine: GDI for .docx, DirectWrite for .oxidocs
- [ ] docs/governance.md: Architectural Guarantee formalized

### Critical path: oxidocs-to-docx Generator

Without a complete generator that produces valid .docx from any oxidocs without the original file, the Architectural Guarantee is aspirational, not real. `create_blank_docx` is the foundation. Ra's reverse-engineered specifications feed directly into generation rules.

---

## v2.x — waterdocs + Collaboration

### waterdocs (oxidocs + hyde encryption)

```
waterdocs
├── core layer (after decryption → .docx exportable)
└── hyde layer (encryption metadata, not included in .docx)
```

- [ ] waterdocs format definition (oxi-hyde Extension)
- [ ] Encryption/decryption flow preserving core layer guarantee

### Real-time collaboration

Real-time co-editing using CRDTs.

- `oxi-collab` crate: CRDT document model using [yrs](https://github.com/y-crdt/y-crdt) (Yjs Rust port)
- Plain text and formatting sync across browsers
- Cursor/selection sharing (Awareness protocol)
- Lightweight WebSocket relay server (stateless)
- E2E encryption (zero-knowledge relay)
- PWA / offline support with auto-merge on reconnect

```
Browser A ◄──── WebSocket ────► Relay (stateless) ◄──── WebSocket ────► Browser B
   oxi-wasm + yrs                                           oxi-wasm + yrs
```

---

## v3 — Platform

Evolve from a document tool into a productivity platform.

- Plugin system (WASM-sandboxed third-party extensions)
- Desktop app via Tauri (Windows, macOS, Linux)
- Mobile app (iOS, Android)
- Workflow automation (approval flows, template-based generation)
- AI-assisted document processing (summarization, translation, form filling)
- oxi-argo (zero-knowledge proofs for document provenance)
- oxi-mcp (AI agent workflow integration)

---

## v4 — Enterprise

- Enterprise interoperability (SharePoint, Google Drive, WebDAV connectors)
- Compliance & governance (audit trails, DLP, GDPR/APPI)
- Industry verticals via Forks (legal, healthcare, education, government)
- Developer ecosystem (REST API, CLI, npm/crates.io packages)

---

## Version Summary

| Version | Theme | Status |
|---------|-------|--------|
| **v1** | Foundation — parse, render, edit in browser | **Active development** |
| **v1.x** | Word parity — SSIM 0.95+, IME, selection | In progress |
| **v2** | oxidocs native format + dual font engine | Planned |
| **v2.x** | waterdocs + real-time collaboration | Planned |
| **v3** | Platform — plugins, desktop/mobile, AI | Exploratory |
| **v4** | Enterprise — compliance, integrations | Long-term vision |
