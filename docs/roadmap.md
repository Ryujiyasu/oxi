# Oxi Roadmap

> **Note**: v2–v4 are long-term goals, not commitments. Priorities may change based on community feedback and contributor interest.

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
- WASM build + web demo
- Basic formula evaluation (.xlsx: SUM, AVERAGE, IF, MIN, MAX, COUNT, COUNTIF)

### In progress
- .docx layout accuracy (font metrics precision, CJK justification)
- .xlsx cell rendering and chart support
- .pptx slide layout engine and masters
- Vertical writing & ruby (furigana)

---

## v2 — Collaboration (planned)

Real-time co-editing using CRDTs.

- `oxi-collab` crate: CRDT document model using [yrs](https://github.com/y-crdt/y-crdt) (Yjs Rust port)
- Plain text and formatting sync across browsers
- Cursor/selection sharing (Awareness protocol)
- Lightweight WebSocket relay server
- E2E encryption (zero-knowledge relay)
- PWA / offline support with auto-merge on reconnect

### Architecture concept

```
Browser A ◄──── WebSocket ────► Relay (stateless) ◄──── WebSocket ────► Browser B
   oxi-wasm + yrs                                           oxi-wasm + yrs
```

---

## v3 — Platform (exploratory)

Evolve from a document tool into a productivity platform.

- Plugin system (WASM-sandboxed third-party extensions)
- Desktop app via Tauri (Windows, macOS, Linux)
- Mobile app (iOS, Android)
- Workflow automation (approval flows, template-based generation)
- AI-assisted document processing (summarization, translation, form filling)

---

## v4 — Enterprise (long-term vision)

- Enterprise interoperability (SharePoint, Google Drive, WebDAV connectors)
- Compliance & governance (audit trails, DLP, GDPR/APPI)
- Industry verticals (legal, healthcare, education, government)
- Developer ecosystem (REST API, CLI, npm/crates.io packages)

---

## Version Summary

| Version | Theme | Status |
|---------|-------|--------|
| **v1** | Foundation — parse, render, edit in browser | **Active development** |
| **v2** | Collaboration — real-time co-editing | Planned |
| **v3** | Platform — plugins, desktop/mobile, AI | Exploratory |
| **v4** | Enterprise — compliance, integrations | Long-term vision |
