# Oxi v2 Roadmap

## Vision

Oxi v2 transforms the document processing suite into a **real-time collaborative platform** — fully WASM-powered, privacy-first, and OOXML-compatible.

## Architecture Overview

```
┌─────────────┐     WebSocket      ┌──────────────┐     WebSocket      ┌─────────────┐
│  Browser A  │◄──────────────────►│ Relay Server │◄──────────────────►│  Browser B  │
│             │                    │  (stateless) │                    │             │
│ ┌─────────┐ │                    └──────────────┘                    │ ┌─────────┐ │
│ │oxi-wasm │ │    yrs sync protocol (binary)                         │ │oxi-wasm │ │
│ │+collab  │ │◄─────────────────────────────────────────────────────►│ │+collab  │ │
│ └─────────┘ │                                                       │ └─────────┘ │
└─────────────┘                                                       └─────────────┘
```

## New Crates

| Crate | Role | Target |
|-------|------|--------|
| `oxi-collab` | CRDT document model (yrs), sync protocol | wasm32 + native |
| `tools/collab-server/` | Lightweight WebSocket relay (workspace-excluded) | native (tokio) |

## CRDT Data Model (yrs mapping)

```
YDoc
├── YMap("meta")           → title, author, last modified
├── YArray("paragraphs")   → paragraph list
│   └── YMap               → single paragraph
│       ├── "text": YText  → CRDT text with formatting attributes
│       ├── "style": str   → style ID
│       └── "align": str   → alignment
├── YArray("tables")       → tables (Phase 2d)
└── YMap("awareness")      → cursor positions, user info
```

## Core API (oxi-collab)

```rust
CollabDoc::new() -> YDoc
CollabDoc::from_docx(data: &[u8]) -> YDoc
CollabDoc::to_docx(&self) -> Vec<u8>

CollabDoc::create_sync_message() -> Vec<u8>
CollabDoc::apply_update(update: &[u8])
CollabDoc::on_change(callback)

CollabDoc::insert_text(para: usize, offset: usize, text: &str)
CollabDoc::delete_text(para: usize, offset: usize, len: usize)
CollabDoc::format_text(para: usize, range, attrs)
```

## Dependencies

| Crate | Version | Purpose | License |
|-------|---------|---------|---------|
| `yrs` | 0.25 | CRDT engine | MIT |
| `tokio` | 1.x | Async runtime (server) | MIT |
| `tokio-tungstenite` | 0.24 | WebSocket (server) | MIT |

---

## Phase 2a — Real-time Collaboration + Comments

- [ ] `oxi-collab` crate: YDoc ↔ docx round-trip
- [ ] Plain text sync across browsers (paragraph-level)
- [ ] Formatting sync (bold, color, font)
- [ ] Cursor/selection sharing (Awareness protocol)
- [ ] Comment threads with @mentions
- [ ] `tools/collab-server/` WebSocket relay (~80 lines)
- [ ] Web UI: share button, room URL, colored cursors
- [ ] Deploy relay to fly.io free tier

## Phase 2b — AI Assist + Slash Commands

- [ ] Claude API integration (text generation, summarization, translation)
- [ ] `/` slash command palette (table, image, heading, AI prompt)
- [ ] AI-powered proofreading and style suggestions
- [ ] Multi-language translation (JP ↔ EN priority)

## Phase 2c — Version History + E2E Encryption

- [ ] CRDT-based version timeline (who changed what, when)
- [ ] Snapshot & rollback to any point
- [ ] E2E encryption (AES-256-GCM, key exchange via URL fragment)
- [ ] Zero-knowledge server (relay sees only ciphertext)

## Phase 2d — Templates + Markdown

- [ ] Business document templates (invoice, minutes, contract, resume)
- [ ] Japanese business templates (稟議書, 見積書, 納品書)
- [ ] Markdown import/export (.md ↔ docx)
- [ ] Table and image sync in collaboration

## Phase 2e — PWA + Offline

- [ ] Service Worker for offline access
- [ ] PWA manifest (installable)
- [ ] Offline editing → auto-merge on reconnect
- [ ] IndexedDB local document storage

---

## Oxi Differentiators

| Feature | Why only Oxi |
|---------|-------------|
| **Full WASM** | No server-side processing → ultimate privacy |
| **OOXML fidelity** | True Word/Excel/PowerPoint compatibility |
| **Hanko (oxihanko)** | Japanese enterprise killer feature |
| **PDF signatures** | End-to-end digital contract workflow |
| **OSS** | Self-hostable, no vendor lock-in |
| **Local-first** | Data never leaves the browser unless shared |
| **E2E encrypted collab** | Even the relay server can't read documents |

---

# Oxi v3 Roadmap — Platform

## Vision

Oxi v3 evolves from a **document tool** into a **productivity platform** — extensible, cross-device, and enterprise-ready.

## Architecture Overview

```
┌──────────────────────────────────────────────────────────┐
│                    Oxi Platform                          │
│                                                          │
│  ┌──────────┐  ┌──────────┐  ┌──────────┐  ┌─────────┐  │
│  │   Docs   │  │  Cells   │  │  Slides  │  │  Forms  │  │
│  │  (docx)  │  │  (xlsx)  │  │  (pptx)  │  │  (new)  │  │
│  └────┬─────┘  └────┬─────┘  └────┬─────┘  └────┬────┘  │
│       └──────────────┼──────────────┼─────────────┘       │
│              ┌───────▼───────┐                            │
│              │  Workspace    │  ← Unified notebook view   │
│              │  (all-in-one) │                            │
│              └───────┬───────┘                            │
│       ┌──────────────┼──────────────┐                     │
│  ┌────▼────┐   ┌─────▼─────┐  ┌────▼─────┐              │
│  │ Collab  │   │  Plugins  │  │ AI Agent │              │
│  │ (CRDT)  │   │  (WASM)   │  │ (Claude) │              │
│  └─────────┘   └───────────┘  └──────────┘              │
│                                                          │
│  ┌─────────┐   ┌───────────┐  ┌──────────┐              │
│  │ Desktop │   │  Mobile   │  │ SaaS     │              │
│  │ (Tauri) │   │ (Tauri/   │  │ (Docker) │              │
│  │         │   │  Capacitor)│  │          │              │
│  └─────────┘   └───────────┘  └──────────┘              │
└──────────────────────────────────────────────────────────┘
```

---

## Phase 3a — Multi-app Collaboration

- [ ] Real-time collaborative spreadsheet (oxicells + yrs)
- [ ] Real-time formula recalculation across connected clients
- [ ] Collaborative slide editing with live preview
- [ ] Presentation mode with remote control (presenter → audience sync)

## Phase 3b — Unified Workspace

- [ ] Notebook-style workspace: mix docs, cells, slides in one page
- [ ] Drag-and-drop blocks between document types
- [ ] Cross-document references and embeds (embed a chart from xlsx in docx)
- [ ] Workspace-level search across all document types

## Phase 3c — Plugin System

- [ ] WASM Plugin API (third-party extensions)
- [ ] Plugin manifest format and sandboxed execution
- [ ] Plugin marketplace / registry
- [ ] Built-in plugins: diagram editor (Mermaid), code highlighting, math (LaTeX)
- [ ] Custom stamp/seal plugins for oxihanko

## Phase 3d — Form Builder

- [ ] Drag-and-drop form designer (text, checkbox, dropdown, date, file upload)
- [ ] Form → docx/xlsx export (responses as spreadsheet)
- [ ] Conditional logic and validation rules
- [ ] QR code generation for form sharing
- [ ] Japanese business forms: 届出書, 申請書, アンケート

## Phase 3e — Workflow Automation

- [ ] Approval flows (submit → review → approve/reject)
- [ ] Notification system (in-app, email, Slack webhook)
- [ ] Conditional branching (if approved → generate PDF → email)
- [ ] Template-based document generation from form data
- [ ] Hanko stamp insertion in approval flow (oxihanko integration)

## Phase 3f — AI Agent

- [ ] "Create a report from this spreadsheet" → auto-generate docx
- [ ] "Summarize this document in 3 bullet points" → AI summary
- [ ] Multi-document reasoning (compare two contracts, find differences)
- [ ] Voice-to-document (speech → text → formatted docx)
- [ ] AI-powered data analysis in spreadsheets (natural language → formula)

## Phase 3g — Desktop App (Tauri)

- [ ] Native desktop app for Windows, macOS, Linux via Tauri
- [ ] System file associations (.docx, .xlsx, .pptx open with Oxi)
- [ ] Native file system access (no upload needed)
- [ ] System tray with quick document access
- [ ] Auto-update mechanism

## Phase 3h — Mobile App

- [ ] iOS and Android app via Tauri Mobile or Capacitor
- [ ] Touch-optimized editor UI
- [ ] Camera → document scan → OCR → editable docx
- [ ] Offline sync with desktop/web versions
- [ ] Push notifications for collaboration events

## Phase 3i — Self-hosted SaaS

- [ ] Docker one-click deploy (`docker compose up`)
- [ ] Team/organization management (users, roles, permissions)
- [ ] Admin dashboard (usage stats, storage, audit log)
- [ ] SSO integration (SAML, OIDC)
- [ ] S3-compatible storage backend
- [ ] Backup and restore

---

## Version Summary

| Version | Theme | Key Deliverable |
|---------|-------|-----------------|
| **v1** | Foundation | OOXML parse/render/edit in browser (docx, xlsx, pptx, pdf, hanko) |
| **v2** | Collaboration | Real-time co-editing, AI assist, E2E encryption, PWA |
| **v3** | Platform | Plugin ecosystem, desktop/mobile apps, workflow automation, self-hosted SaaS |
| **v4** | Enterprise | Interop, compliance, industry verticals, ecosystem |

---

# Oxi v4 Roadmap — Enterprise

## Vision

Oxi v4 becomes **enterprise-grade** — ready to replace Microsoft 365 and Google Workspace in organizations with strict compliance, industry-specific needs, and complex integrations.

## Phase 4a — Enterprise Interoperability

- [ ] SharePoint / OneDrive connector (read/write via Graph API)
- [ ] Google Drive connector (read/write via Drive API)
- [ ] WebDAV / NextCloud integration
- [ ] CalDAV calendar integration (meeting minutes auto-linked to events)
- [ ] Email integration (attach/send documents directly from Oxi)
- [ ] Import/export: ODF (.odt, .ods, .odp), RTF, legacy .doc/.xls/.ppt

## Phase 4b — Compliance & Governance

- [ ] Audit trail (immutable log of all document operations)
- [ ] Data retention policies (auto-archive, auto-delete)
- [ ] DLP (Data Loss Prevention) — block sharing of documents containing PII
- [ ] GDPR / APPI (個人情報保護法) compliance toolkit
- [ ] Electronic signature legal compliance (eIDAS, 電子署名法)
- [ ] Document classification labels (confidential, internal, public)
- [ ] ISO 27001 / SOC 2 readiness documentation

## Phase 4c — Industry Verticals

### Legal (法務)
- [ ] Contract lifecycle management (draft → negotiate → sign → archive)
- [ ] Redline comparison (diff two document versions with tracked changes)
- [ ] Clause library (reusable legal text blocks)
- [ ] Bates numbering for litigation documents

### Healthcare (医療)
- [ ] HL7 FHIR document generation
- [ ] Patient consent forms with digital signature
- [ ] HIPAA-compliant document handling

### Education (教育)
- [ ] Assignment submission and grading workflow
- [ ] Rubric-based feedback system
- [ ] Student collaboration with teacher oversight
- [ ] LTI integration (Canvas, Moodle, Google Classroom)

### Government (行政)
- [ ] e-Gov 連携 (日本電子政府)
- [ ] マイナンバーカード署名 (JPKI) — oxihanko integration
- [ ] 公文書フォーマット準拠
- [ ] 長期保存形式 (PDF/A) エクスポート

## Phase 4d — Advanced AI

- [ ] RAG (Retrieval-Augmented Generation) over document corpus
- [ ] "Search across all my documents" with semantic understanding
- [ ] Auto-fill forms from existing documents (AI extracts fields)
- [ ] Meeting transcript → formatted minutes (speech-to-doc pipeline)
- [ ] Contract risk analysis (AI highlights risky clauses)
- [ ] Multi-modal: image/chart understanding in documents

## Phase 4e — Developer Ecosystem

- [ ] REST API for headless document operations (parse, convert, merge)
- [ ] CLI tool (`oxi convert report.docx report.pdf`)
- [ ] npm / crates.io packages for embedding Oxi in other apps
- [ ] GitHub Action: auto-convert markdown → docx on PR merge
- [ ] Webhook system for document lifecycle events
- [ ] SDK documentation and developer portal

## Phase 4f — Performance & Scale

- [ ] Streaming parser for 100MB+ documents
- [ ] Web Worker parallelization (parse pages concurrently)
- [ ] Incremental layout (re-render only changed pages)
- [ ] CDN-based WASM distribution (sub-second load)
- [ ] Server-side rendering for SEO / link previews
- [ ] Multi-region relay servers for global collaboration
