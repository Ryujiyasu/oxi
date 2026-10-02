# Changelog

## How versions are numbered

Since 0.8.2 the version number is the **penalized mean SSIM against Microsoft Word on the
current frozen blind sets** — fifty English and fifty Japanese documents that nothing has
been tuned against, scored on Word's own 150 DPI render of the same file, with every page
Oxi adds or drops scored zero: `0.8.2` means the two-language penalized mean is 0.82
(0.845 / 0.809, two-decimal truncation). A blind set is spent the moment its
number is written down, so the next release is measured on a new one frozen beforehand.
`1.0` is the project's stop condition, not a maturity label — it means no blind document is
distinguishable from Word's output.

0.8.0 used a different definition (the floor of the tuned development corpus: no document
under 0.80). The change is recorded here so the two numbers are not read as one series.

---

## 0.8.2 — 2026-10-02

Forty days and 1,360 commits after 0.8.0, almost all of them layout: the Word engine now
derives its rules from measured Word behaviour on every blind rotation since (E, F, G) and
from a page-1 line census over a hundred wild documents. Engine `17d1d0a4`.

### Rendering fidelity — blind-H, frozen 2026-10-02 before measurement

Fifty English and fifty Japanese documents that no gate, probe, census or agent had
touched, drawn by the declared SHA-order rule (`pipeline_data/FROZEN_BLIND_H.md`), every
engine rendering the same files, scored against Microsoft Word's own 150 DPI render
(Microsoft 365 16.0.20430.20092). Mean SSIM on common pages / penalized (missing pages
score 0) / page counts equal to Word's.

| Engine | English (50) | Japanese (50) |
|---|---|---|
| ONLYOFFICE 9.3.1.8 | **0.918 / 0.912 / 48** | 0.787 / 0.746 / 39 |
| **Oxi 0.8.2** | 0.874 / 0.845 / 43 | **0.857 / 0.809 / 41** |
| LibreOffice 26.8.0.3 | 0.858 / 0.838 / 42 | 0.807 / 0.714 / 32 |
| Polaris Office 11.115 | 0.827 / 0.792 / 41 (49 scored) | 0.828 / 0.768 / 37 (48 scored) |
| GenOffice 6a241c4 | 0.772 / 0.737 / 39 | 0.800 / 0.696 / 29 (48 scored) |
| SILURUS @silurus/ooxml 0.86.1 | 0.762 / 0.726 / 35 (48 scored) | 0.819 / 0.741 / 34 (47 scored) |
| eigenpal @eigenpal/docx-editor-react 1.9.0 | 0.739 / 0.672 / 30 | 0.761 / 0.654 / 29 (49 scored) |

Head-to-head against LibreOffice, Oxi is closer to Word on 32 of 50 English and 41 of 50
Japanese documents. The worst Oxi document scores 0.579 (English) and 0.551 (Japanese; every
engine is in the 0.50s on that one). BetterOffice and OfficeCLI, the two
remaining engines of the 0.8.0 comparison, were not completed on this rotation (the host ran
out of memory during their browser renders); a blind set is reported once, so their columns
stay empty.

### What changed since 0.8.0

- **Floating objects**: text flows beside page/margin-anchored pictures and into lanes as
  narrow as Word allows; tblpY without vertAnchor is measured from the top margin; floating
  tables break between rows and keep footnotes with their lines
- **Headers and footers**: an ink-less header reserves its line only when its paragraph is
  "touched" (direct formatting or a second paragraph) — the same law the footer already had
- **Ruby**: distributeLetter / distributeSpace placement, wide-ruby base distribution, body
  ruby line height and raise, mixed-size baselines on the character grid
- **Line placement**: atLeast Latin lines seat on their descent; CJK lines seat by their
  largest East Asian run; firstLineChars draws the cached twips; justify shrink allowance
  derived from 7,599 measured Word lines (space 25% / last word 0.35)
- **Tables**: per-cell row splitting, tblPrEx row borders, compat-15 autofit, keepNext
  before a table, CJK row page-bottom rules, grid cells with inline images
- **Symbols and fonts**: w16se symEx runs, Segoe UI Symbol / Emoji advance widths,
  fontTable altName substitution
- **Build**: the four giant layout methods got their own codegen units (crate build halved);
  pre-commit checks are memory-staged (`tools/metrics/precommit.sh`)

### Gates in this release

- Pagination oracle (per-paragraph page match against Word): development corpus **96/96**;
  the declared 1,115-document strict set (every rotated blind set included) **1,115/1,115**;
  blind-G at its measurement 94/100, now 100/100 as a development pool
- Page-1 line census (Blind-G, 98 documents, glyph-dump baselines vs Word PDF): 65 documents
  with no line more than 2pt off, 246 lines off across the other 33 — the next work queue
- SSIM sentinel 369 documents / 1,113 pages; adversarial probes; PPTX render gate; spreadsheet
  oracles; `cargo test` (572) and `cargo clippy` in CI

### Known limits

- The version number is now the blind-set penalized mean, not a floor: the worst blind-H documents
  score 0.579 (English) and 0.551 (Japanese), and the tuned development sentinel (275 real
  documents, mean 0.947) has a real-document floor of 0.658
- ONLYOFFICE remains ahead on English within-page pixels (0.918 vs 0.874)
- Thirty-three Blind-G documents still have page-1 lines more than 2pt from Word's
- The .xlsx and .pptx engines are younger than the .docx one and gated on narrower corpora

### Install

`cargo install oxidocs-cli`, `npm install @oxidocs/wasm`, the `oxidocs-*` crates on crates.io,
and desktop installers from GitHub Releases (the desktop app keeps its own version line, `desktop-v0.8.x`, because its auto-updater only moves forward).

---

## 0.8.0 — 2026-08-24

First tagged release. The floor reached 0.80 on 2026-07-14 and has held since: over the
235 development documents currently scored against stored Word renders, the lowest-scoring
one is at **0.8018** and the mean is 0.9591 per document.

### Rendering fidelity

Measured against Microsoft Office's own renders at 150 DPI, on **blind sets frozen before
measurement and never fixed against** (details in [README](README.md#layout-accuracy-vs-microsoft-word)):

| Blind set | Oxi | best other engine measured |
|---|---|---|
| English, 50 documents | **0.875** mean SSIM, **48/50** page counts match Word | ONLYOFFICE 0.902 / 41 of 50 |
| Japanese, 50 documents | **0.842** mean SSIM, 43/50 page counts match Word | LibreOffice 0.816 / 41 of 50 |
| PowerPoint, 48 decks | **0.953** mean SSIM, 48/48 slide counts match | LibreOffice 0.913 |

Oxi places page breaks better than any engine measured on both Word corpora, leads outright
on Japanese and on PowerPoint, is a statistical tie with LibreOffice on English within-page
pixels, and trails ONLYOFFICE there by 0.027.

### What is in the box

- **.docx** — parser, layout engine and renderer built against Word as ground truth,
  including Japanese typography as a first-class target: JIS X 4051 kinsoku, character
  grid (docGrid), vertical writing with tate-chu-yoko, ruby, warichu, emphasis marks
- **.pptx** — parser, IR and renderer: slide-master placeholder inheritance, group
  transforms, embedded fonts, preset geometry, tables, charts
- **.xlsx** — parser, IR, renderer, and a dependency-graph formula engine (61 functions)
  whose recalculation is diffed against Excel's own cached results across 285 real workbooks
- **Browser VBA host** — workbook macros run client-side: 95 members across 11 host
  objects, each derived from and A/B-verified against real Excel COM behaviour
- **PDF** — parsing, text extraction and generation; hanko (Japanese digital stamps) with
  PAdES signatures
- **Round-trip editing** — .docx / .xlsx / .pptx edits patch only the changed XML text
  nodes inside the original ZIP; a no-edit save is byte-identical, and that is a test
- **Distribution** — WebAssembly bindings and a Canvas editor, a CLI (`oxidocs`) and a
  Tauri desktop app. All processing is client-side; nothing leaves the device

### Gates in this release

- Pagination oracle: per-paragraph page match against Word on the 96-document development
  corpus — **96/96**
- SSIM regression sentinel: 238 documents pixel-compared against stored Word renders
- Adversarial probe harness: 95 synthetic documents gated against real Word output
- PPTX render gate: 40 development decks (886 slides) against PowerPoint's own render
  (mean SSIM 0.957), plus 156 probe decks byte-compared and a determinism check
- Spreadsheet oracles: 285 workbooks recalculated against Excel's cached results; row
  heights agree with Excel on 281 of them
- Golden parse suite: 504 real-world files, 100% parse success
- `cargo test`, `cargo clippy` and the WebAssembly build now run in CI on every push

### Known limits

- The development-corpus floor is 0.8018 — the lowest-scoring document families are
  Latin justified wrap, form-heavy tables, and vector shape groups
- ONLYOFFICE remains ahead on English within-page pixels (0.902 vs 0.875)
- The .xlsx and .pptx layout engines are younger than the .docx one and are gated on
  narrower corpora
- No IME support in the browser editor yet; .odt rendering is not implemented
- Nothing is published to crates.io / npm / PyPI yet — build from source, or use the
  desktop app and the web demo
