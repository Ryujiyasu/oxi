# Changelog

## How versions are numbered

Oxi's version number is the **fidelity floor of its measurement corpus**, not a marketing
milestone: `0.8` means every document in the development corpus renders at **SSIM ≥ 0.80**
against Microsoft Word's own render of the same file. The number moves when the *worst*
document moves, so it cannot be improved by getting good documents better. `1.0` is
therefore the project's stop condition, not a maturity label — it means no document in the
corpus is distinguishable from Word's output.

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
