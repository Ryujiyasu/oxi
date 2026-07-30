# Why modern font metrics cannot reproduce Word pagination

Internet Explorer 11 changed text layout from legacy GDI-compatible metrics to DirectWrite natural metrics. The result was familiar to anyone who maintained a pixel-tight site: text that had fit inside a 330px box could wrap onto another line. Microsoft documented both the regression and the switch back to GDI-compatible layout.

- [Microsoft: Turn off natural metrics](https://learn.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-it-pro/internet-explorer-11/ie11-deploy-guide/turn-off-natural-metrics)
- [Microsoft: Site text layout is different in Internet Explorer 11](https://support.microsoft.com/en-gb/topic/site-text-layout-is-different-in-internet-explorer-11-ab56f2d8-5d08-8f4b-8d70-4c0f8136da60)
- [Microsoft: Web page layout broken due to natural metrics in IE11](https://learn.microsoft.com/en-us/archive/blogs/asiatech/web-page-layout-broken-issue-due-to-natural-metrics-in-ie11)

Word pagination has the same compatibility problem. A modern text stack preserves fractional design metrics. Parts of Word's layout path first quantize them to integer pixels at 96 DPI, then pass those integers downstream. A renderer can therefore look close to Word while still wrapping a line—or a page—at a different place.

Oxi is a Rust and WebAssembly DOCX engine. To match Word, I had to reproduce not merely “GDI-like” numbers, but which component is rounded, in which unit, in which order, and in which direction.

## How I measured a closed-source layout engine

The formulas below are not guesses derived from screenshots. I generated controlled DOCX matrices, changed one axis at a time, and joined three independent oracles.

### Word COM: structural positions

For paragraph pitch, I created consecutive single-line paragraphs with identical formatting and subtracted their `Paragraph.Range.Information(6)` values. Constant 6 is `wdVerticalPositionRelativeToPage`. `ParagraphFormat.LineSpacing` and `LineSpacingRule` verify the input setting; they are not treated as the rendered height of a Single-spaced line.

Inside a table cell, I walked `Table.Cell(row, col).Range` character by character, collected each character's `Information(6)`, merged nearby Y coordinates into line levels, and took the median distance between levels. This avoids confusing cell margins or total row height with the line box itself.

### Win32 GDI: integer metrics

I created an `HFONT` with the same face, size, and weight, selected it into a DC, and read `tmHeight`, `tmAscent`, `tmDescent`, and `tmExternalLeading` through `GetTextMetricsW`. Width probes used `GetCharWidth32W` and `GetTextExtentPoint32W`; `GetTextFaceW` and `GetGlyphIndicesW` detected silent font fallback.

### Word PDF: render truth

COM coordinates can use surprising reference frames in complex tables. I therefore exported each probe with Word's `ExportAsFixedFormat` and extracted span baselines, bounding boxes, and page numbers with PyMuPDF. COM is the structural oracle; the PDF is the final render oracle.

The generator expands font, size, character, body/cell context, and document-grid settings from a manifest. Results are stored as JSON or TSV and joined against the Rust predictions. The width study cited below—13 font/size configurations and 181 characters—is the output of that matrix, not a hand-picked sample.

## Quantize the font size to ppem first

At 96 DPI, the first quantization is:

```text
ppem = round(font_size_pt × 96 / 72)
```

In Rust:

```rust
fn pixel_round(value_normalized: f32, ppem: f32) -> f32 {
    (value_normalized * ppem).round()
}

let ppem = (font_size * 96.0 / 72.0).round();
```

A 12pt font becomes `round(12 × 96 / 72) = 16` ppem. The important part is what happens next: rounding a sum is not equivalent to rounding its components and then adding them.

## Latin line height is integerized component by component

Conceptually, the natural line height is the larger of the OpenType `hhea` total and the `OS/2` Windows-metrics total:

```text
max(
    hhea_ascent + hhea_descent + hhea_lineGap,
    winAscent + winDescent
)
```

But evaluating that `max()` in floating point and rounding once did not match Word. The measured GDI-compatible path was:

```text
ppem        = round(fontSize × 96 / 72)
ascent_px   = round(winAscent × ppem)
descent_px  = round(winDescent × ppem)
hhea_excess = max(0, hhea_total − win_total)
leading_px  = round(hhea_excess × ppem)

line_height =
    (ascent_px + descent_px + leading_px) × 72 / 96
```

```rust
let ppem = (font_size * 96.0 / 72.0).round();
let font_ascent = pixel_round(self.win_ascent, ppem);
let font_descent = pixel_round(self.win_descent, ppem);

let win_total = self.win_ascent + self.win_descent;
let hhea_total = self.ascent + self.descent + self.line_gap;
let hhea_excess = (hhea_total - win_total).max(0.0);
let extra_leading = pixel_round(hhea_excess, ppem);

let height_pt =
    (font_ascent + font_descent + extra_leading) * 72.0 / 96.0;
```

Before rounding,

```text
win_total + max(0, hhea_total - win_total)
    = max(win_total, hhea_total)
```

After component-wise integerization, the two expressions are no longer interchangeable. The conceptual model is `max`; the compatibility implementation is independently quantized components. And `72/96 = 0.75` is not a tuning constant—it converts 96-DPI pixels back to points.

## A table cell floors only the descender

Word does not use the same quantization in every context. In one table-cell path, the descender is truncated rather than rounded:

```rust
pub fn word_line_height_table_cell(&self, font_size: f32) -> f32 {
    let ppem = (font_size * 96.0 / 72.0).round();
    let font_ascent = pixel_round(self.win_ascent, ppem);
    let font_descent = (self.win_descent * ppem).floor();
    (font_ascent + font_descent) * 72.0 / 96.0
}
```

Times New Roman 10pt is a real branch point. The font I measured has UPM=2048, `winAscent=1825`, and `winDescent=443`:

```text
ppem        = round(10 × 96 / 72) = 13
ascent_raw  = 1825 / 2048 × 13 = 11.5845...
descent_raw =  443 / 2048 × 13 =  2.8120...

ascent_px   = round(11.5845...) = 12
descent_px  = floor( 2.8120...) =  2
cell        = (12 + 2) × 72 / 96 = 10.5pt
```

Rounding the descender produces 3px and an 11.25pt cell line: a 0.75pt error per line.

Calibri 10.5pt also branches with the installed Version 6.27 font. Its `OS/2` values are UPM=2048, `winAscent=1950`, and `winDescent=550`:

```text
ppem        = round(10.5 × 96 / 72) = 14
descent_raw = 550 / 2048 × 14 = 3.7598...
round       = 4
floor       = 3
```

Calibri 11pt does not branch: `550 / 2048 × 15 = 4.0283...`, so both operations produce 4. The often-quoted Calibri values 1536/512 are `hhea` ascent/descent, not `OS/2` `winAscent/winDescent`. A reproducible claim must identify the font version and the table.

## Where 83/64 came from

MS Gothic and MS Mincho required another line-spacing path:

```text
raw = (winAscent + winDescent) / UPM × fontSize × 83 / 64
line_height = floor(raw × 8) / 8
```

`83/64` is neither a published Microsoft constant nor a field copied from `OS/2` or `hhea`. It is the rational representation I derived from Word measurements.

I removed document grids, fixed spacing, and paragraph margins; swept the font size in 0.5pt steps; measured adjacent paragraph Y deltas through COM; isolated the slope because MS Gothic and MS Mincho have `winAscent + winDescent = UPM = 256`; incorporated the observed 1/8pt floor; then searched coefficients with power-of-two denominators that satisfied all measured quantization intervals.

The selected coefficient was:

```text
1.296875 = 83 / 64
```

| Size | `fontSize × 83/64` | 1/8pt floor |
|---:|---:|---:|
| 10.5pt | 13.6171875pt | 13.5pt |
| 12pt | 15.5625pt | 15.5pt |
| 14pt | 18.15625pt | 18.125pt |

Replacing it with `1.3` changes the 1/8pt result at 107 of the 129 half-point sizes from 8pt through 72pt.

The more revealing test is a slightly low approximation. `1.2968` changes exactly nine sizes:

```text
8, 16, 24, 32, 40, 48, 56, 64, 72pt
```

Why those sizes?

```text
fontSize × 83/64 × 8 = fontSize × 83/8
```

On the half-point grid, this is an integer exactly when the font size is a multiple of 8pt. Those nine inputs sit on the knife edge immediately before `floor()`. Move the coefficient slightly downward and every one falls by 1/8pt.

The fraction is also exactly representable in binary:

```text
83/64 = 1.296875₁₀ = 1.010011₂
```

Both the coefficient and half-point font sizes are dyadic rationals, so this sweep's `fontSize × 83/64 × 8` is exact in `f32`. Decimal `1.3` repeats in binary. Keeping the measured slope as a dyadic rational is therefore not cosmetic: it makes the quantization boundary stable in the implementation.

I do not claim Microsoft stores the literal fraction `83/64`. I claim that the Oxi model is reproducible from the sweep, the observed 1/8pt quantization, and a dyadic-rational search.

Nor does `units_per_em == 256` select this path by itself. Proportional CJK fonts and families with different natural heights require measured family classification.

## Character widths use another quantization

The 13-configuration, 181-character Word sweep matched a 0.5pt—or 10-twip—layout-width quantization for known Latin metrics:

```rust
let advance_em = metrics.char_width_em(c);

// Positive layout widths: round-half-up to 10 twips.
let width_tw =
    (advance_em * font_size * 20.0 / 10.0 + 0.5).floor() * 10.0;
let width_pt = width_tw / 20.0;
```

For positive inputs this spells out round-half-up. A named helper would communicate the policy better; replacing it mechanically with a generic rounding function would hide the reason.

Some hinted GDI metrics cannot be reconstructed from OpenType ratios and ppem alone. Oxi stores those measurements by font and ppem instead of accumulating font-name special cases:

```rust
// font -> ppem -> (height, ascent, descent)
gdi_heights: HashMap<String, HashMap<u32, (u32, u32, u32)>>,

// font -> ppem -> codepoint -> width_px
gdi_widths: HashMap<String, HashMap<u32, HashMap<u32, u32>>>,
```

This is intentionally plain storage; typed or flat keys would be cleaner. The important separation is between a semantic layout rule and host-measured data.

Arial Narrow 10pt demonstrates why the table exists. At 13 ppem, `GetTextMetricsW` returns:

```text
tmHeight  = 16px
tmAscent  = 13px
tmDescent = 3px
```

That is 12pt after multiplying by 72/96. Word's 1.15 line spacing becomes `12 × 1.15 = 13.8pt`, matching COM. A direct `hhea` calculation does not.

## 96-DPI layout versus 150-DPI evaluation

The 96 DPI in these formulas is the compatibility quantization that decides wrapping and pagination. Oxi's benchmark images are rasterized at 150 DPI. That is an evaluation resolution after layout, not a replacement for 96 in the layout formula.

## Is `f32` safe at a rounding boundary?

Oxi's IR and layout coordinates use `f32`. The 83/64 path is a case where safety is concrete: the coefficient and half-point inputs are dyadic, so even the 8pt-multiple integer boundaries remain exact.

That does not make arbitrary decimal sizes and coefficients safe. Where tie behavior is itself the specification, retaining integer OpenType metrics and using a `MulDiv`-style operation is stronger. Rust's `f32::round()` rounds halfway cases away from zero; if Word uses another rule in a context, the function name and boundary tests must say so.

## Four rules that survived the implementation

1. Do not add components before rounding if Word/GDI integerizes them separately.
2. Separate unit conversions such as 72/96 and 20 twips/pt from coefficients derived from measurements.
3. Treat `round`, `floor`, `ceil`, and half-up as contextual layout rules, not interchangeable approximations.
4. When hinting cannot be expressed by a stable formula, use a narrowly scoped, reproducible measurement table.

As of July 29, 2026, Oxi's frozen 50-document blind sets score 0.828 SSIM with Word and 44/50 page-count matches in Japanese, and 0.825 with 48/50 page-count matches in English. English page-count fidelity is the highest among the engines measured, although Oxi still trails mature native suites in within-page pixel placement.

The English and Japanese sets contain different documents, and page-count equality is a discrete metric while SSIM also measures glyph placement and rasterization. The higher English page-count result alongside lower SSIM is therefore not a contradiction or a direct ranking between languages.

Eliminate enough integer-pixel errors, and eventually a whole page of error disappears.
