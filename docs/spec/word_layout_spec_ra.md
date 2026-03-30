# Word Layout Specification (Ra + Manual Measurement Integrated Edition)

Clarified through COM API black-box measurements. No DLL analysis.

---

## 1. Line Height (line_height)

### 1.1 Basic Calculation (lineSpacingRule = auto/Single)

```
fn gdi_line_height(font_metrics, font_size_pt) -> f32:
    ppem = round(font_size_pt * 96.0 / 72.0)

    // Method 1: GDI tmHeight table lookup (most accurate)
    if GDI_HEIGHT_TABLE.has(font, ppem):
        return GDI_HEIGHT_TABLE[font][ppem] * 72.0 / 96.0

    // Method 2: round formula (accurate for most fonts)
    asc_px = round(win_ascent * ppem / upm)  // round, NOT ceil!
    des_px = round(win_descent * ppem / upm)
    return (asc_px + des_px) * 72.0 / 96.0

    // Method 3: UPM=256 fonts (MS Gothic/Mincho family)
    // gdi_h = ppem (direct)
```

**GDI measurement confirmed (2026-03-29):**

Formula comparison (ppem 7-49, all fonts):

| Formula | Calibri | Cambria | Meiryo | MS Gothic | Arial | TNR | Century |
|------|---------|---------|--------|-----------|-------|-----|---------|
| **round+round** | **0 err** | **0** | **0** | **1** | 16 | 23 | 29 |
| ceil+ceil (old) | 69 err | 70 | 70 | 93 | 60 | 69 | 73 |

- **round(winAsc*ppem/upm)+round(winDes*ppem/upm)** is the correct formula
- Calibri/Cambria/Meiryo: **exact match**
- MS Gothic/Mincho (UPM=256): **gdi_h = ppem** (wa+wd=upm so round=ppem)
- Arial/TNR/Century: some mismatches even with round formula -> GDI table lookup required
  - Cause: TrueType bytecode hinting adjusts tmAscent independently per ppem
  - Arial ppem=11-13: matches neither ceil/round/floor (hinting correction)
- **GDI tmHeight table** (`gdi_height_table.json`, 54.6KB) provides full coverage for all fonts

**3-tier approach:**
1. **GDI table lookup** (highest priority) -- always accurate. All 21 fonts x ppem 5-100
2. **round formula** (for fonts not in table) -- 0 errors for Calibri/Cambria/Meiryo family
3. **ppem direct** (UPM=256) -- MS Gothic/Mincho family

**Font classification:**
- round formula OK (0 errors): **Calibri, Calibri Bold, Cambria, Cambria Bold, Meiryo, Yu Mincho Demibold** (6 entries)
- ppem direct (1 error): **MS Gothic, MS PGothic, MS Mincho, MS PMincho** (UPM=256)
- Table required: **Arial, Arial Bold, TNR, TNR Bold, Century, Yu Gothic Reg/Bold, Yu Mincho Reg** (8 entries)
  - Yu Gothic/Yu Mincho Regular have hinting so round=39 errors -> table required

### 1.2 CJK 83/64 Multiplier

**Application condition:** Font family is in the CJK whitelist

Whitelist (COM measurement 2026-03-29):
  MS Gothic, MS Mincho, MS PGothic, MS PMincho,
  Yu Gothic, Yu Mincho, Meiryo
  **HG-series fonts are NOT applied** (HGGothicM ratio=1.125 != 1.297)

```
fn word_line_height(font_metrics, font_size_pt) -> f32:
    base = gdi_line_height(font_metrics, font_size_pt)
    if is_cjk_font(font_family):
        return base * 83.0 / 64.0
    else:
        return base
```

**COM measurement confirmed:**
- MS Gothic 10.5pt noGrid: 14.25-15.00pt (average approx. 14.6pt, CJK 83/64 applied)
- Calibri 10.5pt noGrid: 13.5pt
- MS Gothic 10.5pt grid(18pt): 18.0pt (1x grid)
- Meiryo 10.5pt grid(18pt): 36.0pt (2x grid, natural height is large)

### 1.3 By lineSpacingRule

```
fn line_height(rule, value, font_metrics, font_size, grid_pitch) -> f32:
    match rule:
        "auto" | "single":
            base = word_line_height(font_metrics, font_size)
            factor = value / 240.0  // w:line="240" = 1.0x
            lh = base * factor
            // COM measurement confirmed (2026-03-29, Calibri 11pt gdi_h=13.5pt, noGrid):
            //   With XML w:line setting: gap = gdi_h x factor
            //     Single(240): 13.5x1.0 = 13.5pt ✓
            //     1.15(276):   13.5x1.15 = 15.5pt ✓
            //     1.5(360):    13.5x1.5  = 20.0pt ✓
            //     Double(480): 13.5x2.0  = 27.0pt ✓
            //   No difference between compat=14 and 15 (previous discrepancy was a style inheritance issue)
        "exact":
            lh = value  // w:line in twips / 20
            // grid snap is NOT applied (COM confirmed)
            // COM measurement (2026-03-29):
            // Value is used directly as twips/20 pt (no rounding)
            // Word pixel-snaps Y coordinates (0.5pt quantization), so
            // non-integer pt values cause alternating micro-variations in line spacing (e.g., 9.15pt -> 9.0,9.5,9.0,9.0)
            // However the average matches the expected value (183tw avg=9.125 approx. 9.15, error 0.025pt)
            // Implementation: use twips/20 pt value as-is. Pixel-snap at render time
            return lh
        "atLeast":
            natural = word_line_height(font_metrics, font_size)
            // grid snap is applied to natural, but NOT to specified
            if grid_pitch > 0:
                natural = grid_snap(natural, grid_pitch)
            lh = max(natural, value)
            return lh  // already grid-snapped or specified (no snap needed)

    // grid snap (only for type="lines" or "linesAndChars")
    if grid_pitch > 0:
        lh = grid_snap(lh, grid_pitch)

    return lh
```

### 1.4 Grid Snap

```
fn grid_snap(lh, pitch) -> f32:
    // round-half-up method (confirmed in COM Session 2)
    n = floor((lh + pitch / 2.0) / pitch + 0.5)  // equiv. round_half_up(lh / pitch)
    if n < 1: n = 1
    return n * pitch
```

**Application conditions:**
- docGrid type="lines" -> applied
- docGrid type="linesAndChars" -> applied
- docGrid type absent (linePitch only) -> applied (COM measurement: confirmed with gen_027)
- docGrid element absent -> NOT applied
- exact lineSpacingRule -> NOT applied

### 1.5 Inside Table Cells

- adjustLineHeightInTable=false (default): CJK 83/64 enabled
- adjustLineHeightInTable=true: CJK 83/64 disabled
- **grid snap: depends on compat_mode** (see section 15.3)
  - compat=15 (Word 2013+): grid snap **enabled**
  - compat=14 (Word 2010): grid snap **disabled**
- **lineSpacing/spaceAfter/spaceBefore: no automatic reset**
  - COM confirmed (2026-03-29): sa/sb/ls from docDefaults/Normal style are preserved inside table cells and TextBoxes
  - Behavior previously seen as "reset" was actually the table style overriding Normal's spacing
  - Paragraph spacing inside cells depends on the style inheritance chain (there is no cell-specific automatic reset feature)

### 1.6 Inside TextBoxes

- **grid snap enabled** (in compat=15, grid snap is applied the same as regular text)
  - COM measurement: MS Gothic 10.5pt inside TB gap=18pt = grid snap(17.85pt -> 18pt) ✓
  - Old spec "no grid snap" was incorrect (corrected 2026-03-29)
- CJK 83/64 is applied
- **spacing reset: none (follows style inheritance chain, same as table cells)** (COM confirmed 2026-03-29)

### 1.7 Mixed Font Run (CJK + Latin mixed line)

```
fn mixed_line_height(runs, grid_pitch) -> f32:
    lh = max(word_line_height(run.font_metrics, run.font_size) for run in runs)
    if grid_pitch > 0:
        lh = grid_snap(lh, grid_pitch)
    return lh
```

**COM measurement confirmed (2026-03-29):**
- Latin(Calibri 10.5pt)=14.5pt, CJK(MS Gothic 10.5pt)=15.5pt -> mixed=15.5pt = max ✓
- grid snap is applied after max (max -> snap, NOT snap each -> max)
- +/-0.5pt fluctuation in Y positions (COM measurement precision limit, not a spec issue)

---

## 2. Paragraph Spacing (spacing)

### 2.1 Basic Rules

```
fn paragraph_gap(prev_para, next_para) -> f32:
    sa = prev_para.space_after
    sb = next_para.space_before
    spacing = max(sa, sb)  // Collapse: use the larger value
    return prev_line_height + spacing
```

**COM measurement confirmed:**
- sa=10, sb=24 -> spacing=24pt (max) ✓
- sa=24, sb=10 -> spacing=24pt (max) ✓
- sa=10, sb=10 -> spacing=9.75pt <- grid snap effect
- sa=15, sb=15 -> spacing=15pt ✓
- sa=0, sb=24 -> spacing=24pt ✓

**spaceAfter precision (2026-03-29, noGrid):**
gap = line_height + sa (exact match, twips/20 added directly)
- sa=0 -> gap=13.5, sa=60tw(3pt) -> 16.5, sa=120tw(6pt) -> 19.5, sa=240tw(12pt) -> 25.5

**Note:** grid snap may also be applied to spacing.
sa=sb=10pt -> 10pt -> grid snap -> 9.75pt (= 13px * 0.75)

### 2.2 spaceBefore Suppression at Page Top

```
fn effective_space_before(para, is_first_on_page) -> f32:
    if is_first_on_page:
        return 0.0  // Completely suppressed
    else:
        return para.space_before
```

**COM measurement confirmed:**
- First paragraph on page 2: sb=6pt -> y=74.25 (same as without sb)
- Second paragraph on page 2: sb=12pt -> applied normally

### 2.3 contextualSpacing

```
fn apply_contextual_spacing(prev_para, next_para) -> (f32, f32):
    // If either has contextualSpacing=True AND same style, set both sa/sb to 0
    if (prev_para.contextual_spacing || next_para.contextual_spacing)
       && prev_para.style == next_para.style:
        return (0.0, 0.0)  // (effective_sa, effective_sb)
    else:
        return (prev_para.space_after, next_para.space_before)
```

**COM measurement confirmed (2026-03-29):**
- Same style + ctx=ON: gap=15.5pt (line height only, sa/sb completely suppressed)
- Same style + ctx=OFF: gap=27.5pt (line height + spacing)
- Different style + ctx=ON: gap=27.5pt (no effect)
- **Even if only one side has ctx=ON, suppression occurs** (even if only P1 is True, spacing between P1-P2 = 0)
- asymmetric (sa=20, sb=10) ctx=ON: gap=15.5pt -> both set to 0 regardless of values

### 2.4 beforeLines / afterLines

```
fn before_lines_to_pt(value, line_pitch) -> f32:
    return value / 100.0 * line_pitch
    // Result gets grid snap applied (if snap_to_grid=true)
```

---

## 3. Grid Snap (grid_snap)

### 3.1 Application Conditions

| docGrid | type | Line height snap | spacing snap |
|---------|------|-----------|-------------|
| type="lines" | present | ✓ | ✓ |
| type="linesAndChars" | present | ✓ | ✓ |
| linePitch only (no type) | present | ✓ | ✓ |
| No docGrid element | absent | ✗ | ✗ |
| snap_to_grid=false | -- | ✗ | ✗ |

### 3.2 Rounding Method

```
fn grid_snap(value, pitch) -> f32:
    // round-half-up (2.5 -> 3)
    n = floor(value / pitch + 0.5)
    if n < 1: n = 1
    return n * pitch
```

### 3.3 Y Position of First Paragraph

- Use margin_top directly (no grid offset)
- However, since line height itself is grid-snapped, alignment occurs as a result

---

## 4. Character Width (char_width)

### 4.1 Basic Calculation

```
fn char_width_pt(font_metrics, char, font_size) -> f32:
    ppem = round(font_size * 96.0 / 72.0)
    advance = font_metrics.advance_width(char)  // in UPM units
    pixel_width = round(advance * ppem / upm)
    return pixel_width * 72.0 / 96.0
```

**GDI measurement vs Oxi calculation discrepancies (2026-03-29):**

| Font | ppem | Mismatches/63 chars | Cause |
|---------|------|----------------|------|
| Calibri | 12 (9pt) | **17** | GDI hinting |
| Calibri | 14 (10.5pt) | **0** | Exact match |
| Calibri | 15 (11pt) | **16** | GDI hinting |
| Arial | 12-15 | **11-18** | GDI hinting |
| MS Gothic/Mincho | all sizes | **0** | No rounding needed at UPM=256 |

**Note:** `round(advance * ppem / upm)` is an approximation when GDI hinting is not applied.
TrueType fonts (UPM=2048) can have up to 1px difference because GDI adjusts widths via hinting instructions.
This difference is the primary cause of line-wrap position -> line count -> cumulative Y-coordinate drift.
**Solution: GDI width override tables**
- `gdi_pixel_overrides.json` (14.8KB): only 1888 entries where Oxi calculation differs
- `gdi_width_overrides.json` (1055KB): complete table for all fonts
- 9 fonts x ppem 7-20 x 894 characters measured on Windows GDI
- Bold fonts have more discrepancies than Regular (Arial Bold: 500 entries)

### 4.6 Line Wrap判定 (Line Break Decision)

```
fn needs_line_break(accumulated_width_px, content_width_px) -> bool:
    return accumulated_width_px > content_width_px
    // Note: > not >= (does NOT wrap when exactly equal to content_width)
```

**GDI measurement confirmed (2026-03-29):**
- Calibri 11pt 'A'(9px) x 86 = 774px... x 87=783px: content=602px
  - n=86 string_w=602px -> **1 line** (width == content -> no wrap)
  - n=87 string_w=609px -> **2 lines** (width > content -> wraps)
- `GetTextExtentPoint32W(full string)` = sum of individual character widths
  - string_width == sum(char_widths) (same in Word)
- Mixed text: line 1 gdi_w=459.75pt > content(451.3pt) -> correctly wraps

### 4.2 Font Fallback

**When a Latin font is specified for CJK characters, GDI automatically falls back.**

```
fn resolve_char_width(font_name, char, font_size) -> f32:
    if is_cjk_char(char) && is_latin_font(font_name):
        // GDI fallback: uses MS UI Gothic
        return char_width_pt(ms_ui_gothic_metrics, char, font_size)
    else:
        return char_width_pt(font_metrics, char, font_size)
```

**COM/GDI measurement confirmed (ppem=14):**

| Character | Calibri | MS UI Gothic | MS Gothic |
|------|---------|-------------|-----------|
| あ (U+3042) | 11px | 11px | 14px |
| 一 (U+4E00) | 14px | 14px | 14px |
| A (U+0041) | 8px | 9px | 7px |

Calibri with "あ" = 11px = MS UI Gothic with "あ" = 11px -> **fallback target is MS UI Gothic**

### 4.3 Character Spacing (w:spacing w:val)

```
fn apply_cs(cs_twips) -> f32:
    // GDI MulDiv rounding
    cs_px = MulDiv(cs_twips, 96, 1440)  // = (cs_twips * 96 + 720) / 1440
    return cs_px * 72.0 / 96.0
```

**COM measurement confirmed (2026-03-29, MS Gothic 9pt):**

| cs(tw) | GDI px | GDI pt | Measured CJK gap | base gap |
|--------|--------|--------|-----------|----------|
| 0 | 0 | 0 | 9.0pt | 9.0pt |
| -9 | -2 | -1.5 | 8.5pt | -0.5pt |
| 9 | 1 | 0.75 | 9.5pt | +0.5pt |
| 20 | 1 | 0.75 | 10.0pt | +1.0pt |
| -20 | -2 | -1.5 | 8.0pt | -1.0pt |

**Note:** COM coordinates are affected by 0.5pt quantization. Implementation should use the GDI MulDiv calculated values directly.

### 4.4 Monospaced CJK Fonts (UPM=256)

MS Gothic, MS Mincho:

```
fn cjk_fullwidth_px(font_size_pt) -> i32:
    ppem = round(font_size_pt * 96.0 / 72.0)
    // Round up to even pixels (GDI measurement 2026-03-29)
    return (ppem + 1) & !1  // ceil to even
    // Half-width = fullwidth / 2
```

**GDI measurement confirmed (2026-03-29):**

| fontSize | ppem | CJK fullwidth px | Formula |
|----------|------|----------|------|
| 7pt | 9 | 10 | ceil_even(9)=10 ✓ |
| 8pt | 11 | 12 | ceil_even(11)=12 ✓ |
| 9pt | 12 | 12 | 12(even) ✓ |
| 10pt | 13 | 14 | ceil_even(13)=14 ✓ |
| 10.5pt | 14 | 14 | 14(even) ✓ |

**Note:** `ceil_even` is only for MS Gothic/MS Mincho (UPM=256 bitmap monospaced fonts).

### 4.5 Fullwidth Width of Other CJK Fonts

Yu Gothic, Yu Mincho, Meiryo: **fullwidth = ppem** (no even rounding)

```
fn cjk_fullwidth_other(font_size_pt) -> i32:
    return round(font_size_pt * 96.0 / 72.0)  // ppem direct
```

MS PGothic, MS PMincho: **proportional** (character widths vary per character via GDI)

**GDI measurement confirmed (2026-03-29, all patterns ppem=5-20 verified):**
- MS Gothic/Mincho: ceil_even ALL MATCH (ppem 5-29)
- Yu Gothic/Mincho/Meiryo: CJK fullwidth = ppem (all sizes match)
- MS PGothic/PMincho: proportional ("あ" != ppem, individual GDI width calculation required)

---

## 5. Page Break (page_break)

### 5.1 Basic Condition

```
fn needs_page_break(cursor_y, line_height, page_bottom) -> bool:
    return cursor_y + line_height > page_bottom
```

### 5.2 Widow/Orphan Control

- WidowControl=True: If only the first line of a paragraph remains at the page bottom, move the entire paragraph to the next page
- WidowControl=True: If only the last line of a paragraph goes to the next page, move one additional line from the previous page

**COM confirmed:** With WidowControl=True, a 3-line paragraph completely moved from page 1 (y=740) to page 2 (y=74)

### 5.3 keepWithNext / keepTogether

- keepWithNext: Keep this paragraph and the next paragraph on the same page
- keepTogether: Keep all lines within a paragraph on the same page

**COM confirmed:** With KeepTogether=True, a long paragraph moved to page 2

### 5.4 Table Row Splitting

- AllowBreakAcrossPages = True (default): Row is split across pages
- AllowBreakAcrossPages = False: Entire row moves to the next page

### 5.5 spaceBeforeAutoSpacing

spaceBefore for the first paragraph on a page is **completely suppressed** (treated as 0pt).

---

## 6. Tab Stops (tab_stops)

### 6.1 Default Tab

```
fn default_tab_interval() -> f32:
    // Document.DefaultTabStop (typically 36pt = 0.5 inch)
    return doc.default_tab_stop  // twips / 20
```

**COM measurement confirmed (2026-03-29):**
- DefaultTabStop varies per document
- ja_gov_template.docx: 36pt
- Normal.dotm (Japanese Word): 42pt (= 4 chars x 10.5pt)
- Value is obtained from `w:settings/w:defaultTabStop w:val` (twips) or Document.DefaultTabStop

### 6.2 Tab Position Reference

Tab positions are **absolute positions from the left margin** (twips converted to pt).

```
fn tab_position_pt(tab_pos_twips, margin_left_pt) -> f32:
    // Position is distance from left margin
    return margin_left_pt + tab_pos_twips / 20.0
```

### 6.3 Placement by Tab Type

```
fn apply_tab(tab_type, tab_pos_pt, text_before_width, text_after) -> f32:
    match tab_type:
        "left":
            // Text start position = tab position
            return tab_pos_pt
        "center":
            // Text center = tab position
            return tab_pos_pt - text_after_width / 2.0
        "right":
            // Text right edge = tab position
            return tab_pos_pt - text_after_width
        "decimal":
            // Decimal point position = tab position
            return tab_pos_pt - width_before_decimal_point
```

**COM measurement confirmed (2026-03-29):**
- Left tab @72pt (1440tw): text starts at margin+72pt ✓
- Center tab @216pt (4320tw): "Center"(~30pt) -> x=201pt, center approx. 216pt ✓
- Right tab @432pt (8640tw): "Right"(~23pt) -> x=409pt, right approx. 432pt ✓
- Decimal tab @216pt (4320tw): "123.45" -> decimal point at approx. 216pt ✓

### 6.4 Default Tab Application

When no custom tabs are defined, automatic tab positions are generated at DefaultTabStop intervals.

```
fn next_tab_position(current_x, margin_left, default_interval) -> f32:
    // First default tab position beyond current position
    offset = current_x - margin_left
    n = floor(offset / default_interval) + 1
    return margin_left + n * default_interval
```

### 6.5 Interaction Between Tabs and Indents

**Custom tab positions are absolute positions from the margin and are NOT affected by indent.**

```
fn next_custom_tab(current_x_from_margin, tab_stops, effective_indent) -> Option<f32>:
    // Skip tab positions before indent
    for tab in tab_stops:
        if tab.position > current_x_from_margin && tab.position >= effective_indent:
            return Some(tab.position)  // Absolute position from margin
    return None
```

**COM measurement confirmed (2026-03-29, fully verified with Selection.Information(5)):**

| Setting | Seg0(text start) | Seg1(tab1) | Seg2(tab2) |
|------|-----------------|-----------|-----------|
| indent=0, tab@144,288 | margin+0 | margin+144 | margin+288 |
| indent=36, tab@144,288 | margin+36 | margin+144 | margin+288 |
| indent=72, tab@144,288 | margin+72 | margin+144 | margin+288 |
| indent=180, tab@144,288 | margin+180 | margin+288 | margin+336 |
| hanging=36/indent=72, P1 | margin+36 | margin+72* | margin+144 |
| hanging=36/indent=72, P2 | margin+72 | margin+144 | margin+288 |

\* P1's Seg1(margin+72) is an **implicit tab** at the leftIndent position (auto-generated by hanging indent)
| firstLine=36, P1 | margin+36 | margin+144 | margin+288 |
| firstLine=36, P2 | margin+0 | margin+144 | margin+288 |

- Text start position = margin + effective_indent (exact match, 0.0pt error)
- indent=180, tab@144: **tab@144 < indent -> skipped**, tab@288 used
- hanging indent: P1 effective=36(72-36), P2 effective=72

### 6.6 Tab Leaders

- `dot`: dotted leader
- `hyphen`: dash leader
- `underscore`: underscore leader
- Leader characters fill the tab space (used in table of contents, etc.)

---

## 7. Multi-Column Layout (columns)

### 7.1 Column Position Calculation

```
fn column_x_positions(margin_left, columns) -> Vec<f32>:
    x = margin_left
    positions = []
    for i, col in columns:
        positions.push(x)
        x += col.width + col.space_after
    return positions
```

**COM measurement confirmed (2026-03-29):**

| Setting | Col1 x | Col2 x | Col3 x |
|------|--------|--------|--------|
| 2col equal (w=215, sp=21.25) | 72.0 | 308.5 (approx. 308.25) | - |
| 3col equal (w=136.25, sp=21.25) | 72.0 | 229.5 | 387.0 |
| 2col gap=36 (w=207.65) | 72.0 | 315.5 (approx. 315.65) | - |
| 2col unequal (w1=150, w2=265.3, sp=36) | 72.0 | 258.0 | - |

### 7.2 Equal Width Column Calculation

```
fn equal_column_width(text_width, num_cols, spacing) -> f32:
    // text_width = page_width - margin_left - margin_right
    return (text_width - spacing * (num_cols - 1)) / num_cols
```

### 7.3 Text Flow

- Text flows in order: Column 1 -> Column 2 -> ... -> next page Column 1
- Column height is normally the same as the page body area (top_margin ~ bottom_margin)
- Column Break (wdColumnBreak): forcibly moves to the next column

### 7.4 Column Y Coordinate

- Y start position for each column is the same, starting from the page top margin
- Text is laid out independently from the page top into each column

### 7.5 Mid-Paragraph Column Break

**When paragraph lines exceed the column bottom, remaining lines continue from the top (=start_y) of the next column.**

```
fn column_line_overflow(cursor_y, line_height, col_bottom, next_col_start_y) -> f32:
    if cursor_y + line_height > col_bottom:
        return next_col_start_y  // Y start position of next column
    return cursor_y
```

**COM measurement confirmed (2026-03-29):**
- After filling column 1 with 35 short paragraphs, a long paragraph (16 lines):
  - 3 lines in column 1 (y=704.5, 722.5, 740.5)
  - 13 lines in column 2 (y=74.5 onwards, col2 x=308.5)
  - **Y coordinate resets to top_margin (74.5)**
- keepTogether=True: entire paragraph moves to column 2 top (x=308.5, y=74.5)
- **Same logic as within-page line splitting**

---

## 12. Numbered Lists (numbering)

### 12.1 Basic Layout

```
fn list_paragraph_layout():
    // Numbered list = hanging indent + list marker
    // leftIndent = text start position (from margin)
    // firstLineIndent = -leftIndent (hanging: width of marker area)
    // Marker is placed between margin+0 and margin+leftIndent
```

**COM measurement confirmed (2026-03-29):**
- Basic numbered list (1.2.3.): leftIndent=22pt, firstLineIndent=-22pt
- Text start position = margin+22pt (= leftIndent)
- Bullet list: same (li=22, fli=-22), marker=U+F06C (Wingdings)

### 12.2 Indent by Nesting Level

| Level | leftIndent | firstLineIndent | Text start (margin-relative) |
|-------|-----------|-----------------|----------------------|
| 1 | 22.0pt | -22.0pt | 22.0pt |
| 2 | 32.5pt | -22.0pt | 32.5pt |
| 3 | 43.0pt | -22.0pt | 43.0pt |

- Indent increment between levels: +10.5pt
- firstLineIndent is the same across all levels (-22pt)

### 12.3 Interaction Between Lists and Custom Tabs

- Text within lists starts from the leftIndent position
- Custom tabs (@144pt) are at absolute margin positions (not affected by list indent)
- Tab between number and text is the list's implicit tab

---

## 8. Header/Footer (header_footer)

### 8.1 Header Position

```
fn header_y() -> f32:
    return header_distance  // Distance from page top (used directly)
```

**COM measurement confirmed (2026-03-29):** headerDistance=18 -> y=18, 36 -> 36, 54 -> 54 (exact match)

### 8.2 Body Start Position

```
fn body_start_y(top_margin, header_bottom) -> f32:
    // Normal: topMargin + approx. 2.5pt offset
    // If header exceeds topMargin: header_bottom + gap
    return max(top_margin, header_bottom + gap) + ~2.5pt
```

**COM measurement confirmed:**
- Normal (header < topMargin): body_y = topMargin + 2.5pt
- Tall header (3 lines 14pt, topMargin=72): header_bottom approx. 87 -> body_y=90pt (pushdown)

### 8.3 Footer Position

```
fn footer_y(page_height, footer_distance) -> f32:
    // Y position of footer text
    return page_height - footer_distance - footer_line_height
    // Calibri 11pt: footer_line_height approx. 13.4pt
```

**COM measurement confirmed:**
- footerDist=18: footer_y=810.5, from_bottom=31.4 (18+13.4)
- footerDist=36: footer_y=792.5, from_bottom=49.4 (36+13.4)
- footerDist=54: footer_y=774.5, from_bottom=67.4 (54+13.4)

---

## 9. Footnotes (footnotes)

### 9.1 Footnote Default Style

- Font: 10.5pt (document default)
- LineSpacing: 12pt (Single)
- SpaceBefore/After: 0pt

### 9.2 Footnote Position

```
fn footnote_area(page_height, bottom_margin, footnotes) -> (f32, f32):
    // Footnotes are placed at the bottom of the page body area
    // body_area_bottom = page_height - bottom_margin
    // footnote_area extends upward from body_area_bottom
    area_height = separator_height + sum(fn.line_height for fn in footnotes)
    footnote_start_y = body_area_bottom - area_height
    return (footnote_start_y, body_area_bottom)
```

**COM measurement confirmed (2026-03-29):**
- Single footnote: y=752.5pt (body bottom=769.9 -> 17.4pt above)
- Multiple footnotes (3): y=717.0, 735.0, 752.5 (interval approx. 17-18pt)
- Footnotes compress the body area (body line count decreases)

---

## 11. Section Breaks (section_break)

### 11.1 Continuous Section Break

- Changes section within the same page
- Y coordinate continues from the end of the previous section (no reset)
- Format changes such as margins and column count are applied within the same page

**COM measurement confirmed (2026-03-29):**
- Section 1 last paragraph y=110.5, Section 2 first paragraph y=146.5 (blank line for section break)

### 11.2 nextPage Section Break

- Forced page break + section change
- Section 2 starts from the beginning of the new page
- Margin change: Section 2 leftMargin=108pt -> x=108pt (applied immediately)

### 11.3 Continuous + Column Change

- Section 1 (1 column) -> continuous break -> Section 2 (2 columns)
- Column change is applied within the same page
- Section 2 column area starts below Section 1's body text

---

## 13. Table Cell Padding (cell_padding)

### 13.1 Default Values

| Parameter | Default Value |
|-----------|-----------|
| LeftPadding | **4.95pt** (approx. 0.069in) |
| RightPadding | **4.95pt** |
| TopPadding | **0.0pt** |
| BottomPadding | **0.0pt** |

**COM measurement confirmed (2026-03-29):** Cell.LeftPadding=4.95, Cell.TopPadding=0.0

### 13.2 Cell-Level Override

- Table-level (tbl.LeftPadding=10) and cell-level (cell.LeftPadding=20) can coexist
- **Cell-level takes priority**
- COM confirmed: tbl=10, cell(1,1)=20 -> R1C1 text_x is 10pt right of R2C1

### 13.3 Border Width Effect on Text Position

```
fn text_position_in_cell(cell_x, padding, border_width) -> f32:
    return cell_x + padding + border_width / 2.0
```

**COM measurement confirmed:**
- border=0: text_x=77.0
- border=4halfpt(2pt): text_x=77.0 (no difference? absorbed into padding)
- border=12halfpt(6pt): text_x=77.5 (+0.5pt)
- border=24halfpt(12pt): text_x=78.5 (+1.5pt)

### 13.4 Cell Vertical Alignment (vAlign)

```
fn cell_text_y(valign, row_top, row_height, content_height, top_padding) -> f32:
    match valign:
        "top":    return row_top + top_padding
        "center": return row_top + (row_height - content_height) / 2.0
        "bottom": return row_top + row_height - content_height
```

**COM measurement confirmed (2026-03-29, row_height=60pt):**

| vAlign | 1 line (~18pt) | 2 lines (~36pt) | 3 lines (~54pt) |
|--------|-----------|-----------|-----------|
| top | 102.5 | 102.5 | 102.5 |
| center | 184.0 | 175.0 | 166.0 |
| bottom | 265.5 | 247.5 | 229.5 |

- center 1 line: row_top(162.5) + (60-18)/2 = 183.5 -> measured 184.0 (+/-0.5pt)
- center 3 lines: row_top(162.5) + (60-54)/2 = 165.5 -> measured 166.0 (+/-0.5pt)

---

## 15. Indent Inheritance (indent_inheritance)

### 15.1 leftChars vs left (twips)

```
fn effective_indent(left_twips, left_chars, char_width) -> f32:
    if left_chars is Some:
        // leftChars overrides left (not additive)
        return left_chars / 100.0 * char_width
    else:
        return left_twips / 20.0
```

**COM measurement confirmed (2026-03-29):**
- left=720tw only -> li=36.0pt (720/20)
- leftChars=200 only -> li=21.0pt (200/100 x 10.5pt)
- left=720 + leftChars=400 -> li=**42.0pt** (leftChars takes priority, 400/100 x 10.5)

### 15.2 Style Inheritance Chain

```
priority: direct(paragraph/run XML) > style > basedOn chain > docDefaults
```

**COM measurement confirmed (2026-03-29):**

| Property | docDefaults | Normal style | direct | Result |
|-----------|------------|-------------|--------|------|
| font | Arial | TNR | none | **TNR** (Normal wins) |
| font | Arial | TNR | Calibri | **Calibri** (direct wins) |
| size | 11pt | 12pt | none | **12pt** |
| size | 11pt | 12pt | 9pt | **9pt** |
| spaceAfter | 8pt | 10pt | none | **10pt** |
| spaceAfter | 8pt | 10pt | 0pt | **0pt** (explicit 0 is also effective) |
| leftIndent | - | 18pt | none | **18pt** |
| leftIndent | - | 18pt | 0pt | **0pt** |

### 15.3 Table Cell Line Height and Grid Snap -- compat_mode Dependent

**Grid snap inside table cells differs by compatibility mode.**

```
fn cell_line_gap(font, font_size, grid_pitch, compat_mode) -> f32:
    lh = gdi_line_height(font, font_size)
    if compat_mode >= 15:
        // Word 2013+: grid snap enabled inside cells
        if grid_pitch > 0:
            lh = grid_snap(lh, grid_pitch)
    // else: compat=14 (Word 2010): grid snap disabled
    return lh + border_overhead
```

**COM measurement confirmed (2026-03-29):**

| compat | grid | body gap | table gap | Interpretation |
|--------|------|----------|-----------|------|
| 15 | none | 18.0 | **18.5** | body=18(snap), table=18+0.5 |
| 14 | lines/360(18pt) | 20.5 | **16.0** | body=snap, table=**no snap**(13.5+2.5) |

- compat=15: grid snap inside table **enabled** (even with adjustLineHeightInTable=False)
- compat=14: grid snap inside table **disabled**
- adjustLineHeightInTable=False is common across all tests (no impact)

---

## 14. TextBox Internal Margins (textbox_padding)

### 14.1 Default Values

| Parameter | Default Value |
|-----------|-----------|
| MarginLeft | **7.2pt** (approx. 0.1in) |
| MarginRight | **7.2pt** |
| MarginTop | **3.6pt** (approx. 0.05in) |
| MarginBottom | **3.6pt** |

**COM measurement confirmed (2026-03-29)**

### 14.2 Table vs TextBox Comparison

| | Table Cell | TextBox |
|---|-----------|---------|
| Left pad | 4.95pt | 7.2pt |
| Top pad | 0.0pt | 3.6pt |
| Right pad | 4.95pt | 7.2pt |
| Bottom pad | 0.0pt | 3.6pt |

**TextBox has larger padding (left/right +2.25pt, top/bottom +3.6pt)**

### 14.3 Custom Padding Verification

| Setting | MarginL | text_offset_x | text_offset_y |
|------|---------|--------------|--------------|
| default | 7.2 | close to 7.2 | close to 3.6 |
| zero | 0.0 | ~0 | ~0 |
| large(20) | 20.0 | ~20 | ~15 |
| asymmetric(10) | 10.0 | ~10 | ~8 |

### 13.5 Table Row Height (trHeight)

```
fn row_height(height_rule, specified, content_height) -> f32:
    match height_rule:
        "auto":    return content_height  // Fit to content
        "exact":   return specified       // Use specified value directly
        "atLeast": return max(content_height, specified)
```

**COM measurement confirmed (2026-03-29):**

| rule | specified | content | actual gap |
|------|----------|---------|-----------|
| exact=20pt | 20 | - | 20.0 ✓ |
| exact=30pt | 30 | - | 30.0 ✓ |
| exact=50pt | 50 | - | 50.0 ✓ |
| atLeast=25, 1 line | 25 | ~18 | 25.5 (approx. specified) |
| atLeast=25, 3 lines | 25 | ~54 | 54.5 (content wins) |
| auto, 1 line | - | ~18 | 18.5 |
| auto, 2 lines | - | ~36 | 36.5 |

- auto row height approx. n_paras x grid_snap(line_height) + border
  - Calibri 11pt: natural=13.5 -> grid_snap=18pt. 1-line row gap=18.5pt (+0.5 border)
  - P1 -> P2 gap in cell: 18pt (grid snap confirmed, ls=12pt reported but effectively 18pt)
- Paragraph -> table first row: gap=18.5pt (when spaceAfter=0)

---

## 16. Nested Tables (nested_tables)

### 16.1 Nested Table Width

```
fn nested_table_width(parent_cell_width, parent_padding_l, parent_padding_r) -> f32:
    return parent_cell_width - parent_padding_l - parent_padding_r
```

**COM measurement confirmed (2026-03-29, 3-level nesting):**

| Level | Cell width | Padding L/R | Content area |
|-------|-----------|------------|-------------|
| outer | 400.0 | 4.95 | 390.1 |
| mid | 390.1 | 4.95 | 380.2 |
| inner | 380.2 | 4.95 | 370.3 |

- **text_x increment = parent_padding (5pt/level)**: 77 -> 82 -> 87

### 16.2 Table Inside TextBox

- Table width = TextBox content_width = TB.width - TB.marginL - TB.marginR
- TB w=300, marginL/R=7.2 -> content=285.6 -> table total=285.85 approx. 285.6

---

## 17. Shape Positioning (shape_positioning)

### 17.1 Position Reference (relativePosition)

| h_rel | v_rel | Reference |
|-------|-------|------|
| 0 (page) | 0 (page) | Page top-left |
| 1 (margin) | 1 (margin) | Margin top-left |
| 2 (column) | 2 (paragraph) | Column / anchor paragraph |

### 17.2 Text Flow Effect of Wrap Types

**COM measurement confirmed (2026-03-29, shape w=100, h=60 at margin-left):**

| Wrap | Text behavior |
|------|-------------|
| None(3) | No effect (text overlaps with shape) |
| Square(0) | Adjacent lines pushed right (x+shape_w) |
| TopAndBottom(1) | Adjacent lines pushed right |
| Tight(4) | Entire text moves below shape |

---

## 11. Character Grid (docGrid type="linesAndChars")

### 11.1 Character Pitch Calculation

```
charPitch = content_width / floor(content_width / defaultFontSize)
```

- `content_width = pageWidth - leftMargin - rightMargin`
- `defaultFontSize` = docDefaults rPr sz (typically 10.5pt)
- When charSpace XML attribute is specified: `base = defaultFontSize + charSpace * defaultFontSize / 4096`
- COM GridDistanceHorizontal = charPitch / 2 (half-width portion)

**1ec document example:** content_width=510.2pt, defaultFS=10.5pt -> floor(510.2/10.5)=48 chars -> charPitch=510.2/48=10.629pt

### 11.2 Character Width Grid Snap

Applied to text inside TextBoxes. NOT applied to body paragraphs (adjusted via justify).

```
effective_char_width = ceil(natural_char_width / charPitch) * charPitch
```

**Example:** MS Gothic 18pt (natural=18pt) on charPitch=10.629pt
- ceil(18/10.629) = 2 -> effective = 2 x 10.629 = 21.258pt/char
- inner_width=371.85pt -> floor(371.85/21.258) = 17 chars

### 11.3 Line Grid

Applied to body paragraphs and inside table cells. Disabled inside TextBoxes when docGrid type="linesAndChars".

```
line_height = ceil(natural_height / linePitch) * linePitch
// Round to 10tw (0.5pt)
line_height = round(line_height * 2) / 2
```

---

## 12. Table Cell Line Height and vAlign

### 12.1 Table Row Height

```
content_h = max(ceil(CJK_height / grid_pitch) * grid_pitch, over all cells)
content_h = round_to_10tw(content_h)  // 0.5pt rounding
row_height = max(content_h, trHeight)  // atLeast
```

- Line wrap width inside table cells: `cell_w` (cellMargin is NOT subtracted)
- Word allows text to overflow into cellMargin

### 12.2 vAlign=center

```
text_y = cell_top + (row_height - CJK_content_height) / 2
```

COM Y gap values include vAlign offset differences, so they don't match row_height:
```
COM_gap = row_height + vAlign_offset_next - vAlign_offset_current
```

### 12.3 vAlign Settings

XML `<w:vAlign w:val="center"/>` is set per cell. Cannot be directly obtained via COM but can be verified in XML.

---

## 13. TextBox Internal Layout

### 13.1 Grid Snap

- `docGrid type="linesAndChars"`: grid snap for line height inside TextBox is **disabled**
- `docGrid type="lines"`: grid snap for line height inside TextBox is **enabled**
- COM measurement: Shape2 P2 -> P3 gap=17.0pt (does not match grid pitch 17.85pt)

### 13.2 Page Break / Orphan Control

**Disabled** inside TextBoxes. Overflowing text is clipped.

### 13.3 Justify

Justify (jc=both) is also **disabled** inside TextBoxes. Text is left-aligned.

### 13.4 Text Y Position Offset

- exact/atLeast spacing: extra space is placed **above** the text (text at bottom)
  - `text_y_offset = (line_spacing - natural_height).max(0)`
- single spacing (grid snap): fontSize is **centered** within the grid cell
  - `text_y_offset = (grid_snapped_height - fontSize) / 2`
  - Note: uses fontSize, NOT CJK_natural (GDI TextOutW character cell = fontSize)

### 13.5 Shape Stroke

XML `strokeweight` value is used directly. Separate from table-level borders.
- Shape1-5: strokeweight=1pt
- TextBox IR holds `stroke_color` / `stroke_width`

---

## 14. GDI Character Width (ABC Width)

### 14.1 CJK Monospaced Fonts (UPM=256)

```
advance_px = ceil_even(ppem)  // Round ppem up to even number
advance_pt = advance_px * 72 / 96
```

GDI GetCharABCWidthsW A+B+C total = ceil_even(ppem). Matches for all sizes and all characters.

**Examples:**
| fontSize | ppem | ceil_even | advance_pt |
|----------|------|-----------|------------|
| 10.5pt | 14 | 14 | 10.5pt |
| 12pt | 16 | 16 | 12.0pt |
| 14pt | 19 | 20 | 15.0pt |
| 18pt | 24 | 24 | 18.0pt |
| 20pt | 27 | 28 | 21.0pt |

### 14.2 Proportional Fonts

GDI width override table (`gdi_width_overrides.json`, 1055KB) holds individual character widths.
9 fonts x ppem 7-100 x major character code points.

---

## 15. GDI Renderer Specification

### 15.1 Font Creation

```c
CreateFontW(
    -(int)round(fontSize * dpi / 72),  // lfHeight: round, NOT truncate
    0, 0, 0, weight,
    0, 0, 0,
    DEFAULT_CHARSET,
    0, 0,
    CLEARTYPE_QUALITY,  // ClearType quality
    0, fontName
)
```

### 15.2 Text Rendering

```c
TextOutW(hdc, round(x * dpi/72), round(y * dpi/72), text, len)
```

Coordinate conversion: pt -> pixel uses `round()`. NOT `truncate`.

### 15.3 Page Size

```c
width_px = ceil(pageWidth_pt * dpi / 72)
height_px = ceil(pageHeight_pt * dpi / 72)
```

---

## 16. SSIM=1.0 Verification Method

### 16.1 Word Reference Image

```
Word COM -> doc.Content.CopyAsPicture() -> CF_ENHMETAFILE ->
SetEnhMetaFileBits -> PlayEnhMetaFile(content_area_rect) -> GetDIBits -> PNG
```

EMF rendering target RECT includes margin offsets:
```
RECT = { margin_left, margin_top, width-margin_right, height-margin_bottom }
```

### 16.2 Oxi Reference Image

```
oxidocs-core::layout::LayoutEngine -> GDI TextOutW/FillRect/RoundRect -> GetDIBits -> PNG
```

### 16.3 Comparison

Compare band-by-band using skimage.metrics.structural_similarity (SSIM).
Goal: SSIM=1.0000 for all bands.

---

## Measurement Basis

- All values measured on Windows + Word 365 + COM API (win32com)
- GDI character widths measured via GetCharABCWidthsW / GetTextExtentPoint32W
- GDI font metrics measured via GetTextMetricsW
- Test documents dynamically generated with python-docx or created directly via Word COM
- Integrated results from Ra (automated specification analysis engine) + manual measurements
- GDI vs GDI SSIM verification: SSIM=1.0000 (pixel-perfect match) achieved in 0-6% band (2026-03-30)
