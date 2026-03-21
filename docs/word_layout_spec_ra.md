# Word レイアウト仕様書 (Ra + 手動計測統合版)

COM API ブラックボックス測定により解明。DLL解析なし。

---

## 1. 行高さ (line_height)

### 1.1 基本計算 (lineSpacingRule = auto/Single)

```
fn gdi_line_height(font_metrics, font_size_pt) -> f32:
    ppem = round(font_size_pt * 96.0 / 72.0)
    asc_px = ceil(win_ascent / upm * ppem)
    des_px = ceil(win_descent / upm * ppem)
    gdi_height_pt = (asc_px + des_px) * 72.0 / 96.0
    return gdi_height_pt
```

### 1.2 CJK 83/64 乗数

**適用条件:** フォントファミリーがCJKホワイトリストに含まれる場合

ホワイトリスト: MS Gothic, MS Mincho, Yu Gothic, Yu Mincho, Meiryo,
  MS PGothic, MS PMincho, HG系フォント

```
fn word_line_height(font_metrics, font_size_pt) -> f32:
    base = gdi_line_height(font_metrics, font_size_pt)
    if is_cjk_font(font_family):
        return base * 83.0 / 64.0
    else:
        return base
```

**COM実測確認:**
- MS Gothic 10.5pt noGrid: 14.25-15.00pt (平均≈14.6pt, CJK 83/64適用)
- Calibri 10.5pt noGrid: 13.5pt
- MS Gothic 10.5pt grid(18pt): 18.0pt (1×grid)
- Meiryo 10.5pt grid(18pt): 36.0pt (2×grid, 自然高が大きい)

### 1.3 lineSpacingRule別

```
fn line_height(rule, value, font_metrics, font_size, grid_pitch) -> f32:
    match rule:
        "auto" | "single":
            base = word_line_height(font_metrics, font_size)
            factor = value / 240.0  // w:line="240" = 1.0x
            lh = base * factor
        "exact":
            lh = value  // w:line in twips / 20
            // grid snap は適用されない (COM確定)
            return lh
        "atLeast":
            natural = word_line_height(font_metrics, font_size)
            lh = max(natural, value)

    // grid snap (type="lines" or "linesAndChars" のみ)
    if grid_pitch > 0:
        lh = grid_snap(lh, grid_pitch)

    return lh
```

### 1.4 グリッドスナップ

```
fn grid_snap(lh, pitch) -> f32:
    // round-half-up 方式 (COM Session2で確定)
    n = floor((lh + pitch / 2.0) / pitch + 0.5)  // ≡ round_half_up(lh / pitch)
    if n < 1: n = 1
    return n * pitch
```

**適用条件:**
- docGrid type="lines" → 適用
- docGrid type="linesAndChars" → 適用
- docGrid type なし (linePitchのみ) → 適用 (COM実測: gen_027で確認)
- docGrid 要素なし → 適用しない
- exact lineSpacingRule → 適用しない

### 1.5 テーブルセル内

- adjustLineHeightInTable=false (デフォルト): CJK 83/64 + grid snap 有効
- adjustLineHeightInTable=true: CJK 83/64 + grid snap 両方無効
- **lineSpacing/spaceAfter/spaceBefore: Normalスタイルの継承値はリセット**
  - COM確認: セル内段落は ls=12(Single), sa=0, sb=0
  - 明示的にXMLで指定されていない限り、デフォルトに戻る

### 1.6 TextBox内

- grid snap なし (grid_pitch = None として計算)
- CJK 83/64 は適用
- spacing リセット条件: テーブルセルと同様（未完全確認）

### 1.7 Mixed font run (CJK + Latin 混在行)

- 行高さ = max(各runのfont行高さ)
- COM計測は安定的に取得できず、1px程度の揺らぎあり

---

## 2. 段落間スペーシング (spacing)

### 2.1 基本ルール

```
fn paragraph_gap(prev_para, next_para) -> f32:
    sa = prev_para.space_after
    sb = next_para.space_before
    spacing = max(sa, sb)  // コラプス: 大きい方を採用
    return prev_line_height + spacing
```

**COM実測確認:**
- sa=10, sb=24 → spacing=24pt (max) ✓
- sa=24, sb=10 → spacing=24pt (max) ✓
- sa=10, sb=10 → spacing=9.75pt ← grid snap の影響
- sa=15, sb=15 → spacing=15pt ✓
- sa=0, sb=24 → spacing=24pt ✓

**注意:** spacingにもgrid snapが適用される場合がある。
sa=sb=10pt → 10pt → grid snap → 9.75pt (= 13px * 0.75)

### 2.2 ページ先頭での spaceBefore 抑制

```
fn effective_space_before(para, is_first_on_page) -> f32:
    if is_first_on_page:
        return 0.0  // 完全に抑制
    else:
        return para.space_before
```

**COM実測確認:**
- ページ2の最初の段落: sb=6pt → y=74.25 (sbなしと同じ)
- ページ2の2番目の段落: sb=12pt → 通常通り適用

### 2.3 contextualSpacing

- 同じスタイルの隣接段落間で space_after/space_before を 0 にする
- 異なるスタイルでは効果なし
- (COM計測でエラー発生、追加検証推奨)

### 2.4 beforeLines / afterLines

```
fn before_lines_to_pt(value, line_pitch) -> f32:
    return value / 100.0 * line_pitch
    // 結果にgrid snap適用 (snap_to_grid=true の場合)
```

---

## 3. グリッドスナップ (grid_snap)

### 3.1 適用条件

| docGrid | type | 行高さsnap | spacing snap |
|---------|------|-----------|-------------|
| type="lines" | あり | ✓ | ✓ |
| type="linesAndChars" | あり | ✓ | ✓ |
| linePitchのみ (typeなし) | あり | ✓ | ✓ |
| docGrid要素なし | なし | ✗ | ✗ |
| snap_to_grid=false | — | ✗ | ✗ |

### 3.2 丸め方式

```
fn grid_snap(value, pitch) -> f32:
    // round-half-up (2.5 → 3)
    n = floor(value / pitch + 0.5)
    if n < 1: n = 1
    return n * pitch
```

### 3.3 最初の段落のY位置

- margin_top をそのまま使用 (grid offsetなし)
- ただし行高さ自体がgrid snapされるため、結果的に整列する

---

## 4. 文字幅 (char_width)

### 4.1 基本計算

```
fn char_width_pt(font_metrics, char, font_size) -> f32:
    ppem = round(font_size * 96.0 / 72.0)
    advance = font_metrics.advance_width(char)  // UPM単位
    pixel_width = round(advance * ppem / upm)
    return pixel_width * 72.0 / 96.0
```

### 4.2 フォントフォールバック

**CJK文字にLatinフォントが指定された場合、GDIは自動的にフォールバックする。**

```
fn resolve_char_width(font_name, char, font_size) -> f32:
    if is_cjk_char(char) && is_latin_font(font_name):
        // GDI fallback: MS UI Gothic を使用
        return char_width_pt(ms_ui_gothic_metrics, char, font_size)
    else:
        return char_width_pt(font_metrics, char, font_size)
```

**COM/GDI実測確認 (ppem=14):**

| 文字 | Calibri | MS UI Gothic | MS Gothic |
|------|---------|-------------|-----------|
| あ (U+3042) | 11px | 11px | 14px |
| 一 (U+4E00) | 14px | 14px | 14px |
| A (U+0041) | 8px | 9px | 7px |

Calibri で「あ」= 11px = MS UI Gothic で「あ」= 11px → **フォールバック先は MS UI Gothic**

### 4.3 character spacing (w:spacing w:val)

```
fn apply_cs(cs_twips) -> f32:
    // GDI MulDiv 丸め
    cs_px = MulDiv(cs_twips, 96, 1440)  // = (cs_twips * 96 + 720) / 1440
    return cs_px * 72.0 / 96.0
```

### 4.4 等幅CJKフォント (UPM=256)

MS Gothic, MS Mincho:
- 全角文字幅 = fontSize (pt)
- 半角文字幅 = fontSize / 2 (pt)

---

## 5. ページ分割 (page_break)

### 5.1 基本条件

```
fn needs_page_break(cursor_y, line_height, page_bottom) -> bool:
    return cursor_y + line_height > page_bottom
```

### 5.2 widow/orphan 制御

- WidowControl=True: 段落の最初の1行だけがページ末尾に残る場合、段落全体を次ページへ移動
- WidowControl=True: 段落の最後の1行だけが次ページに行く場合、前のページから1行追加で移動

**COM確認:** WidowControl=True で3行段落が page 1 (y=740) から page 2 (y=74) に完全移動

### 5.3 keepWithNext / keepTogether

- keepWithNext: この段落と次の段落を同じページに保つ
- keepTogether: 段落内の全行を同じページに保つ

**COM確認:** KeepTogether=True で長い段落が page 2 に移動

### 5.4 テーブル行の分割

- AllowBreakAcrossPages = True (デフォルト): 行がページ間で分割される
- AllowBreakAcrossPages = False: 行全体が次ページに移動

### 5.5 spaceBeforeAutoSpacing

ページ先頭の段落のspaceBeforeは**完全に抑制**される（0ptとして扱う）。

---

## 計測根拠

- 全値はWindows + Word 365 + COM API (win32com) で計測
- GDI文字幅は GetTextExtentPoint32W で計測
- GDIフォントメトリクスは GetTextMetricsW で計測
- テスト文書は python-docx で動的生成 or Word COM で直接作成
- Ra (仕様自動解析エンジン) + 手動計測の統合結果
