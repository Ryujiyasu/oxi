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
- **lineSpacing/spaceAfter/spaceBefore: 自動リセットは無い**
  - COM確認 (2026-03-29): docDefaults/Normalスタイル経由のsa/sb/lsは、テーブルセル・TextBox内でも保持される
  - 以前「リセット」と見えた挙動はテーブルスタイルがNormalのspacingをオーバーライドした結果
  - セル内段落のspacingはスタイル継承チェーン次第（テーブルセル固有の自動リセット機能ではない）

### 1.6 TextBox内

- grid snap なし (grid_pitch = None として計算)
- CJK 83/64 は適用
- **spacingリセット: なし（テーブルセルと同様、スタイル継承チェーンに従う）** (COM確定 2026-03-29)

### 1.7 Mixed font run (CJK + Latin 混在行)

```
fn mixed_line_height(runs, grid_pitch) -> f32:
    lh = max(word_line_height(run.font_metrics, run.font_size) for run in runs)
    if grid_pitch > 0:
        lh = grid_snap(lh, grid_pitch)
    return lh
```

**COM実測確定 (2026-03-29):**
- Latin(Calibri 10.5pt)=14.5pt, CJK(MS Gothic 10.5pt)=15.5pt → mixed=15.5pt = max ✓
- grid snap は max の後に適用 (max → snap、not snap each → max)
- Y位置に±0.5ptの揺らぎあり (COM測定精度の限界、仕様上の問題ではない)

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

```
fn apply_contextual_spacing(prev_para, next_para) -> (f32, f32):
    // 片方でもcontextualSpacing=Trueかつ同一スタイルなら両方のsa/sbを0に
    if (prev_para.contextual_spacing || next_para.contextual_spacing)
       && prev_para.style == next_para.style:
        return (0.0, 0.0)  // (effective_sa, effective_sb)
    else:
        return (prev_para.space_after, next_para.space_before)
```

**COM実測確定 (2026-03-29):**
- 同一スタイル + ctx=ON: gap=15.5pt (行高さのみ、sa/sb完全抑制)
- 同一スタイル + ctx=OFF: gap=27.5pt (行高さ+spacing)
- 異なるスタイル + ctx=ON: gap=27.5pt (効果なし)
- **片方のみctx=ON でも抑制される** (P1のみTrueでもP1-P2間のspacing=0)
- asymmetric (sa=20,sb=10) ctx=ON: gap=15.5pt → 値に関わらず両方0に

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

## 6. タブストップ (tab_stops)

### 6.1 デフォルトタブ

```
fn default_tab_interval() -> f32:
    // Document.DefaultTabStop (通常 36pt = 0.5 inch)
    return doc.default_tab_stop  // twips / 20
```

**COM実測確定 (2026-03-29):** DefaultTabStop = 36pt

### 6.2 タブ位置の基準

タブ位置は**左マージン起点**の絶対位置（twips単位をpt変換）。

```
fn tab_position_pt(tab_pos_twips, margin_left_pt) -> f32:
    // 位置は左マージンからの距離
    return margin_left_pt + tab_pos_twips / 20.0
```

### 6.3 タブ種類別の配置

```
fn apply_tab(tab_type, tab_pos_pt, text_before_width, text_after) -> f32:
    match tab_type:
        "left":
            // テキスト開始位置 = タブ位置
            return tab_pos_pt
        "center":
            // テキスト中心 = タブ位置
            return tab_pos_pt - text_after_width / 2.0
        "right":
            // テキスト右端 = タブ位置
            return tab_pos_pt - text_after_width
        "decimal":
            // 小数点位置 = タブ位置
            return tab_pos_pt - width_before_decimal_point
```

**COM実測確定 (2026-03-29):**
- Left tab @72pt (1440tw): text starts at margin+72pt ✓
- Center tab @216pt (4320tw): "Center"(~30pt) → x=201pt, center≈216pt ✓
- Right tab @432pt (8640tw): "Right"(~23pt) → x=409pt, right≈432pt ✓
- Decimal tab @216pt (4320tw): "123.45" → decimal point at ≈216pt ✓

### 6.4 デフォルトタブの適用

カスタムタブが定義されていない場合、DefaultTabStop間隔で自動タブ位置が生成される。

```
fn next_tab_position(current_x, margin_left, default_interval) -> f32:
    // 現在位置より先の最初のデフォルトタブ位置
    offset = current_x - margin_left
    n = floor(offset / default_interval) + 1
    return margin_left + n * default_interval
```

### 6.5 タブとindentの相互作用

**タブ位置はマージン起点の絶対位置であり、indentの影響を受けない。**

```
fn next_tab(current_x_from_margin, tab_stops, indent) -> f32:
    // current_x はマージンからの相対位置
    // indent より前のタブ位置はスキップ
    for tab in tab_stops:
        if tab.position > current_x_from_margin && tab.position >= indent:
            return tab.position  // マージンからの絶対位置
    // カスタムタブがない場合、デフォルトタブへフォールバック
    return next_default_tab(current_x_from_margin)
```

**COM実測確定 (2026-03-29):**
- indent=0, tab@144: text at margin+144 ✓
- indent=36, tab@144: text at margin+144 (タブ位置は変わらない) ✓
- indent=72, tab@144: text at margin+144 ✓
- indent=180, tab@144: **スキップ** (タブ位置 < indent)。tab@288を使用 ✓
- firstLineIndent=36, tab@144: first line text at margin+144 ✓, P2(indent=0) at margin+144 ✓
- hanging=36/indent=72: P1(effective=36) uses tab@144, P2(effective=72) uses tab@144 ✓

### 6.6 タブリーダー

- `dot`: 点線リーダー
- `hyphen`: ダッシュリーダー
- `underscore`: 下線リーダー
- リーダー文字はタブ空白を埋める（目次等で使用）

---

## 7. マルチカラムレイアウト (columns)

### 7.1 カラム位置計算

```
fn column_x_positions(margin_left, columns) -> Vec<f32>:
    x = margin_left
    positions = []
    for i, col in columns:
        positions.push(x)
        x += col.width + col.space_after
    return positions
```

**COM実測確定 (2026-03-29):**

| 設定 | Col1 x | Col2 x | Col3 x |
|------|--------|--------|--------|
| 2col equal (w=215, sp=21.25) | 72.0 | 308.5 (≈308.25) | - |
| 3col equal (w=136.25, sp=21.25) | 72.0 | 229.5 | 387.0 |
| 2col gap=36 (w=207.65) | 72.0 | 315.5 (≈315.65) | - |
| 2col unequal (w1=150, w2=265.3, sp=36) | 72.0 | 258.0 | - |

### 7.2 均等幅カラムの計算

```
fn equal_column_width(text_width, num_cols, spacing) -> f32:
    // text_width = page_width - margin_left - margin_right
    return (text_width - spacing * (num_cols - 1)) / num_cols
```

### 7.3 テキストフロー

- テキストは Column 1 → Column 2 → ... → 次ページ Column 1 の順にフロー
- カラムの高さは通常ページ本文領域と同じ (top_margin ~ bottom_margin)
- Column Break (wdColumnBreak): 強制的に次カラムへ移動

### 7.4 カラムY座標

- 各カラムのY座標開始位置はページ上端マージンから同一
- テキストはページ上端から各カラムに独立してレイアウト

---

## 計測根拠

- 全値はWindows + Word 365 + COM API (win32com) で計測
- GDI文字幅は GetTextExtentPoint32W で計測
- GDIフォントメトリクスは GetTextMetricsW で計測
- テスト文書は python-docx で動的生成 or Word COM で直接作成
- Ra (仕様自動解析エンジン) + 手動計測の統合結果
