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
            // COM実測 (2026-03-29):
            // 値は twips/20 ptをそのまま使用（丸めなし）
            // WordはY座標をピクセルスナップ(0.5pt量子化)するため、
            // 非整数pt値では行間が交互に微小変動（例: 9.15pt→9.0,9.5,9.0,9.0）
            // しかし平均値は期待値と一致（183tw avg=9.125≈9.15, 誤差0.025pt）
            // 実装: twips/20のpt値をそのまま使用。レンダリング時にピクセルスナップ
            return lh
        "atLeast":
            natural = word_line_height(font_metrics, font_size)
            // grid snap は natural に適用、specified にはなし
            if grid_pitch > 0:
                natural = grid_snap(natural, grid_pitch)
            lh = max(natural, value)
            return lh  // grid snap済み or specified(snap不要)

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

**GDI実測 vs Oxi計算の差異 (2026-03-29):**

| フォント | ppem | 不一致数/63文字 | 原因 |
|---------|------|----------------|------|
| Calibri | 12 (9pt) | **17** | GDIヒンティング |
| Calibri | 14 (10.5pt) | **0** | 完全一致 |
| Calibri | 15 (11pt) | **16** | GDIヒンティング |
| Arial | 12-15 | **11-18** | GDIヒンティング |
| MS Gothic/Mincho | 全サイズ | **0** | UPM=256で丸め不要 |

**注意:** `round(advance * ppem / upm)` はGDIヒンティング非適用時の近似値。
TrueTypeフォント(UPM=2048)は、GDIがヒンティング命令で幅を調整するため最大1px差が生じる。
この差が行折り返し位置→行数→Y座標累積ずれの主因。
**解決策: GDI幅オーバーライドテーブル**
- `gdi_pixel_overrides.json` (14.8KB): Oxi計算値と異なる1888箇所のみ
- `gdi_width_overrides.json` (1055KB): 全フォント完全テーブル
- 9フォント × ppem 7-20 × 894文字を Windows GDI で実測済み
- Bold フォントは Regular より差異が多い (Arial Bold: 500箇所)

### 4.6 行折り返し判定

```
fn needs_line_break(accumulated_width_px, content_width_px) -> bool:
    return accumulated_width_px > content_width_px
    // 注意: > であり >= ではない（content_widthちょうどは折り返さない）
```

**GDI実測確定 (2026-03-29):**
- Calibri 11pt 'A'(9px)×86 = 774px... ×87=783px: content=602px
  - n=86 string_w=602px → **1行** (width == content → 折り返さない)
  - n=87 string_w=609px → **2行** (width > content → 折り返す)
- `GetTextExtentPoint32W(全文字列)` = 個別文字幅の合計と一致
  - string_width == sum(char_widths) (Wordでも同様)
- Mixed text: line 1 gdi_w=459.75pt > content(451.3pt) → 正しく折り返し

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

**COM実測確定 (2026-03-29, MS Gothic 9pt):**

| cs(tw) | GDI px | GDI pt | 実測CJK gap | base gap |
|--------|--------|--------|-----------|----------|
| 0 | 0 | 0 | 9.0pt | 9.0pt |
| -9 | -2 | -1.5 | 8.5pt | -0.5pt |
| 9 | 1 | 0.75 | 9.5pt | +0.5pt |
| 20 | 1 | 0.75 | 10.0pt | +1.0pt |
| -20 | -2 | -1.5 | 8.0pt | -1.0pt |

**注意:** COM座標は0.5pt量子化の影響あり。実装ではGDI MulDiv計算値をそのまま使用すべき。

### 4.4 等幅CJKフォント (UPM=256)

MS Gothic, MS Mincho:

```
fn cjk_fullwidth_px(font_size_pt) -> i32:
    ppem = round(font_size_pt * 96.0 / 72.0)
    // 偶数ピクセルに切り上げ (GDI実測 2026-03-29)
    return (ppem + 1) & !1  // ceil to even
    // 半角 = fullwidth / 2
```

**GDI実測確定 (2026-03-29):**

| fontSize | ppem | CJK全角px | 公式 |
|----------|------|----------|------|
| 7pt | 9 | 10 | ceil_even(9)=10 ✓ |
| 8pt | 11 | 12 | ceil_even(11)=12 ✓ |
| 9pt | 12 | 12 | 12(偶数) ✓ |
| 10pt | 13 | 14 | ceil_even(13)=14 ✓ |
| 10.5pt | 14 | 14 | 14(偶数) ✓ |

**注意:** `ceil_even` はMS Gothic/MS Minchoのみ（UPM=256ビットマップ等幅フォント）。

### 4.5 その他のCJKフォントの全角幅

Yu Gothic, Yu Mincho, Meiryo: **fullwidth = ppem** (偶数丸めなし)

```
fn cjk_fullwidth_other(font_size_pt) -> i32:
    return round(font_size_pt * 96.0 / 72.0)  // ppem直接
```

MS PGothic, MS PMincho: **プロポーショナル** (文字ごとにGDI幅が異なる)

**GDI実測確定 (2026-03-29, ppem=5-20全パターン確認):**
- MS Gothic/Mincho: ceil_even ALL MATCH (ppem 5-29)
- Yu Gothic/Mincho/Meiryo: CJK全角 = ppem (全サイズ一致)
- MS PGothic/PMincho: プロポーショナル（「あ」≠ppem、文字幅のGDI個別計算が必要）

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

**COM実測確定 (2026-03-29):**
- DefaultTabStop はドキュメントごとに異なる
- ja_gov_template.docx: 36pt
- Normal.dotm (日本語Word): 42pt (= 4文字 × 10.5pt)
- 値は `w:settings/w:defaultTabStop w:val` (twips) または Document.DefaultTabStop で取得

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

**カスタムタブ位置はマージン起点の絶対位置であり、indentの影響を受けない。**

```
fn next_custom_tab(current_x_from_margin, tab_stops, effective_indent) -> Option<f32>:
    // indent より前のタブ位置はスキップ
    for tab in tab_stops:
        if tab.position > current_x_from_margin && tab.position >= effective_indent:
            return Some(tab.position)  // マージンからの絶対位置
    return None
```

**COM実測確定 (2026-03-29, Selection.Information(5)で完全確認):**

| 設定 | Seg0(text start) | Seg1(tab1) | Seg2(tab2) |
|------|-----------------|-----------|-----------|
| indent=0, tab@144,288 | margin+0 | margin+144 | margin+288 |
| indent=36, tab@144,288 | margin+36 | margin+144 | margin+288 |
| indent=72, tab@144,288 | margin+72 | margin+144 | margin+288 |
| indent=180, tab@144,288 | margin+180 | margin+288 | margin+336 |
| hanging=36/indent=72, P1 | margin+36 | margin+72* | margin+144 |
| hanging=36/indent=72, P2 | margin+72 | margin+144 | margin+288 |

\* P1のSeg1(margin+72)はleftIndent位置への**暗黙タブ**(hanging indent自動生成)
| firstLine=36, P1 | margin+36 | margin+144 | margin+288 |
| firstLine=36, P2 | margin+0 | margin+144 | margin+288 |

- テキスト開始位置 = margin + effective_indent（完全一致、誤差0.0pt）
- indent=180, tab@144: **tab@144 < indent → スキップ**、tab@288使用
- hanging indent: P1 effective=36(72-36), P2 effective=72

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

### 7.5 段落途中のカラム分割 (mid-paragraph column break)

**段落の行がカラム底を超えた場合、残りの行は次カラムのtop(=start_y)から継続。**

```
fn column_line_overflow(cursor_y, line_height, col_bottom, next_col_start_y) -> f32:
    if cursor_y + line_height > col_bottom:
        return next_col_start_y  // 次カラムのY開始位置
    return cursor_y
```

**COM実測確定 (2026-03-29):**
- 35短段落でカラム1を埋めた後、長段落(16行):
  - カラム1に3行(y=704.5, 722.5, 740.5)
  - カラム2に13行(y=74.5から、col2 x=308.5)
  - **Y座標はtop_marginにリセット(74.5)**
- keepTogether=True: 段落全体がカラム2先頭(x=308.5, y=74.5)に移動
- **ページ内行分割と同一ロジック**

---

## 12. 番号リスト (numbering)

### 12.1 基本レイアウト

```
fn list_paragraph_layout():
    // 番号リスト = hanging indent + list marker
    // leftIndent = テキスト開始位置（マージンから）
    // firstLineIndent = -leftIndent（hanging: マーカー領域の幅）
    // マーカーは margin+0 ～ margin+leftIndent に配置
```

**COM実測確定 (2026-03-29):**
- 基本番号リスト(1.2.3.): leftIndent=22pt, firstLineIndent=-22pt
- テキスト開始位置 = margin+22pt (= leftIndent)
- 弾丸リスト: 同一(li=22, fli=-22), マーカー=U+F06C (Wingdings)

### 12.2 ネストレベル別インデント

| Level | leftIndent | firstLineIndent | テキスト開始(margin-rel) |
|-------|-----------|-----------------|----------------------|
| 1 | 22.0pt | -22.0pt | 22.0pt |
| 2 | 32.5pt | -22.0pt | 32.5pt |
| 3 | 43.0pt | -22.0pt | 43.0pt |

- Level間のインデント増分: +10.5pt
- firstLineIndentは全レベル共通(-22pt)

### 12.3 リスト+カスタムタブの相互作用

- リスト内テキストはleftIndent位置から開始
- カスタムタブ(@144pt)はマージン絶対位置（リストindentに影響されない）
- 番号→テキスト間のタブはリストの暗黙タブ

---

## 8. ヘッダー/フッター (header_footer)

### 8.1 ヘッダー位置

```
fn header_y() -> f32:
    return header_distance  // ページ上端からの距離（そのまま）
```

**COM実測確定 (2026-03-29):** headerDistance=18→y=18, 36→36, 54→54（完全一致）

### 8.2 本文開始位置

```
fn body_start_y(top_margin, header_bottom) -> f32:
    // 通常: topMargin + 約2.5ptオフセット
    // ヘッダーが topMargin を超える場合: header_bottom + gap
    return max(top_margin, header_bottom + gap) + ~2.5pt
```

**COM実測確定:**
- 通常（ヘッダー < topMargin）: body_y = topMargin + 2.5pt
- 背の高いヘッダー（3行14pt, topMargin=72）: header_bottom≈87 → body_y=90pt（pushdown）

### 8.3 フッター位置

```
fn footer_y(page_height, footer_distance) -> f32:
    // フッターテキストのY位置
    return page_height - footer_distance - footer_line_height
    // Calibri 11pt: footer_line_height ≈ 13.4pt
```

**COM実測確定:**
- footerDist=18: footer_y=810.5, from_bottom=31.4 (18+13.4)
- footerDist=36: footer_y=792.5, from_bottom=49.4 (36+13.4)
- footerDist=54: footer_y=774.5, from_bottom=67.4 (54+13.4)

---

## 9. 脚注 (footnotes)

### 9.1 脚注デフォルトスタイル

- フォント: 10.5pt（ドキュメントデフォルト）
- LineSpacing: 12pt (Single)
- SpaceBefore/After: 0pt

### 9.2 脚注位置

```
fn footnote_area(page_height, bottom_margin, footnotes) -> (f32, f32):
    // 脚注はページ本文領域の下端に配置
    // body_area_bottom = page_height - bottom_margin
    // footnote_area は body_area_bottom から上に伸びる
    area_height = separator_height + sum(fn.line_height for fn in footnotes)
    footnote_start_y = body_area_bottom - area_height
    return (footnote_start_y, body_area_bottom)
```

**COM実測確定 (2026-03-29):**
- 単一脚注: y=752.5pt (body bottom=769.9 → 17.4pt上)
- 複数脚注(3個): y=717.0, 735.0, 752.5 (間隔≈17-18pt)
- 脚注はbody areaを圧迫（bodyの行数が減る）

---

## 11. セクション区切り (section_break)

### 11.1 continuous section break

- 同一ページ内でセクション変更
- Y座標は前セクションの最後から連続（リセットなし）
- マージン・カラム数等のフォーマット変更が同一ページ内で適用

**COM実測確定 (2026-03-29):**
- Section 1 最終段落 y=110.5, Section 2 最初段落 y=146.5（セクション区切り分の空行あり）

### 11.2 nextPage section break

- 強制改ページ + セクション変更
- Section 2 は新ページの先頭から開始
- マージン変更: Section 2 leftMargin=108pt → x=108pt（即座に反映）

### 11.3 continuous + column変更

- Section 1 (1カラム) → continuous break → Section 2 (2カラム)
- カラム変更は同一ページ内で適用される
- Section 2 のカラム領域は Section 1 の本文の下から開始

---

## 13. テーブルセルパディング (cell_padding)

### 13.1 デフォルト値

| パラメータ | デフォルト値 |
|-----------|-----------|
| LeftPadding | **4.95pt** (≈0.069in) |
| RightPadding | **4.95pt** |
| TopPadding | **0.0pt** |
| BottomPadding | **0.0pt** |

**COM実測確定 (2026-03-29):** Cell.LeftPadding=4.95, Cell.TopPadding=0.0

### 13.2 セルレベルオーバーライド

- テーブルレベル (tbl.LeftPadding=10) とセルレベル (cell.LeftPadding=20) が混在可能
- **セルレベルが優先**
- COM確認: tbl=10, cell(1,1)=20 → R1C1 text_x は R2C1 より10pt右

### 13.3 ボーダー幅のテキスト位置影響

```
fn text_position_in_cell(cell_x, padding, border_width) -> f32:
    return cell_x + padding + border_width / 2.0
```

**COM実測確定:**
- border=0: text_x=77.0
- border=4halfpt(2pt): text_x=77.0 (差なし? padding内に吸収)
- border=12halfpt(6pt): text_x=77.5 (+0.5pt)
- border=24halfpt(12pt): text_x=78.5 (+1.5pt)

### 13.4 セル垂直配置 (vAlign)

```
fn cell_text_y(valign, row_top, row_height, content_height, top_padding) -> f32:
    match valign:
        "top":    return row_top + top_padding
        "center": return row_top + (row_height - content_height) / 2.0
        "bottom": return row_top + row_height - content_height
```

**COM実測確定 (2026-03-29, row_height=60pt):**

| vAlign | 1行(~18pt) | 2行(~36pt) | 3行(~54pt) |
|--------|-----------|-----------|-----------|
| top | 102.5 | 102.5 | 102.5 |
| center | 184.0 | 175.0 | 166.0 |
| bottom | 265.5 | 247.5 | 229.5 |

- center 1行: row_top(162.5) + (60-18)/2 = 183.5 → 実測184.0 (±0.5pt)
- center 3行: row_top(162.5) + (60-54)/2 = 165.5 → 実測166.0 (±0.5pt)

---

## 15. インデント継承 (indent_inheritance)

### 15.1 leftChars vs left (twips)

```
fn effective_indent(left_twips, left_chars, char_width) -> f32:
    if left_chars is Some:
        // leftChars が left を上書き（加算ではない）
        return left_chars / 100.0 * char_width
    else:
        return left_twips / 20.0
```

**COM実測確定 (2026-03-29):**
- left=720tw のみ → li=36.0pt (720/20)
- leftChars=200 のみ → li=21.0pt (200/100 × 10.5pt)
- left=720 + leftChars=400 → li=**42.0pt** (leftCharsが優先、400/100 × 10.5)

### 15.2 スタイル継承

- Normal style li=18pt → 段落に明示設定なし → **li=18pt(継承)**
- Normal style li=18pt → 段落で li=36pt → **li=36pt(上書き)**
- Normal style li=18pt → 段落で li=0pt → **li=0pt(明示0で上書き)**

---

## 14. TextBox内部マージン (textbox_padding)

### 14.1 デフォルト値

| パラメータ | デフォルト値 |
|-----------|-----------|
| MarginLeft | **7.2pt** (≈0.1in) |
| MarginRight | **7.2pt** |
| MarginTop | **3.6pt** (≈0.05in) |
| MarginBottom | **3.6pt** |

**COM実測確定 (2026-03-29)**

### 14.2 テーブル vs TextBox 比較

| | Table Cell | TextBox |
|---|-----------|---------|
| Left pad | 4.95pt | 7.2pt |
| Top pad | 0.0pt | 3.6pt |
| Right pad | 4.95pt | 7.2pt |
| Bottom pad | 0.0pt | 3.6pt |

**TextBoxの方がパディングが大きい（左右+2.25pt、上下+3.6pt）**

### 14.3 カスタムパディング検証

| 設定 | MarginL | text_offset_x | text_offset_y |
|------|---------|--------------|--------------|
| default | 7.2 | 7.2に近い | 3.6に近い |
| zero | 0.0 | ~0 | ~0 |
| large(20) | 20.0 | ~20 | ~15 |
| asymmetric(10) | 10.0 | ~10 | ~8 |

---

## 計測根拠

- 全値はWindows + Word 365 + COM API (win32com) で計測
- GDI文字幅は GetTextExtentPoint32W で計測
- GDIフォントメトリクスは GetTextMetricsW で計測
- テスト文書は python-docx で動的生成 or Word COM で直接作成
- Ra (仕様自動解析エンジン) + 手動計測の統合結果
