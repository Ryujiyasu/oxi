# Wordの改ページ位置は、GDIの整数丸めで決まっている

Internet Explorer 11 は、文字幅の計算を従来の GDI 互換メトリクスから DirectWrite の natural metrics へ切り替えました。Microsoft は後に、互換性が必要なサイトでは `X-UA-TextLayoutMetrics: gdi` によって従来の計算へ戻せることを文書化しています。移行時の検証例には、natural metrics では 330px の箱に収まらず、同じ一行が折り返すケースもあります。

- [Microsoft: Turn off natural metrics](https://learn.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-it-pro/internet-explorer-11/ie11-deploy-guide/turn-off-natural-metrics)
- [Microsoft: Site text layout is different in Internet Explorer 11](https://support.microsoft.com/en-gb/topic/site-text-layout-is-different-in-internet-explorer-11-ab56f2d8-5d08-8f4b-8d70-4c0f8136da60)
- [Microsoft: Web page layout broken issue due to natural metrics in IE11](https://learn.microsoft.com/en-us/archive/blogs/asiatech/web-page-layout-broken-issue-due-to-natural-metrics-in-ie11)

これはブラウザだけの互換性問題ではありません。同じ種類の量子化が、Microsoft Word の文字幅、行高、表セル高、そして最終的な改ページ位置を決めています。

現代的なテキストスタックは、設計単位の小数値をなるべく保ったまま組版します。一方、Word の古いレイアウト経路には、途中で 96 DPI の整数ピクセルへ落とし、その結果を後段へ渡す箇所があります。したがって natural metrics をそのまま使うだけでは、見た目が近くても Word と同じページ分割にはなりません。必要なのは「GDIっぽい係数」ではなく、どの成分を、どの単位で、どの順番に、どちら向きへ丸めたかの再現です。

Rust + WebAssembly で DOCX を描画する Oxi では、この互換経路を測定から実装しました。本稿では、そのうち GDI の整数丸めに関係する部分を説明します。

## まずフォントサイズを ppem に落とす

最初の量子化はフォントサイズです。12pt の文字を 96 DPI で描くとき、em が何ピクセルになるかを求めます。

```text
ppem = round(font_size_pt × 96 / 72)
```

Rust では次の形です。

```rust
fn pixel_round(value_normalized: f32, ppem: f32) -> f32 {
    (value_normalized * ppem).round()
}

let ppem = (font_size * 96.0 / 72.0).round();
```

12pt なら `round(12 × 96 / 72) = 16` なので 16 ppem です。ここから先は、フォント内の比率へ 16 を掛け、GDI と同じ整数ピクセルへ量子化します。

重要なのは、最後の合計値だけを丸めないことです。アセンダーとディセンダーを足してから丸めた結果と、それぞれを整数化してから足した結果は一致しない場合があります。

## 欧文の行高は、成分ごとに整数化する

概念的に欲しい自然行高は、次の大きい方です。

```text
max(
    hhea_ascent + hhea_descent + hhea_lineGap,
    winAscent + winDescent
)
```

これは OpenType の `hhea` 全高と、`OS/2` テーブルの Windows メトリクス全高の比較です。しかし、これを浮動小数点のまま `max()` して最後に丸めても Word には一致しません。

実測と一致したのは、Windows メトリクスのアセンダーとディセンダーを別々に整数化し、`hhea` が上回る分も独立した external leading として整数化する経路でした。

```text
ppem        = round(fontSize × 96 / 72)
ascent_px   = round(winAscent × ppem)
descent_px  = round(winDescent × ppem)
hhea_excess = max(0, hhea_total − win_total)
leading_px  = round(hhea_excess × ppem)

line_height =
    (ascent_px + descent_px + leading_px) × 72 / 96
```

対応する Rust コードは次のようになります。

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

丸める前の代数では、

```text
win_total + max(0, hhea_total - win_total)
    = max(win_total, hhea_total)
```

です。しかし成分別に整数化した後は、両辺を同じ式だと扱えません。概念モデルは `max()`、GDI 互換の実装モデルは「成分別に量子化して加算」です。この違い自体が、丸め順序を仕様として扱う必要がある例です。

`72 / 96 = 0.75` は調整用のマジックナンバーではなく、96 DPI のピクセルを 72 DPI のポイントへ戻す単位変換です。

## 表セルではディセンダーだけ floor になる

さらに厄介だったのが、本文と表セルで量子化規則が同じではなかったことです。Word の表セル内の行ピッチを測ると、特定の経路ではディセンダーが四捨五入ではなく切り捨てになります。

```rust
pub fn word_line_height_table_cell(&self, font_size: f32) -> f32 {
    let ppem = (font_size * 96.0 / 72.0).round();
    let font_ascent = pixel_round(self.win_ascent, ppem);
    let font_descent = (self.win_descent * ppem).floor();
    (font_ascent + font_descent) * 72.0 / 96.0
}
```

Times New Roman 10pt は、`round` と `floor` が実際に分岐する例です。手元のフォントは UPM=2048、`winAscent=1825`、`winDescent=443` でした。

```text
ppem        = round(10 × 96 / 72) = 13
ascent_raw  = 1825 / 2048 × 13 = 11.5845...
descent_raw =  443 / 2048 × 13 =  2.8120...

ascent_px   = round(11.5845...) = 12
descent_px  = floor( 2.8120...) =  2
cell        = (12 + 2) × 72 / 96 = 10.5pt
```

ディセンダーを `round` すると 3px、セル高は 11.25pt になります。1 行で 0.75pt の差です。

Calibri 10.5pt も、現在手元にある Version 6.27 では分岐します。この版の `OS/2` 値は UPM=2048、`winAscent=1950`、`winDescent=550` です。

```text
ppem        = round(10.5 × 96 / 72) = 14
descent_raw = 550 / 2048 × 14 = 3.7598...
round       = 4
floor       = 3
```

なお、Calibri 11pt は同じフォント版で `550 / 2048 × 15 = 4.0283...` となり、`round` も `floor` も 4 です。Calibri の `1536/512` は `hhea` の ascent/descent であり、`OS/2` の `winAscent/winDescent` ではありません。フォント名だけでメトリクス値を固定せず、実際に使用したフォントファイルの版とテーブルを併記する必要があります。

20 行なら 0.75pt の差が 15pt になります。行単位では小さくても、ページネーションでは無視できません。

## 83/64 はどこから出したのか

MS Gothic と MS Mincho には、欧文用の式とは別の行間補正が必要でした。Oxi で使っている式は次です。

```text
raw = (winAscent + winDescent) / UPM × fontSize × 83 / 64
line_height = floor(raw × 8) / 8
```

ここで `83/64` は Microsoft の仕様書から拾った定数でも、フォントの `OS/2` や `hhea` にそのまま格納された値でもありません。私が Word の測定結果から作った、傾きの有理数表現です。

導出は次の手順でした。

1. `docGrid`、段落前後の余白、固定行間を外し、同じフォントとサイズの単行段落を連続させる。
2. Word COM の `Range.Information(6)` から隣接段落の Y 差を取り、フォントサイズを 0.5pt 刻みで sweep する。
3. MS Gothic / MS Mincho は `winAscent + winDescent = UPM = 256` なので、測定値をフォントサイズで割り、純粋な行間係数を取り出す。
4. 出力側で観測された `floor(x × 8) / 8` の 1/8pt 量子化をモデルへ入れる。
5. 2 の冪を分母に持つ係数を探索し、全サイズの量子化区間を同時に満たす値を選ぶ。

この探索で採用したのが、

```text
1.296875 = 83 / 64
```

でした。たとえば MS Gothic / MS Mincho は Windows メトリクス全高がちょうど 1em なので、量子化前後は次のようになります。

| サイズ | `fontSize × 83/64` | 1/8pt floor |
|---:|---:|---:|
| 10.5pt | 13.6171875pt | 13.5pt |
| 12pt | 15.5625pt | 15.5pt |
| 14pt | 18.15625pt | 18.125pt |

したがって、83 と 64 に OpenType 上のフィールド名が対応しているわけではありません。一方で「理由のない補正値」でもありません。測定 sweep、1/8pt の観測量子化、dyadic rational の探索から再導出できる、Oxi 側のモデル係数です。Microsoft 内部でも `83/64` という分数が使われている、とまでは主張しません。

この係数を `1.3` に丸めると、8.0〜72.0pt を 0.5pt 刻みで調べた 129 サイズのうち、107 サイズで 1/8pt floor の結果が変わります。

さらに面白いのは、係数をわずかに下側へずらした場合です。`1.2968` では、次の9サイズだけが正確に 1/8pt 下へ落ちました。

```text
8, 16, 24, 32, 40, 48, 56, 64, 72pt
```

これは偶然ではありません。

```text
fontSize × 83/64 × 8
    = fontSize × 83/8
```

0.5pt 刻みのサイズ集合では、この値が整数になるのは `fontSize` が 8pt の倍数のときです。つまり、この9サイズは `floor()` の直前でちょうど整数境界に乗ります。係数を少しでも下側へ近似すると境界の下へ落ち、結果が 1/8pt 小さくなります。

`83/64` を分数の形で残す理由は、さらにあります。

```text
83/64 = 1.296875₁₀ = 1.010011₂
```

分母が2の冪なので、`83/64` は `f32` で厳密に表現できます。0.5pt 刻みのフォントサイズも二進で厳密なため、この sweep の `fontSize × 83/64 × 8` は丸め誤差なしで評価できます。一方、`1.3` のような10進小数は二進では循環小数です。

8pt の倍数がまさに floor の knife edge に乗る以上、これは単なる表記上の好みではありません。測定で得た傾きを dyadic rational として保持することが、量子化境界を浮動小数点誤差から守る実装上の意味を持ちます。

また、`units_per_em == 256` なら常にこの経路、という判定にはしていません。MS PGothic のようなプロポーショナル CJK フォントや、異なる自然行高を持つフォントがあるため、Word 実測で分類したファミリーだけを通します。

## 文字幅は別の単位で量子化される

行高が GDI の整数ピクセルで説明できても、文字幅まで同じ式になるとは限りません。

13種類のフォント・サイズと 181 文字を Word で測ったところ、既知の欧文メトリクスからレイアウト幅を求める経路では、0.5pt、つまり 10twip 単位の量子化が一致しました。

```rust
let advance_em = metrics.char_width_em(c);

// 正のレイアウト幅を 10twip 単位で round-half-up する。
let width_tw =
    (advance_em * font_size * 20.0 / 10.0 + 0.5).floor() * 10.0;
let width_pt = width_tw / 20.0;
```

正の値だけを扱うため、この式は round-half-up を明示しています。`round()` へ置き換えることもできますが、この記事の主題は丸め方そのものなので、実装では「10twip へ half-up」という方針を名前付き helper に閉じ込める方が安全です。

一方、GDI hinting の結果が単純な OpenType 比率から再現できないフォントや文字は、ppem 別の実測値を使います。MS Gothic / MS Mincho の全角・半角はさらに別で、等幅の全角を `fontSize`、半角を `fontSize / 2` と扱う経路が Word と一致しました。

つまり「GDIを再現する」という名前でも、すべてを一つの `round()` へ統一してはいません。Word がその場面で使った量子化規則を再現します。

## 数式で閉じない部分は測定表にする

GDI の hinting によって、`tmHeight` や文字幅が単純な

```text
round(metric / UPM × ppem)
```

と一致しない組み合わせがあります。

たとえば Arial Narrow 10pt は 13 ppem です。Windows の `GetTextMetricsW` で測ると、

```text
tmHeight  = 16px
tmAscent  = 13px
tmDescent = 3px
```

でした。ポイントへ戻すと 12pt です。Word の 1.15 倍行間は `12 × 1.15 = 13.8pt` となり、COM 実測とも一致しました。フォントの `hhea` 値を直接使う経路では、この値に届きません。

Oxi では、この種の値をフォント名の `if` 文へ埋め込まず、Windows で再生成したテーブルとして分離しています。

```rust
// font -> ppem -> (height, ascent, descent)
gdi_heights: HashMap<String, HashMap<u32, (u32, u32, u32)>>,

// font -> ppem -> codepoint -> width_px
gdi_widths: HashMap<String, HashMap<u32, HashMap<u32, u32>>>,
```

これは説明用にそのまま抜いた素朴な格納形です。型付きキーや flat key に整理する余地はありますが、データ構造の洗練と測定値の意味は分けています。

## オラクルをどう作ったか

この調査で一番重要なのは、式より測定系です。Word の内部規則を直接読めるわけではないため、入力を一軸ずつ変え、複数の独立したオラクルを突き合わせました。

### Word COM: レイアウト位置

段落の行ピッチは、同一書式の単行段落を複数作り、各 `Paragraph.Range.Information(6)`、つまり `wdVerticalPositionRelativeToPage` の差から求めました。`ParagraphFormat.LineSpacing` と `LineSpacingRule` は設定値の確認には使いますが、Single の実レイアウト高そのものとはみなしません。

表セルでは、`Table.Cell(row, col).Range` を一文字ずつ走査して各文字の `Information(6)` を採取します。近接する Y を同じ行へまとめ、行レベル間の差の中央値をセル内の行ピッチとしました。これにより、セル全体の高さや余白を行高と取り違えずに済みます。

### Win32 GDI: 整数メトリクス

同じフォント名、サイズ、weight で `CreateFontIndirectW` した `HFONT` を DC へ選択し、`GetTextMetricsW` から `tmHeight`、`tmAscent`、`tmDescent`、`tmExternalLeading` を取得しました。文字幅は `GetCharWidth32W` と `GetTextExtentPoint32W` を使い、`GetTextFaceW` と `GetGlyphIndicesW` で意図しないフォールバックも検査しました。

### Word PDF: 描画結果

COM の位置情報だけでは、複雑な表や改ページ境界で参照フレームを誤ることがあります。そのため Word の `ExportAsFixedFormat` で PDF を作り、PyMuPDF で span の baseline、bbox、ページ番号を抽出しました。COM は構造とプロパティ、PDF は最終的な render truth、という役割分担です。

### sweep の再生成

フォント、サイズ、文字、本文／セル、grid の有無を manifest から DOCX へ展開し、Word COM と PDF 出力を自動実行します。結果は JSON/TSV に保存し、Rust 側の予測表と join して誤差を出します。13種類のフォント・サイズと181文字という数字は、手で拾ったサンプル数ではなく、この matrix の行数です。

フォントファイル自体はリポジトリへ入れません。Windows runner 上で測定表を再生成し、版、ppem、コードポイントをキーにした数値だけを格納します。

実装と測定パイプラインは公開しています。

- [Oxiソースリポジトリ](https://gitlab.com/Ryujiyasu/oxi)
- [ベンチマーク結果と測定方法](https://gitlab.com/Ryujiyasu/oxi/-/blob/main/README.md#layout-accuracy-vs-microsoft-word)
- [レイアウト再現度の技術ノート](https://gitlab.com/Ryujiyasu/oxi/-/blob/main/docs/layout_accuracy.md)

## 96 DPI と 150 DPI は別の話

本文で 96 DPI を使っているのは、Word 互換のレイアウト量子化を説明するためです。一方、Oxi の画像比較ベンチマークは 150 DPI でラスタライズします。

前者は「どこで改行・改ページするか」を決める互換計算、後者は「決まったページを何ピクセルの画像として比較するか」という評価解像度です。150 DPI で比較するからといって、レイアウト式の 96 を 150 に置き換えてはいけません。

## f32 で境界を扱ってよいのか

現在の Oxi の IR とレイアウト座標は `f32` なので、上の例も `f32` で書いています。通常のフォントサイズ範囲では、OpenType の整数値を正規化してすぐ規定単位へ snap し、境界ケースをテストで固定することで運用しています。

83/64 の経路は、その中でも安全性を具体的に説明できる例です。MS Gothic / MS Mincho ではメトリクス係数も `(winAscent + winDescent) / UPM = 256/256 = 1` です。この係数、83/64、0.5pt 刻みの入力がすべて二進で厳密なので、上で示した 8pt 倍数の整数境界を `f32` でも正確に表現できます。ここでは「測定値に近い小数」ではなく「測定区間を満たし、二進浮動小数点で厳密な分数」を選んだことが効いています。

ただし、これは任意の係数や任意の小数フォントサイズまで安全だという意味ではありません。丸め境界そのものが仕様である以上、一般論として「f32なら常に安全」とは言えません。とくにちょうど 0.5 に乗る入力の同点処理まで保証したい箇所は、元の整数メトリクスを保持したまま `MulDiv` 相当の整数演算にする方が強い設計です。Rust の `f32::round()` は halfway を 0 から遠ざかる向きへ丸めます。Word/GDI 側の規則が別なら、関数名とテストでその差を明示する必要があります。

## 実装で守った原則

最終的に、次の四つが重要でした。

1. 合計してから丸めず、Word/GDI が整数化する成分と順番を再現する。
2. `72/96`、20twip/pt、10twip のような単位変換・量子化単位と、測定から導いた係数を区別する。
3. `round`、`floor`、`ceil`、half-up を交換可能な近似だと思わず、文脈を含む仕様として名前を付ける。
4. 数式で再現できない hinting 結果は、適用範囲を限定した再生成可能な測定表にする。

GDI 互換実装で難しかったのは、Rust の `round()` や `floor()` の使い方ではありません。Word がどの場面でどの値を、どの単位へ、どの順番で、どちら向きに丸めたかを切り分けることでした。

Oxi はこの方法を含む互換実装により、2026年7月29日時点の blind 50 文書で、日本語は平均 SSIM 0.828・Word とページ数一致 44/50、英語は平均 SSIM 0.825・ページ数一致 48/50 です。[公開比較表](https://oxi-dd65f4.gitlab.io/#accuracy)では、正解を Microsoft Word（Microsoft 365 16.0.20131.20154）とし、Oxi、ONLYOFFICE 9.3.1.8、LibreOffice 26.2.1.2、SILURUS 0.72.2、BetterOffice 0.0.4、eigenpal 1.9.0 を150 DPIで比較しています。英語のページ数一致はこの6エンジン中で最多です。英語 SSIM は最初の 0.800、途中の 0.807 から改善しましたが、ページ内の画素配置では成熟したネイティブ実装にまだ負けています。

英語のページ数一致が日本語より高い一方で SSIM は低い、という結果は矛盾ではありません。二つは別の文書集合であり、ページ数一致は改ページという離散的な結果、SSIM はページ内の文字位置やラスタライズまで含む指標です。言語間の優劣として直接比較できる数字ではありません。

整数ピクセルの差を消すと、最後には一ページの差が消えます。
