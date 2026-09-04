# フォント解決の macOS 実測

`crates/oxidocs-cli/FONT_CHECKS.md` の手順を macOS 実機で通した記録。
[font-resolution-linux.md](font-resolution-linux.md) と同じ先端・同じ blindC50 50冊を、
フォント事情が正反対の機械で回したもの。

SSIM の採点は含まない（真値が Word の PDF なので Windows が要る）。
ここにあるのはレンダ結果と、その構造・ink の検査。

| 項目 | 値 |
|---|---|
| OS | macOS 26.5.2 (25F84) / arm64 |
| cargo | 1.96.0 |
| python3 | 3.9.6（Xcode 同梱） |
| fontTools / PyMuPDF | **既定では両方とも無い**。検査のため venv に導入（fontTools 4.60.2） |
| 索引された face | 740 |
| 測ったもの | `feat/pdf-font-resolution` の `88d13003` |

**Linux とちょうど裏返しの機械。** Microsoft のコア書体が実物で入っている一方、
Office 由来の Calibri / Cambria / Segoe UI は無い。

## ビルドと試験

```
cargo build --release -p oxidocs-cli   → Finished in 1m 01s / rc=0 / error 0
cargo test  -p oxipdf-core             → 59 passed + 13 passed / 0 failed
cargo test  -p oxidocs-cli             → 2 passed / 0 failed
test writer::tests::every_composite_font_carries_its_font_file ... ok
```

Linux 固有・macOS 固有の分岐は踏んでいない。どちらでも同じ数の試験が通る。

## 索引 — 740 face、0.57 秒

既定クエリ 11 族 44 行の内訳は、解決 28 / `DEGRADED` 8 / `ABSENT` 8。

実物が見つかる族:

```
  Arial            /System/Library/Fonts/Supplemental/Arial.ttf#0        （4スタイル揃い）
  Times New Roman  …/Supplemental/Times New Roman.ttf#0                  （4スタイル揃い）
  Courier New      …/Supplemental/Courier New.ttf#0                      （4スタイル揃い）
  Verdana          …/Supplemental/Verdana.ttf#0                          （4スタイル揃い）
  Georgia          …/Supplemental/Georgia.ttf#0                          （4スタイル揃い）
  Tahoma           …/Supplemental/Tahoma.ttf#0                           （regular / bold のみ）
  Wingdings        …/Supplemental/Wingdings.ttf#0                        （regular のみ）
  Symbol           /System/Library/Fonts/Symbol.ttf#0                    （regular のみ）
```

`ABSENT` は `Calibri` / `Cambria` / `Segoe UI` の3族。前2つは同梱の Carlito / Caladea が受ける。

### `DEGRADED` が出るのはこの機械

`Tahoma italic` や `Wingdings bold` のように、族はあるがそのスタイルが無い場合に出る:

```
  Tahoma      italic       DEGRADED -> …/Supplemental/Tahoma.ttf (regular)
  Wingdings   bold         DEGRADED -> …/Supplemental/Wingdings.ttf (regular)
```

Linux では該当が無く、この出力形は macOS で初めて見えた。

### コレクション

システムに `.ttc` が 128 本ある。`Helvetica` / `Courier` / `Hiragino Sans` などは
コレクションとして索引に載る:

```
  Helvetica       /System/Library/Fonts/Helvetica.ttc#0
  Courier         /System/Library/Fonts/Courier.ttc#0
  Hiragino Sans   /System/Library/Fonts/ヒラギノ角ゴシック W8.ttc#0
```

ただし **`#1` 以降に解決した族は今回は 0 件**。コーパスが名指しした族はすべて
`#0`（コレクションの先頭か、単体の `.ttf`）に当たった。
`resolve_face_path` がコレクションの2番目以降を外す制限は、この機械では踏まない。

## blindC50 50冊

| | 冊数 | 所要 | PDF 合計 | 一時ファイル残 |
|---|---|---|---|---|
| fontTools あり (4.60.2) | 50/50 | **72 s** | **10.9 MB** | **0** |
| fontTools 無し | 50/50 | **27 s** | **61.8 MB** | **0** |

`TMPDIR` は macOS では `/tmp` ではなく `/var/folders/…/T`。
`std::env::temp_dir()` 化がここで効いていて、実 `TMPDIR` に残骸は出ない。

### 3つの検査

| 検査 | fontTools あり | fontTools 無し |
|---|---|---|
| `composite_ok` | **50/50 成立** | **50/50 成立** |
| `ink_rows == 0` | **0 冊** | **0 冊** |
| `type1 != 0` | **7 冊** | **7 冊** |
| `ink_rows` 合計 | 527,947 | 528,198 |

**subset の有無で未埋め込みの族は1つも変わらない。** 手順書が「見た目は変わらないはず」と
書いているとおりで、変わるのはサイズ（5.7倍）と所要時間だけ。

### 未埋め込みで残った 7 面

| 族名 | 面数 |
|---|---|
| `Noto Sans Symbols` | 2 |
| `Aptos` / `Cambria Math` / `Century Gothic` / `Garamond` / `Segoe UI` | 各1 |

すべて `Warning: no font file for X` として stderr に申告されている。
この機械に実在しない族ばかりで、手順書の立場ではこれが正しい状態。

`Symbol` と `Wingdings` は実物があるので、記号フォントの経路
（`Font subset (X) mapped no glyph; embedding the whole font instead`）を通って
丸ごと埋め込まれている——`FSymbol` 13回、`FSymbol_B` 2回、`FWingdings` 4回、`FArialBlack` 1回。

代替が使われたのは Calibri 36回、Cambria 6回、Calibri Light 2回。

## Linux との突き合わせ

同じ 50 冊、同じ `88d13003`。

| | Linux | macOS |
|---|---|---|
| 索引された face | 2,069 | 740 |
| `type1 != 0` の冊数 | 24 / 50 | **7 / 50** |
| 未埋め込みの面数 | 33 | **7** |
| `composite_ok` | 50/50 | 50/50 |
| 白紙 | 0 | 0 |
| 総頁数 | 650 | **650** |
| `ink_rows` 合計 | 528,125 | 527,947 |
| PDF 合計（subset 有） | 8.87 MB | 10.9 MB |

**頁数は 1 冊の違いもなく一致した。** レイアウトはフォントの当たり外れに依らない。
ink 合計の差も 0.03% で、最大の1冊でも 331 行（`technical__008ae1fa42620401`、
Linux 17,896 / Mac 17,565）。実際に違う書体で描かれている以上、この程度の差は出る。

未埋め込みの族は**補完関係**にある。macOS は Linux で落ちた
`Verdana` `Tahoma` `Arial Narrow` `Arial Black` `Trebuchet MS` `Comic Sans MS`
`Helvetica Neue` `Symbol` `Wingdings` をすべて実物で解決する。
逆に `Noto Sans Symbols` は Linux にあって macOS に無い。
両方に無いのは `Aptos` `Cambria Math` `Century Gothic` `Garamond` `Segoe UI` の5族。

## この機械で分かったこと

1. **`DEGRADED` の経路は macOS でしか踏まない。** 族はあるがスタイルが無い、という形は
   Microsoft 書体を実物で持つ機械に固有
2. **`TMPDIR` が `/tmp` でない環境でも残骸ゼロ。** `std::env::temp_dir()` 化の効き目が
   Linux より分かりやすく出る場所
3. **macOS には既定で fontTools も PyMuPDF も無い。** 手順書の検査を回すには
   venv などで足す必要がある（Xcode 同梱の python3 は 3.9.6）
