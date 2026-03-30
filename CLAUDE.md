# Oxi Development Guide

## Project Overview

Oxi is an OSS document processing suite built with Rust + WebAssembly.
The goal is to parse, render, and eventually edit .docx / .xlsx / .pptx files natively in the browser.

## Architecture

- **oxi-common**: Shared OOXML utilities (ZIP, XML, relationships)
- **oxidocs-core**: .docx engine — parser, IR, layout, font metrics
- **oxicells-core**: .xlsx engine — parser, IR
- **oxislides-core**: .pptx engine — parser, IR
- **oxi-wasm**: WebAssembly bindings via wasm-bindgen
- **web/**: React + Canvas demo app

## IR Design Principles

The Intermediate Representation (IR) must be language-agnostic and NOT depend on Word-specific internals.
Structure: Document → Page → Block (Paragraph | Table | Image) → Run

## Font Metrics

Font files are NEVER committed to the repository. Only pre-computed metrics tables are included.
Metrics are measured on GitHub Actions Windows runners and stored as data tables.

## Japanese Typography (Kinsoku)

Priority order:
1. Kinsoku processing (line-start/line-end prohibited characters)
2. Character spacing (justification)
3. Ruby (furigana)
4. Vertical writing (basics only)

Reference: JIS X 4051

## Testing

- Golden tests: render .docx with Oxi, compare pixel-by-pixel against Word screenshots
- Test fixtures go in tests/fixtures/
- CI: `cargo test`, `cargo clippy`, `wasm-pack build`

## Build Commands

```bash
cargo build                          # Build all
cargo test                           # Run tests
cargo clippy                         # Lint
cd crates/oxi-wasm && wasm-pack build --target web  # Wasm build
```

## Ra: Word仕様自動解析ループ

各セッション開始時に以下を確認し、自律的に仕様解析を継続する。

### ステート管理
- 仕様書: `docs/spec/word_layout_spec_ra.md`
- 計測データ: `pipeline_data/ra_manual_measurements.json`
- SSIMベースライン: `pipeline_data/ssim_baseline.json`

### 自律ループ手順
1. `docs/spec/word_layout_spec_ra.md` を読み、未解決の疑問を確認
2. 最も影響の大きい未解決疑問を選択
3. Python COM計測スクリプトを `tools/metrics/` に作成
4. 実行して結果を `pipeline_data/ra_manual_measurements.json` に追記
5. 結果を分析し、仕様書を更新
6. 仕様が確定したドメインは Rust 実装に反映
7. `python -m pipeline.verify` で SSIM リグレッションチェック
8. net positive なら commit、negative なら revert
9. 1 に戻る

### ドメイン状況（2026-03-28）
- **char_width**: フォールバック実装済み（MS UI Gothic）。現テスト文書では効果なし
- **page_break**: widow/orphan、keepNext/keepTogether 実装済み。段落途中改ページ修正済み（net +0.041）
- **spacing**: コラプス（max(sa,sb)）実装済み。net +0.71
- **line_height**: テーブルセル内リセット実装済み。net +0.66
- **grid_snap**: 実装済み
- **justify**: docDefaults jc=both 継承修正済み。Justify(均等割付)が全文書で有効化
- **SSIM: 0.7496 → 0.7884（+0.039）** ベースライン: 147文書399ページ
- **GDI幅オーバーライド**: 9フォント完全GDI幅テーブル組み込み済み（1055KB）
- **残りの改善余地**: 1ec文書72.7ptオーバーフロー、見出し行高さ、Desktop GDIレンダラー

### 計測テンプレート
行高さの正しい計測方法は「2段落のY座標差分」:
```python
y1 = doc.Paragraphs(1).Range.Information(6)  # wdVerticalPositionRelativeToPage
y2 = doc.Paragraphs(2).Range.Information(6)
gap = y2 - y1  # = line_height + spacing
```
`Format.LineSpacing` は設定値を返すだけで、実際のレンダリング高さではない。

### 重要ルール
- DLL解析禁止。COM API経由のブラックボックス測定のみ
- 推測で実装しない。必ずCOM実測で値を確定してから実装
- SSIMが下がる変更は revert（net positive ルール）

### 言い訳不可能な設計
Raは「言い訳できない環境」を前提としている。
- Wordのレイアウトは**決定論的**（同じ入力→同じ出力）
- COM APIで**全ての値が計測可能**（Y座標、行高さ、文字幅、段落間隔…）
- 差がある = 未実装の仕様がある = COM計測で特定 → 実装で解消
- **仕様は有限個。計測結果は永久資産。**一度計測すれば二度とやり直す必要がない
- 1つの仕様修正が複数文書を同時に改善する（収束する構造）
- 「できない」ではなく「まだやっていない」— 時間と計測回数の問題でしかない

## License

MIT. All third-party crate licenses must be MIT-compatible.
