# Word Text Shaping Metrics Pipeline

Word のテキストシェイピング（グリフ配置・文字幅補正）をブラックボックス観測し、
HarfBuzz との差分を数値化するパイプライン。

## 目的

Word は仕様書に記載されていない独自のテキスト補正を行っている。
このパイプラインは：

1. 既知の入力（docx）を Word に食わせて PDF を出力
2. PDF からグリフ座標を抽出
3. 同じフォント・サイズで HarfBuzz の出力と比較
4. 差分テーブルを生成 → Oxi のレイアウトエンジンに補正値として組み込む

## テストケース（36パターン）

| フォント | サイズ | 言語パターン |
|---|---|---|
| 游明朝 | 10.5pt / 11pt / 12pt | 日本語 / 英語 / 混在 |
| 游ゴシック | 10.5pt / 11pt / 12pt | 日本語 / 英語 / 混在 |
| Century | 10.5pt / 11pt / 12pt | 日本語 / 英語 / 混在 |
| Times New Roman | 10.5pt / 11pt / 12pt | 日本語 / 英語 / 混在 |

## 前提条件

- Windows 10/11
- Microsoft Word（Microsoft 365 または Office 2019 以降）
- Python 3.10+
- PowerShell 5.1+

## 実行方法

```powershell
# 一発実行
.\run_pipeline.ps1
```

または個別に：

```powershell
# 1. Python 依存パッケージをインストール
pip install -r requirements.txt

# 2. テスト用 docx を生成
python generate_test_docx.py

# 3. Word COM Automation で PDF に変換
.\word_to_pdf.ps1

# 4. メトリクス抽出・比較
python extract_metrics.py
```

## 出力

```
output/
├── pdfs/               # Word が生成した PDF（中間ファイル）
└── metrics_diff.json   # 差分テーブル（これが本体）
```

### metrics_diff.json の構造

```json
{
  "test_case": { "font": "游明朝", "size_pt": 10.5, "lang": "ja" },
  "max_diff_pt": 0.142,
  "avg_diff_pt": 0.067,
  "diffs": [
    {
      "char": "吾",
      "word_width": 10.512,
      "harfbuzz_width": 10.370,
      "diff": 0.142,
      "diff_pct": 1.37
    }
  ]
}
```

## 法的注意事項

- Word のバイナリを逆アセンブルしていない（クリーンルーム手法）
- 入力と出力のみを観測するブラックボックス解析
- Samba / Wine と同じアプローチ（相互運用目的のリバースエンジニアリング）
