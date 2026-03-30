# Oxi Word互換 SSIM自動計測パイプライン — AI実装指示書 v2

## 方針（重要）

このパイプラインは **計測・可視化まで** を自動化します。
修正はCOM API実測に基づいて人間が判断します。

```
Step 1: テスト文書生成       ← 自動（Claude API）
Step 2: Wordレンダリング     ← 自動（COM API → PNG）
Step 3: Oxiレンダリング      ← 自動（Canvas → PNG）
Step 4: SSIM計測             ← 自動
Step 5: ヒートマップ生成     ← 自動（人間が確認する）
-------------------------------------------------
Step 6: 修正                 ← 手動（COM実測ベース）
```

### なぜ自動修正しないか

- **no-speculation ルール**: Word仕様不明のまま推測で変更しない
- COM API実測に基づかない修正はClean-Room方針に反する
- AIが自動でRustを書き換えるのは品質管理上リスクが高い

---

## プロジェクト概要

- リポジトリ: https://gitlab.com/Ryujiyasu/oxi
- コア言語: Rust
- ターゲット: oxidocs-core（.docx エンジン）のレンダリング精度向上
- 正解の定義: Word 365（Windows・日本語ロケール）

---

## システム全体像

```
Claude 1
（テスト文書仕様生成）
      ↓
python-docx
（.docxファイル生成）
      ↓
       ├──→ Word COM API → PNG（正解画像）
       │
       └──→ Oxi Canvas  → PNG（現在の出力）
                ↓
          SSIM計測
                ↓
          ヒートマップ生成
          （Word | Oxi | 差分）
                ↓
          SSIMレポート出力
          （スコア低い順にソート）
                ↓
          👤 人間が確認して修正判断
```

---

## ディレクトリ構成

```
oxi/
├── pipeline/
│   ├── __init__.py
│   ├── main.py              # メインスクリプト
│   ├── generator.py         # Claude 1: テスト文書生成
│   ├── word_renderer.py     # Word COM API → PNG
│   ├── oxi_renderer.py      # Oxi Canvas → PNG
│   ├── ssim_calculator.py   # SSIM計測＋ヒートマップ
│   ├── reporter.py          # HTMLレポート生成
│   └── config.py            # 設定
├── pipeline_data/
│   ├── docx/                # テスト文書
│   ├── word_png/            # 正解画像
│   ├── oxi_png/             # Oxi出力画像
│   ├── heatmaps/            # 差分ヒートマップ
│   ├── reports/             # HTMLレポート
│   └── ssim_scores/         # スコア履歴JSON
└── requirements.txt
```

---

## ステップ1：環境セットアップ

### requirements.txt

```
anthropic>=0.25.0
pywin32>=306
Pillow>=10.0.0
scikit-image>=0.22.0
numpy>=1.26.0
matplotlib>=3.8.0
python-docx>=1.1.0
pdf2image>=1.17.0
jinja2>=3.1.0
```

### config.py

```python
# pipeline/config.py

import os

# Claude API（テスト文書生成のみに使用）
ANTHROPIC_API_KEY = os.environ["ANTHROPIC_API_KEY"]
CLAUDE_MODEL = "claude-sonnet-4-20250514"

# レンダリング設定
WORD_VERSION_TARGET = "Word 365 Windows JA"
RENDER_DPI = 150

# パス設定
OXI_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
DATA_DIR = os.path.join(OXI_ROOT, "pipeline_data")
DOCX_DIR       = os.path.join(DATA_DIR, "docx")
WORD_PNG_DIR   = os.path.join(DATA_DIR, "word_png")
OXI_PNG_DIR    = os.path.join(DATA_DIR, "oxi_png")
HEATMAP_DIR    = os.path.join(DATA_DIR, "heatmaps")
REPORTS_DIR    = os.path.join(DATA_DIR, "reports")
SSIM_SCORES_DIR = os.path.join(DATA_DIR, "ssim_scores")

# SSIMの警告閾値（これ以下のページをレポートに強調表示）
SSIM_WARN_THRESHOLD = 0.95
SSIM_ERROR_THRESHOLD = 0.90

# 1バッチのテスト文書数
BATCH_SIZE = 20
```

---

## ステップ2：テスト文書生成（Claude 1）

### generator.py

```python
# pipeline/generator.py

import anthropic
import json
from pathlib import Path
from datetime import datetime
from docx import Document
from docx.shared import Pt, Cm
from .config import *

client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

SYSTEM_PROMPT = """
あなたはWordのレイアウトエンジンのエキスパートQAエンジニアです。
Wordのレンダリングエンジンのバグを引き出しやすい.docxの構造を設計してください。

## 方針
- ターゲット: Word 365 Windows 日本語ロケール
- 1テストケース = 1つの要素に絞ったシンプルな文書
- 日本語文字列を積極的に使う

## 重点カテゴリ
1. フォントメトリクス — 行の高さ、ベースライン、行間
2. 禁則処理 — 行頭・行末禁則文字
3. 表のレイアウト — セル幅、結合セル、ボーダー
4. 画像の配置 — インライン、テキスト回り込み
5. ヘッダー・フッター — ページ番号
6. 段落スタイル — 見出し、インデント、スペーシング
7. 混在フォント — 日本語+英語の混在行
8. ページ境界 — ページをまたぐ段落・表

## 出力形式（JSON配列のみ）
[
  {
    "id": "一意のID（英数字とアンダースコアのみ）",
    "category": "カテゴリ名",
    "description": "何を検証するか",
    "difficulty": "LOW | MEDIUM | HIGH",
    "elements": [
      {
        "type": "paragraph | table | image | header | footer",
        "content": "要素の詳細仕様（日本語テキストを含める）",
        "style": "スタイル名（省略可）",
        "font": "フォント名（省略可）",
        "font_size": 数値（省略可）
      }
    ]
  }
]
"""

def generate_test_documents(count: int = BATCH_SIZE) -> list[str]:
    """テスト文書を生成して.docxファイルのパスリストを返す"""

    response = client.messages.create(
        model=CLAUDE_MODEL,
        max_tokens=8192,
        system=SYSTEM_PROMPT,
        messages=[{
            "role": "user",
            "content": (
                f"Wordレンダリングのエッジケースを {count} 件設計してください。"
                "日本語文書で発生しやすいレイアウトのズレに注目してください。"
            )
        }]
    )

    raw = response.content[0].text.strip()
    if "```json" in raw:
        raw = raw.split("```json")[1].split("```")[0].strip()
    elif "```" in raw:
        raw = raw.split("```")[1].split("```")[0].strip()

    specs = json.loads(raw)

    Path(DOCX_DIR).mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    generated_paths = []

    for i, spec in enumerate(specs):
        doc_id = spec.get("id", f"{timestamp}_{i:04d}")
        docx_path = str(Path(DOCX_DIR) / f"{doc_id}.docx")
        _build_docx(spec, docx_path)
        generated_paths.append(docx_path)
        print(f"  生成: {doc_id}.docx ({spec.get('category', '?')})")

    print(f"✓ テスト文書 {len(generated_paths)} 件を生成")
    return generated_paths


def _build_docx(spec: dict, output_path: str):
    """仕様からpython-docxで.docxファイルを構築する"""
    doc = Document()
    doc.styles['Normal'].font.name = '游ゴシック'
    doc.styles['Normal'].font.size = Pt(10.5)

    for element in spec.get("elements", []):
        el_type = element.get("type", "paragraph")
        content  = element.get("content", "")
        font_name = element.get("font", "游ゴシック")
        font_size = element.get("font_size", 10.5)

        if el_type == "paragraph":
            para = doc.add_paragraph()
            run = para.add_run(content)
            run.font.name = font_name
            run.font.size = Pt(font_size)

        elif el_type == "table":
            try:
                table_spec = json.loads(content) if isinstance(content, str) else content
                rows = table_spec.get("rows", 2)
                cols = table_spec.get("cols", 2)
                table = doc.add_table(rows=rows, cols=cols)
                table.style = 'Table Grid'
                for r, row_data in enumerate(table_spec.get("cells", [])[:rows]):
                    for c, cell_text in enumerate(row_data[:cols]):
                        table.cell(r, c).text = str(cell_text)
            except Exception:
                doc.add_paragraph(f"[table placeholder: {str(content)[:50]}]")

        elif el_type == "header":
            section = doc.sections[0]
            section.header.paragraphs[0].text = content

        elif el_type == "footer":
            section = doc.sections[0]
            section.footer.paragraphs[0].text = content

        else:
            doc.add_paragraph(content)

    doc.save(output_path)
```

---

## ステップ3：Word COM API → PNG

### word_renderer.py

```python
# pipeline/word_renderer.py
# Windowsのみで動作します

import os
import win32com.client
from pathlib import Path
from .config import *


def render_with_word(docx_paths: list[str]) -> dict[str, list[str]]:
    """
    Word COM APIで各.docxをページごとにPNG化する。
    戻り値: {docx_path: [page1.png, page2.png, ...]}
    """

    word = None
    results = {}

    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False

        for docx_path in docx_paths:
            doc_id  = Path(docx_path).stem
            out_dir = Path(WORD_PNG_DIR) / doc_id
            out_dir.mkdir(parents=True, exist_ok=True)

            doc = None
            try:
                doc = word.Documents.Open(
                    os.path.abspath(docx_path),
                    ReadOnly=True
                )
                page_count = doc.ComputeStatistics(2)  # wdStatisticPages
                png_paths  = []

                for page_num in range(1, page_count + 1):
                    pdf_path = str(out_dir / f"page_{page_num:04d}.pdf")
                    png_path = str(out_dir / f"page_{page_num:04d}.png")

                    # 1ページずつPDFに書き出す
                    doc.ExportAsFixedFormat(
                        OutputFileName=pdf_path,
                        ExportFormat=17,   # wdExportFormatPDF
                        OpenAfterExport=False,
                        OptimizeFor=0,     # wdExportOptimizeForPrint
                        Range=3,           # wdExportFromTo
                        From=page_num,
                        To=page_num,
                    )

                    _pdf_to_png(pdf_path, png_path)

                    if os.path.exists(pdf_path):
                        os.unlink(pdf_path)

                    png_paths.append(png_path)

                results[docx_path] = png_paths
                print(f"  Word: {doc_id} ({page_count}ページ)")

            except Exception as e:
                print(f"✗ Word エラー ({doc_id}): {e}")
                results[docx_path] = []
            finally:
                if doc:
                    doc.Close(SaveChanges=False)

    finally:
        if word:
            word.Quit()

    print(f"✓ Word レンダリング完了: {len(results)} 件")
    return results


def _pdf_to_png(pdf_path: str, png_path: str):
    """
    pdf2image + popplerでPDFをPNGに変換する。

    事前準備:
      pip install pdf2image
      popplerをWindowsにインストールしてPATHに追加
      → https://github.com/oschwartz10612/poppler-windows/releases
    """
    from pdf2image import convert_from_path
    images = convert_from_path(pdf_path, dpi=RENDER_DPI)
    if images:
        images[0].save(png_path, "PNG")
```

---

## ステップ4：Oxi Canvas → PNG

### oxi_renderer.py

```python
# pipeline/oxi_renderer.py

import subprocess
import os
from pathlib import Path
from .config import *


def render_with_oxi(docx_paths: list[str]) -> dict[str, list[str]]:
    """
    OxiのWASM APIで各.docxをページごとにPNG化する。
    戻り値: {docx_path: [page1.png, page2.png, ...]}
    """

    results = {}
    cli_path = Path(OXI_ROOT) / "tools/oxi-render-cli/index.js"

    if not cli_path.exists():
        raise NotImplementedError(
            "tools/oxi-render-cli/index.js を先に実装してください。\n"
            "仕様は下記「Node.js CLIの仕様」を参照。"
        )

    for docx_path in docx_paths:
        doc_id  = Path(docx_path).stem
        out_dir = Path(OXI_PNG_DIR) / doc_id
        out_dir.mkdir(parents=True, exist_ok=True)

        result = subprocess.run(
            ["node", str(cli_path),
             "--input",  os.path.abspath(docx_path),
             "--output", str(out_dir),
             "--dpi",    str(RENDER_DPI)],
            capture_output=True,
            text=True
        )

        if result.returncode != 0:
            print(f"✗ Oxi エラー ({doc_id}):\n{result.stderr[:300]}")
            results[docx_path] = []
            continue

        png_paths = sorted([
            str(out_dir / f)
            for f in os.listdir(str(out_dir))
            if f.endswith(".png")
        ])
        results[docx_path] = png_paths
        print(f"  Oxi: {doc_id} ({len(png_paths)}ページ)")

    print(f"✓ Oxi レンダリング完了: {len(results)} 件")
    return results
```

### Node.js CLIの仕様（tools/oxi-render-cli/index.js）

```
以下の仕様でindex.jsを実装してください。

【引数】
  --input  : 入力.docxファイルのパス
  --output : 出力ディレクトリ
  --dpi    : 解像度（デフォルト150）

【処理】
  1. oxi_wasm.jsをimport（crates/oxi-wasm/でwasm-pack buildしたもの）
  2. layout_document(bytes) を呼び出してpositioned elementsを取得
  3. node-canvasでCanvasを作成してelementsを描画
  4. ページごとにpage_0001.png, page_0002.png ... として保存

【注意】
  - フォントは游ゴシック・游明朝・メイリオをregisterFontで登録すること
  - Windowsのフォントパス: C:/Windows/Fonts/
  - npm install canvas で node-canvas をインストール
```

---

## ステップ5：SSIM計測＋ヒートマップ生成

### ssim_calculator.py

```python
# pipeline/ssim_calculator.py

import json
import numpy as np
from pathlib import Path
from datetime import datetime
from PIL import Image
from skimage.metrics import structural_similarity as ssim
import matplotlib
matplotlib.use("Agg")  # GUIなし環境用
import matplotlib.pyplot as plt
from .config import *


def calculate_ssim(
    word_results: dict[str, list[str]],
    oxi_results:  dict[str, list[str]]
) -> list[dict]:
    """
    SSIMを計測してスコアリストを返す（スコア低い順）。
    """

    scores = []

    for docx_path, word_pages in word_results.items():
        doc_id    = Path(docx_path).stem
        oxi_pages = oxi_results.get(docx_path, [])

        for page_idx, word_png in enumerate(word_pages):

            if page_idx >= len(oxi_pages):
                scores.append({
                    "doc_id":     doc_id,
                    "page":       page_idx + 1,
                    "ssim_score": 0.0,
                    "word_png":   word_png,
                    "oxi_png":    None,
                    "heatmap_png": None,
                    "diff_regions": [],
                    "error": "Oxiがこのページを生成できていない",
                })
                continue

            oxi_png  = oxi_pages[page_idx]
            word_img = _load_rgb(word_png)
            oxi_img  = _load_rgb(oxi_png)
            oxi_img  = _resize_to_match(oxi_img, word_img)

            score, diff_map = ssim(
                word_img, oxi_img,
                full=True,
                channel_axis=2,
                data_range=255
            )

            heatmap_path = _save_heatmap(
                doc_id, page_idx + 1,
                word_img, oxi_img, diff_map
            )

            diff_regions = _find_diff_regions(diff_map)

            scores.append({
                "doc_id":       doc_id,
                "page":         page_idx + 1,
                "ssim_score":   float(score),
                "word_png":     word_png,
                "oxi_png":      oxi_png,
                "heatmap_png":  heatmap_path,
                "diff_regions": diff_regions,
            })

            # コンソール表示（色分け）
            flag = "✓" if score >= SSIM_WARN_THRESHOLD else ("⚠" if score >= SSIM_ERROR_THRESHOLD else "✗")
            print(f"  {flag} SSIM: {doc_id} p.{page_idx+1} = {score:.4f}")

    scores.sort(key=lambda x: x["ssim_score"])

    # 保存
    Path(SSIM_SCORES_DIR).mkdir(parents=True, exist_ok=True)
    timestamp   = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = Path(SSIM_SCORES_DIR) / f"ssim_{timestamp}.json"
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(scores, f, ensure_ascii=False, indent=2)

    if scores:
        avg = sum(s["ssim_score"] for s in scores) / len(scores)
        low = sum(1 for s in scores if s["ssim_score"] < SSIM_WARN_THRESHOLD)
        print(f"✓ SSIM計測完了: 平均={avg:.4f} / 要改善={low}件")

    return scores


def _load_rgb(path: str) -> np.ndarray:
    return np.array(Image.open(path).convert("RGB"))


def _resize_to_match(img: np.ndarray, ref: np.ndarray) -> np.ndarray:
    h, w = ref.shape[:2]
    return np.array(Image.fromarray(img).resize((w, h), Image.LANCZOS))


def _save_heatmap(
    doc_id: str, page_num: int,
    word_img: np.ndarray,
    oxi_img:  np.ndarray,
    diff_map: np.ndarray
) -> str:
    """Word | Oxi | 差分ヒートマップ を横並びで保存する"""

    Path(HEATMAP_DIR).mkdir(parents=True, exist_ok=True)
    out_path = str(Path(HEATMAP_DIR) / f"{doc_id}_p{page_num:04d}.png")

    fig, axes = plt.subplots(1, 3, figsize=(18, 8))

    axes[0].imshow(word_img)
    axes[0].set_title("Word 365（正解）", fontsize=12)
    axes[0].axis("off")

    axes[1].imshow(oxi_img)
    axes[1].set_title("Oxi（現在）", fontsize=12)
    axes[1].axis("off")

    diff_gray = 1.0 - np.mean(diff_map, axis=2)
    im = axes[2].imshow(diff_gray, cmap="hot", vmin=0, vmax=1)
    axes[2].set_title("差分ヒートマップ（赤=大きな差異）", fontsize=12)
    axes[2].axis("off")
    plt.colorbar(im, ax=axes[2], fraction=0.046, pad=0.04)

    plt.suptitle(f"{doc_id}  page {page_num}", fontsize=14)
    plt.tight_layout()
    plt.savefig(out_path, dpi=100, bbox_inches="tight")
    plt.close()

    return out_path


def _find_diff_regions(diff_map: np.ndarray, threshold: float = 0.1) -> list[dict]:
    """差分が大きい領域を矩形として検出する（人間の確認用）"""
    from scipy import ndimage

    diff_gray = 1.0 - np.mean(diff_map, axis=2)
    mask = diff_gray > threshold
    if not mask.any():
        return []

    labeled, num_features = ndimage.label(mask)
    regions = []

    for label in range(1, min(num_features + 1, 10)):
        comp = labeled == label
        rows_any = np.any(comp, axis=1)
        cols_any = np.any(comp, axis=0)
        y_min, y_max = np.where(rows_any)[0][[0, -1]]
        x_min, x_max = np.where(cols_any)[0][[0, -1]]
        region_diff = diff_gray[y_min:y_max, x_min:x_max]
        regions.append({
            "x": int(x_min), "y": int(y_min),
            "width": int(x_max - x_min),
            "height": int(y_max - y_min),
            "max_diff":  float(region_diff.max()),
            "mean_diff": float(region_diff.mean()),
        })

    regions.sort(key=lambda r: r["mean_diff"], reverse=True)
    return regions
```

---

## ステップ6：HTMLレポート生成

### reporter.py

人間が確認しやすいHTMLレポートを生成します。

```python
# pipeline/reporter.py

import json
import shutil
from pathlib import Path
from datetime import datetime
from jinja2 import Template
from .config import *

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <title>Oxi SSIM レポート — {{ timestamp }}</title>
  <style>
    body { font-family: 'Segoe UI', sans-serif; background: #f5f5f5; margin: 0; padding: 20px; }
    h1 { color: #333; }
    .summary { background: white; padding: 16px; border-radius: 8px; margin-bottom: 24px; }
    .summary-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 16px; margin-top: 12px; }
    .metric { text-align: center; padding: 12px; border-radius: 6px; background: #f0f0f0; }
    .metric .value { font-size: 2em; font-weight: bold; }
    .metric .label { font-size: 0.85em; color: #666; margin-top: 4px; }
    .good  { background: #d4edda; } .good  .value { color: #155724; }
    .warn  { background: #fff3cd; } .warn  .value { color: #856404; }
    .error { background: #f8d7da; } .error .value { color: #721c24; }
    .page-card { background: white; border-radius: 8px; margin-bottom: 20px; overflow: hidden;
                 box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
    .page-header { padding: 12px 16px; display: flex; justify-content: space-between; align-items: center; }
    .page-header.good  { background: #d4edda; }
    .page-header.warn  { background: #fff3cd; }
    .page-header.error { background: #f8d7da; }
    .ssim-badge { font-size: 1.4em; font-weight: bold; }
    .page-title { font-size: 1em; font-weight: 600; }
    .heatmap img { width: 100%; display: block; }
    .regions { padding: 12px 16px; font-size: 0.85em; color: #555; }
    .error-msg { padding: 12px 16px; color: #721c24; background: #f8d7da; }
  </style>
</head>
<body>
  <h1>Oxi SSIM レポート</h1>
  <p>生成日時: {{ timestamp }} | ターゲット: {{ target }}</p>

  <div class="summary">
    <h2>サマリー</h2>
    <div class="summary-grid">
      <div class="metric {{ 'good' if avg_ssim >= 0.95 else ('warn' if avg_ssim >= 0.90 else 'error') }}">
        <div class="value">{{ "%.4f"|format(avg_ssim) }}</div>
        <div class="label">平均SSIMスコア</div>
      </div>
      <div class="metric good">
        <div class="value">{{ good_count }}</div>
        <div class="label">✓ 良好 (≥{{ warn_threshold }})</div>
      </div>
      <div class="metric warn">
        <div class="value">{{ warn_count }}</div>
        <div class="label">⚠ 要注意 (≥{{ error_threshold }})</div>
      </div>
      <div class="metric error">
        <div class="value">{{ error_count }}</div>
        <div class="label">✗ 要修正 (<{{ error_threshold }})</div>
      </div>
    </div>
  </div>

  {% for entry in scores %}
  {% set cls = 'good' if entry.ssim_score >= warn_threshold else ('warn' if entry.ssim_score >= error_threshold else 'error') %}
  <div class="page-card">
    <div class="page-header {{ cls }}">
      <span class="page-title">{{ entry.doc_id }}  page {{ entry.page }}</span>
      <span class="ssim-badge">SSIM: {{ "%.4f"|format(entry.ssim_score) }}</span>
    </div>
    {% if entry.error %}
    <div class="error-msg">{{ entry.error }}</div>
    {% elif entry.heatmap_png %}
    <div class="heatmap">
      <img src="{{ entry.heatmap_png_rel }}" alt="heatmap">
    </div>
    {% if entry.diff_regions %}
    <div class="regions">
      <strong>差分の大きい領域 (上位{{ entry.diff_regions|length }}件):</strong>
      {% for r in entry.diff_regions %}
        x={{ r.x }}, y={{ r.y }}, {{ r.width }}×{{ r.height }}px,
        mean_diff={{ "%.3f"|format(r.mean_diff) }}
      {% endfor %}
    </div>
    {% endif %}
    {% endif %}
  </div>
  {% endfor %}
</body>
</html>
"""

def generate_report(scores: list[dict]) -> str:
    """HTMLレポートを生成してパスを返す"""

    Path(REPORTS_DIR).mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    report_dir  = Path(REPORTS_DIR) / timestamp
    report_dir.mkdir(parents=True, exist_ok=True)
    report_path = str(report_dir / "index.html")

    # ヒートマップをレポートディレクトリにコピーして相対パスを設定
    for entry in scores:
        heatmap = entry.get("heatmap_png")
        if heatmap and Path(heatmap).exists():
            dest = report_dir / Path(heatmap).name
            shutil.copy2(heatmap, dest)
            entry["heatmap_png_rel"] = Path(heatmap).name
        else:
            entry["heatmap_png_rel"] = ""

    avg_ssim    = sum(s["ssim_score"] for s in scores) / len(scores) if scores else 0
    good_count  = sum(1 for s in scores if s["ssim_score"] >= SSIM_WARN_THRESHOLD)
    warn_count  = sum(1 for s in scores if SSIM_ERROR_THRESHOLD <= s["ssim_score"] < SSIM_WARN_THRESHOLD)
    error_count = sum(1 for s in scores if s["ssim_score"] < SSIM_ERROR_THRESHOLD)

    html = Template(HTML_TEMPLATE).render(
        timestamp       = datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        target          = WORD_VERSION_TARGET,
        scores          = scores,
        avg_ssim        = avg_ssim,
        good_count      = good_count,
        warn_count      = warn_count,
        error_count     = error_count,
        warn_threshold  = SSIM_WARN_THRESHOLD,
        error_threshold = SSIM_ERROR_THRESHOLD,
    )

    with open(report_path, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"✓ HTMLレポートを生成: {report_path}")
    return report_path
```

---

## ステップ7：メインスクリプト

### main.py

```python
# pipeline/main.py

import argparse
import os
from .generator       import generate_test_documents
from .word_renderer   import render_with_word
from .oxi_renderer    import render_with_oxi
from .ssim_calculator import calculate_ssim
from .reporter        import generate_report
from .config          import *


def run(batch_size: int = BATCH_SIZE):
    print("=" * 60)
    print("Oxi Word互換 SSIM計測パイプライン")
    print(f"ターゲット: {WORD_VERSION_TARGET}")
    print("=" * 60)

    print("\n[1/5] テスト文書生成中（Claude 1）...")
    docx_paths = generate_test_documents(count=batch_size)

    print("\n[2/5] Wordでレンダリング中...")
    word_results = render_with_word(docx_paths)

    print("\n[3/5] Oxiでレンダリング中...")
    oxi_results = render_with_oxi(docx_paths)

    print("\n[4/5] SSIM計測中...")
    ssim_scores = calculate_ssim(word_results, oxi_results)

    print("\n[5/5] レポート生成中...")
    report_path = generate_report(ssim_scores)

    print("\n" + "=" * 60)
    print(f"完了。レポートを確認してください:")
    print(f"  {report_path}")
    print("=" * 60)

    # ブラウザで開く
    os.startfile(report_path)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Oxi Word SSIM計測パイプライン")
    parser.add_argument("--batch-size", type=int, default=BATCH_SIZE,
                        help=f"テスト文書数（デフォルト: {BATCH_SIZE}）")
    args = parser.parse_args()
    run(batch_size=args.batch_size)
```

---

## 起動手順

```bash
# 1. 依存関係インストール
cd oxi
pip install -r pipeline/requirements.txt

# 2. popplerをWindowsにインストールしてPATHに追加
#    https://github.com/oschwartz10612/poppler-windows/releases

# 3. OxiのWASMビルド
cd crates/oxi-wasm
wasm-pack build --target nodejs
cd ../..

# 4. Node.js CLIの依存関係インストール
cd tools/oxi-render-cli
npm install canvas
cd ../..

# 5. 環境変数設定
set ANTHROPIC_API_KEY=your_api_key_here

# 6. 実行
python -m pipeline.main --batch-size 5
```

---

## 実装の優先順位

**優先度1（これがないと動かない）**
- `tools/oxi-render-cli/index.js` — OxiのCanvas→PNG変換CLI

**優先度2（コア機能）**
- `word_renderer.py` — popplerが正しくインストールされているか確認
- `ssim_calculator.py` — scipyが入っているか確認（`pip install scipy`）

**優先度3（仕上げ）**
- `reporter.py` — HTMLレポートの表示確認

---

## 将来：BugBounty統合

ユーザーが「このdocxがWordと違う」と報告したら自動でパイプラインに追加する。
修正判断は引き続き人間が行う。

```python
def import_bug_report(docx_path: str):
    """報告されたdocxをパイプラインに追加して即時計測する"""
    word_results = render_with_word([docx_path])
    oxi_results  = render_with_oxi([docx_path])
    scores       = calculate_ssim(word_results, oxi_results)
    report_path  = generate_report(scores)
    os.startfile(report_path)
```
