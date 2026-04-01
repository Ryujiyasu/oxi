"""HTMLレポート生成"""

import shutil
from pathlib import Path
from datetime import datetime
from jinja2 import Template
from .config import (
    REPORTS_DIR, WORD_VERSION_TARGET,
    SSIM_WARN_THRESHOLD, SSIM_ERROR_THRESHOLD,
)

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <title>Oxi SSIM Report — {{ timestamp }}</title>
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
  <h1>Oxi SSIM Report</h1>
  <p>{{ timestamp }} | Target: {{ target }}</p>

  <div class="summary">
    <h2>Summary</h2>
    <div class="summary-grid">
      <div class="metric {{ 'good' if avg_ssim >= 0.95 else ('warn' if avg_ssim >= 0.90 else 'error') }}">
        <div class="value">{{ "%.4f"|format(avg_ssim) }}</div>
        <div class="label">Avg SSIM</div>
      </div>
      <div class="metric good">
        <div class="value">{{ good_count }}</div>
        <div class="label">Good (>={{ warn_threshold }})</div>
      </div>
      <div class="metric warn">
        <div class="value">{{ warn_count }}</div>
        <div class="label">Warning (>={{ error_threshold }})</div>
      </div>
      <div class="metric error">
        <div class="value">{{ error_count }}</div>
        <div class="label">Error (<{{ error_threshold }})</div>
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
      <strong>Top diff regions:</strong>
      {% for r in entry.diff_regions %}
        x={{ r.x }}, y={{ r.y }}, {{ r.width }}x{{ r.height }}px,
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
    timestamp   = datetime.now().strftime("%Y%m%d_%H%M%S")
    report_dir  = Path(REPORTS_DIR) / timestamp
    report_dir.mkdir(parents=True, exist_ok=True)
    report_path = str(report_dir / "index.html")

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
    warn_count  = sum(1 for s in scores
                      if SSIM_ERROR_THRESHOLD <= s["ssim_score"] < SSIM_WARN_THRESHOLD)
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

    print(f"[OK] HTMLレポートを生成: {report_path}")
    return report_path
