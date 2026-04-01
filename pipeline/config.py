import os

# Claude API（テスト文書生成のみに使用）
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
CLAUDE_MODEL = "claude-sonnet-4-20250514"

# レンダリング設定
WORD_VERSION_TARGET = "Word 365 Windows JA"
RENDER_DPI = 150

# パス設定
OXI_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
DATA_DIR = os.path.join(OXI_ROOT, "pipeline_data")
DOCX_DIR        = os.path.join(DATA_DIR, "docx")
WORD_PNG_DIR    = os.path.join(DATA_DIR, "word_png")
OXI_PNG_DIR     = os.path.join(DATA_DIR, "oxi_png")
HEATMAP_DIR     = os.path.join(DATA_DIR, "heatmaps")
REPORTS_DIR     = os.path.join(DATA_DIR, "reports")
SSIM_SCORES_DIR = os.path.join(DATA_DIR, "ssim_scores")

# SSIMの警告閾値
SSIM_WARN_THRESHOLD  = 0.95
SSIM_ERROR_THRESHOLD = 0.90

# 1バッチのテスト文書数
BATCH_SIZE = 20
