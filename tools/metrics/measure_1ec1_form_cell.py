"""Measure form cell content positions for 1ec1 (-1.9pt opposite shift hypothesis)."""
import sys
from PIL import Image

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

WORD = "pipeline_data/word_png/1ec1091177b1_006/page_0001.png"
OXI = "pipeline_data/oxi_png/1ec1091177b1_006/page_0001.png"


def measure_at(png, y_pt, x_min_pt, x_max_pt):
    img = Image.open(png).convert("L")
    w, h = img.size
    page_w = 595.30; page_h = 841.90
    pix = img.load()
    ymin = max(0, int((y_pt - 5) * h / page_h))
    ymax = min(h, int((y_pt + 5) * h / page_h))
    xmin = max(0, int(x_min_pt * w / page_w))
    xmax = min(w, int(x_max_pt * w / page_w))
    for x in range(xmin, xmax):
        for y in range(ymin, ymax):
            if pix[x, y] < 100:
                return x * page_w / w
    return None


# Body, form area, "○ 納付計画記載欄" header (y≈592), then form cells
TEST_POINTS = [
    ("body p27 ○納付計画 (y≈594)", 594, 35, 50),
    ("form header label (y≈617)", 617, 35, 60),  # 課税期間 cell
    ("form 課税期間 col2 (y≈617)", 617, 270, 290),  # right column
    ("form 期限内に金額 (y≈660)", 660, 35, 60),
    ("form 残額 (y≈678)", 678, 35, 60),
    ("body 'は、このチェック表' (y≈775)", 775, 35, 60),
]

for label, y, xmin, xmax in TEST_POINTS:
    word_x = measure_at(WORD, y, xmin, xmax)
    oxi_x = measure_at(OXI, y, xmin, xmax)
    diff = (oxi_x - word_x) if (word_x and oxi_x) else None
    print(f"{label}:")
    print(f"  Word: {word_x}")
    print(f"  Oxi:  {oxi_x}")
    print(f"  Diff: {diff:+.2f}" if diff is not None else "  Diff: N/A")
