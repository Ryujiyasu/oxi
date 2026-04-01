"""SSIM計測 + ヒートマップ生成"""

import json
import numpy as np
from pathlib import Path
from datetime import datetime
from PIL import Image
from skimage.metrics import structural_similarity as ssim
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from .config import (
    SSIM_SCORES_DIR, HEATMAP_DIR,
    SSIM_WARN_THRESHOLD, SSIM_ERROR_THRESHOLD,
)


def calculate_ssim(
    word_results: dict[str, list[str]],
    oxi_results:  dict[str, list[str]],
    skip_heatmap: bool = False,
) -> list[dict]:
    """SSIMを計測してスコアリストを返す（スコア低い順）。"""

    scores = []

    for docx_path, word_pages in word_results.items():
        doc_id    = Path(docx_path).stem
        oxi_pages = oxi_results.get(docx_path, [])

        for page_idx, word_png in enumerate(word_pages):

            if page_idx >= len(oxi_pages):
                scores.append({
                    "doc_id":      doc_id,
                    "page":        page_idx + 1,
                    "ssim_score":  0.0,
                    "word_png":    word_png,
                    "oxi_png":     None,
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

            if skip_heatmap:
                heatmap_path = None
                diff_regions = []
            else:
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

            flag = "[OK]" if score >= SSIM_WARN_THRESHOLD else (
                "[!!]" if score >= SSIM_ERROR_THRESHOLD else "[NG]")
            print(f"  {flag} SSIM: {doc_id} p.{page_idx+1} = {score:.4f}")

    scores.sort(key=lambda x: x["ssim_score"])

    Path(SSIM_SCORES_DIR).mkdir(parents=True, exist_ok=True)
    timestamp   = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = Path(SSIM_SCORES_DIR) / f"ssim_{timestamp}.json"
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(scores, f, ensure_ascii=False, indent=2)

    if scores:
        avg = sum(s["ssim_score"] for s in scores) / len(scores)
        low = sum(1 for s in scores if s["ssim_score"] < SSIM_WARN_THRESHOLD)
        print(f"[OK] SSIM計測完了: 平均={avg:.4f} / 要改善={low}件")

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
    axes[0].set_title("Word 365", fontsize=12)
    axes[0].axis("off")

    axes[1].imshow(oxi_img)
    axes[1].set_title("Oxi", fontsize=12)
    axes[1].axis("off")

    diff_gray = 1.0 - np.mean(diff_map, axis=2)
    im = axes[2].imshow(diff_gray, cmap="hot", vmin=0, vmax=1)
    axes[2].set_title("Diff heatmap", fontsize=12)
    axes[2].axis("off")
    plt.colorbar(im, ax=axes[2], fraction=0.046, pad=0.04)

    plt.suptitle(f"{doc_id}  page {page_num}", fontsize=14)
    plt.tight_layout()
    plt.savefig(out_path, dpi=100, bbox_inches="tight")
    plt.close()

    return out_path


def _find_diff_regions(diff_map: np.ndarray, threshold: float = 0.1) -> list[dict]:
    """差分が大きい領域を矩形として検出する"""
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
        if region_diff.size == 0:
            continue
        regions.append({
            "x": int(x_min), "y": int(y_min),
            "width": int(x_max - x_min),
            "height": int(y_max - y_min),
            "max_diff":  float(region_diff.max()),
            "mean_diff": float(region_diff.mean()),
        })

    regions.sort(key=lambda r: r["mean_diff"], reverse=True)
    return regions
