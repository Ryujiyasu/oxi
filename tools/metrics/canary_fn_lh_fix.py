"""13-doc focused canary for footnote line-height fix.

Before/after SSIM on bottom-bucket + a few control docs.
"""
import os, glob, subprocess
from PIL import Image
import numpy as np
from skimage.metrics import structural_similarity as ssim

CANARY_DOCS = [
    "b837808d0555_20240705_resources_data_guideline_02",
    "1ec1091177b1_006",
    "29dc6e8943fe_order_01",
    "34140b9c5662_index-14",
    "0e7af1ae8f21_20230331_resources_open_data_contract_sample_00",  # DWrite winner
    "459f05f1e877_kyodokenkyuyoushiki01",
    "04b88e7e0b25_index-19",
    "2ea81a8441cc_0025006-192",
    "3a4f9fbe1a83_001620506",
    # No-fn controls
    "683ffcab86e2_20230331_resources_open_data_contract_addon_00",
    "0e7af1ae8f21_20230331_resources_open_data_contract_sample_00",
]


def compute_ssim(doc):
    word_dir = f"pipeline_data/word_png/{doc}"
    oxi_dir = f"pipeline_data/oxi_png/{doc}"
    if not os.path.isdir(word_dir):
        return None
    pages = sorted(glob.glob(f"{word_dir}/page_*.png"))
    results = []
    for p in pages:
        page_num = int(os.path.basename(p).replace("page_", "").replace(".png", ""))
        # Try both naming conventions
        oxi_p = f"{oxi_dir}/page_p{page_num}.png"
        if not os.path.exists(oxi_p):
            oxi_p = f"{oxi_dir}/page_{page_num:04d}.png"
        if not os.path.exists(oxi_p):
            continue
        word = np.array(Image.open(p).convert("L"))
        oxi = Image.open(oxi_p).convert("L").resize(Image.open(p).size)
        s = ssim(word, np.array(oxi))
        results.append(s)
    return results


for doc in CANARY_DOCS:
    res = compute_ssim(doc)
    if res is None:
        print(f"{doc}: NOT FOUND")
        continue
    avg = sum(res) / len(res)
    print(f"{doc}: avg={avg:.4f} ({len(res)} pages: {[f'{x:.3f}' for x in res]})")
