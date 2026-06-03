# -*- coding: utf-8 -*-
# Triage 0e7af1ae8f21: structure counts + worst-page SSIM + blur test.
# ASCII-only output. Results written to JSON; no Japanese needles.
import json, re, zipfile, glob, os
import numpy as np
from PIL import Image, ImageFilter
from skimage.metrics import structural_similarity as ssim

DOCX = glob.glob(r"c:/Users/ryuji/oxi-main/tools/golden-test/documents/docx/0e7af1ae8f21*.docx")[0]
WORD_DIR = r"c:/Users/ryuji/oxi-main/pipeline_data/word_png/0e7af1ae8f21_20230331_resources_open_data_contract_sample_00"
OXI_PREFIX = r"c:/tmp/_tri_0e7af1ae8f21_p"  # _p<N>.png
OUT = r"c:/tmp/triage_0e7af_result.json"

res = {"doc_id": "0e7af1ae8f21", "docx": os.path.basename(DOCX)}

# --- 1. STRUCTURE ---
with zipfile.ZipFile(DOCX) as z:
    xml = z.read("word/document.xml").decode("utf-8", "replace")
res["textbox_count"] = len(re.findall(r"<w:txbxContent\b", xml))
res["table_count"] = len(re.findall(r"<w:tbl\b", xml))
m = re.search(r'<w:docGrid\b[^>]*w:type="([^"]+)"', xml)
res["docgrid_type"] = m.group(1) if m else (
    "present_no_type" if "<w:docGrid" in xml else "none")

# --- helpers ---
def load_gray(path):
    return np.asarray(Image.open(path).convert("L"), dtype=np.float64)

def load_pil_gray_resized(path, size):
    im = Image.open(path).convert("L")
    if im.size != size:
        im = im.resize(size, Image.LANCZOS)
    return im

def ssim_pair(word_path, oxi_path):
    w = Image.open(word_path).convert("L")
    o = Image.open(oxi_path).convert("L")
    if o.size != w.size:
        o = o.resize(w.size, Image.LANCZOS)
    wa = np.asarray(w, dtype=np.float64)
    oa = np.asarray(o, dtype=np.float64)
    return float(ssim(wa, oa, data_range=255.0))

# --- 2/3. WORST PAGE over all Word pages with an Oxi counterpart ---
word_pages = sorted(glob.glob(os.path.join(WORD_DIR, "page_*.png")))
per_page = {}
worst_page, worst_ssim = None, 2.0
for wp in word_pages:
    mm = re.search(r"page_(\d+)\.png$", wp)
    n = int(mm.group(1))
    oxi = f"{OXI_PREFIX}{n}.png"
    if not os.path.exists(oxi):
        per_page[n] = None
        continue
    s = ssim_pair(wp, oxi)
    per_page[n] = round(s, 4)
    if s < worst_ssim:
        worst_ssim, worst_page = s, n

res["per_page_ssim"] = per_page
res["worst_page"] = worst_page
res["worst_page_ssim"] = round(worst_ssim, 4)

# --- 4. BLUR TEST on worst page ---
wp = os.path.join(WORD_DIR, f"page_{worst_page:04d}.png")
op = f"{OXI_PREFIX}{worst_page}.png"
w_im = Image.open(wp).convert("L")
o_im = Image.open(op).convert("L")
if o_im.size != w_im.size:
    o_im = o_im.resize(w_im.size, Image.LANCZOS)

blur = {}
for r in (0, 2, 4, 6):
    if r == 0:
        wb, ob = w_im, o_im
    else:
        wb = w_im.filter(ImageFilter.GaussianBlur(r))
        ob = o_im.filter(ImageFilter.GaussianBlur(r))
    s = float(ssim(np.asarray(wb, dtype=np.float64),
                   np.asarray(ob, dtype=np.float64), data_range=255.0))
    blur[r] = round(s, 4)

res["blur_ssim"] = {str(k): v for k, v in blur.items()}
rise = blur[6] - blur[0]
res["blur_rise_r6_minus_r0"] = round(rise, 4)
res["blur_r0"] = blur[0]
position_capped = (rise > 0.12) and (blur[0] < 0.92)
res["position_capped"] = bool(position_capped)

# --- 5. SUITABILITY ---
tb = res["textbox_count"]
if not position_capped or tb > 8:
    suit = "poor"
elif tb <= 3:
    suit = "ideal"
else:
    suit = "ok"
res["emf_target_suitability"] = suit

with open(OUT, "w", encoding="utf-8") as f:
    json.dump(res, f, indent=2)
print("WROTE", OUT)
print("textbox_count", res["textbox_count"], "table_count", res["table_count"],
      "docgrid", res["docgrid_type"])
print("worst_page", worst_page, "worst_ssim", res["worst_page_ssim"])
print("blur", res["blur_ssim"], "rise", res["blur_rise_r6_minus_r0"])
print("position_capped", res["position_capped"], "suit", suit)
