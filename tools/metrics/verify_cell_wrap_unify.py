"""Render bottom-5 + sample docs with current binary; compute SSIM vs baseline."""
import os, sys, subprocess, shutil, json
import numpy as np
from PIL import Image
from skimage.metrics import structural_similarity as ssim

RENDERER = r"tools\oxi-gdi-renderer\target\release\oxi-gdi-renderer.exe"
DOC_DIR = r"tools\golden-test\documents\docx"
WORD_DIR = r"pipeline_data\word_png"
OXI_DIR = r"pipeline_data\oxi_png"

BASELINE_JSON = r"pipeline_data\ssim_baseline.json"

# Bottom-5 + sample for blast radius check
TARGETS = [
    # Bottom-5
    "683fcab86e22_20230315_resources_data_contract_sample_02",
    "0e7af1ae8f21_20230331_resources_open_data_contract_sample_00",
    "d77a58485f16_20240705_resources_data_outline_08",
    "b35123fe8efc_tokumei_08_01",
    "b837808d0555_20240705_resources_data_guideline_02",
    # Sample docs with tables (blast radius check)
    "04b88e7e0b25_index-19",
    "2ea81a8441cc_0025006-192",
    "1ec1d05b4ab4_tokumei_08_02",
    "b9d8d9b58cbd_tokumei_08_06",
    "e3c545fac7a7_LOD_Handbook",
]


def render_one(stem):
    docx = os.path.join(DOC_DIR, f"{stem}.docx")
    if not os.path.exists(docx):
        return False
    oxi_dir = os.path.join(OXI_DIR, stem)
    if os.path.exists(oxi_dir):
        shutil.rmtree(oxi_dir)
    os.makedirs(oxi_dir)
    result = subprocess.run(
        [RENDERER, docx, os.path.join(oxi_dir, "page"), "96"],
        capture_output=True, text=True, timeout=180,
    )
    return result.returncode == 0


def compute_ssim(stem):
    """Returns dict {page: ssim}"""
    word_dir = os.path.join(WORD_DIR, stem)
    oxi_dir = os.path.join(OXI_DIR, stem)
    result = {}
    i = 1
    while True:
        wpath = os.path.join(word_dir, f"page_{i:04d}.png")
        opath = os.path.join(oxi_dir, f"page_p{i}.png")
        if not os.path.exists(wpath):
            break
        if not os.path.exists(opath):
            result[str(i)] = None
            i += 1
            continue
        try:
            w = np.array(Image.open(wpath).convert("L"))
            o = np.array(Image.open(opath).convert("L"))
            if w.shape != o.shape:
                oh = Image.open(opath).convert("L").resize(w.shape[::-1])
                o = np.array(oh)
            s = float(ssim(w, o))
            result[str(i)] = round(s, 4)
        except Exception as e:
            result[str(i)] = f"err: {e}"
        i += 1
    return result


def main():
    with open(BASELINE_JSON) as f:
        baseline = json.load(f)

    new_results = {}
    print(f"{'doc':70} {'page':>5} {'base':>7} {'new':>7} {'delta':>8}")
    print("-" * 100)
    total_changes = {"better": 0, "worse": 0, "same": 0}
    per_doc_deltas = []
    for stem in TARGETS:
        if not render_one(stem):
            print(f"[render failed] {stem}")
            continue
        new = compute_ssim(stem)
        new_results[stem] = new
        base = baseline.get(stem, {})
        doc_delta = 0.0
        worst_pg_delta = 0.0
        worst_pg = None
        for pg, s in sorted(new.items(), key=lambda x: int(x[0])):
            b = base.get(pg)
            if b is None or isinstance(s, str):
                continue
            delta = s - b
            doc_delta += delta
            mark = ""
            if abs(delta) > 0.005:
                mark = "**" if delta > 0 else "--"
            print(f"{stem:70} p{pg:>4} {b:7.4f} {s:7.4f} {delta:+8.4f} {mark}")
            if delta > 0.005:
                total_changes["better"] += 1
            elif delta < -0.005:
                total_changes["worse"] += 1
            else:
                total_changes["same"] += 1
            if delta < worst_pg_delta:
                worst_pg_delta = delta
                worst_pg = pg
        per_doc_deltas.append((stem, doc_delta, worst_pg, worst_pg_delta))

    print("\n=== Per-doc total delta ===")
    for stem, dd, wp, wd in per_doc_deltas:
        print(f"  {stem}: Δ={dd:+.4f}  worst page p{wp} ({wd:+.4f})")
    print(f"\nChange tally (per page, threshold 0.005): {total_changes}")

    # Bottom-N floor check
    print("\n=== Bottom-5 floor check ===")
    all_page_ssims_pre = []
    all_page_ssims_post = []
    for stem in baseline:
        for pg, s in baseline[stem].items():
            all_page_ssims_pre.append((s, stem, pg))
    # Use new_results if available, else baseline
    for stem, pages in baseline.items():
        for pg, s in pages.items():
            if stem in new_results and pg in new_results[stem] and not isinstance(new_results[stem][pg], str):
                all_page_ssims_post.append((new_results[stem][pg], stem, pg))
            else:
                all_page_ssims_post.append((s, stem, pg))
    all_page_ssims_pre.sort()
    all_page_ssims_post.sort()
    bottom5_pre = all_page_ssims_pre[:5]
    bottom5_post = all_page_ssims_post[:5]
    print("Pre-fix bottom-5:")
    for s, stem, pg in bottom5_pre:
        print(f"  {s:.4f} {stem} p{pg}")
    print("Post-fix bottom-5:")
    for s, stem, pg in bottom5_post:
        print(f"  {s:.4f} {stem} p{pg}")
    pre_sum = sum(s for s, _, _ in bottom5_pre)
    post_sum = sum(s for s, _, _ in bottom5_post)
    print(f"\nBottom-5 floor sum: pre={pre_sum:.4f} post={post_sum:.4f} delta={post_sum-pre_sum:+.4f}")

    with open("pipeline_data/cell_wrap_unify_verify.json", "w") as f:
        json.dump(new_results, f, indent=2)
    print("\n[OK] results saved")


if __name__ == "__main__":
    main()
