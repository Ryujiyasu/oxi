"""Layer ablation analysis: identify which element types hurt SSIM most."""
import subprocess
import sys
import os
import numpy as np
from PIL import Image
from skimage.metrics import structural_similarity as ssim

def run_renderer(docx_path, output_prefix, dpi=150, exclude=None):
    cmd = [
        os.path.join(os.path.dirname(__file__), '..', 'oxi-gdi-renderer', 'target', 'release', 'oxi-gdi-renderer.exe'),
        docx_path, output_prefix, str(dpi)
    ]
    if exclude:
        cmd.append(f'--exclude={exclude}')
    subprocess.run(cmd, capture_output=True)

def render_word_pdf(docx_path, output_path, dpi=150):
    """Render Word reference via COM PDF export."""
    import win32com.client
    import fitz
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(os.path.abspath(docx_path), ReadOnly=True)
    pdf_path = output_path.replace('.png', '.pdf')
    doc.ExportAsFixedFormat(OutputFileName=os.path.abspath(pdf_path), ExportFormat=17, OpenAfterExport=False)
    page_count = doc.ComputeStatistics(2)
    doc.Close(False)
    word.Quit()

    pdf_doc = fitz.open(pdf_path)
    paths = []
    for i in range(pdf_doc.page_count):
        pix = pdf_doc[i].get_pixmap(matrix=fitz.Matrix(dpi/72, dpi/72))
        p = output_path.replace('.png', f'_p{i+1}.png')
        pix.save(p)
        paths.append(p)
    pdf_doc.close()
    return paths

def analyze(docx_path, dpi=150):
    tmp = os.path.join(os.environ.get('TEMP', '/tmp'), 'layer_ablation')
    os.makedirs(tmp, exist_ok=True)

    print(f"Analyzing: {os.path.basename(docx_path)}")

    # 1. Render Word reference
    print("  Rendering Word reference...")
    word_paths = render_word_pdf(docx_path, os.path.join(tmp, 'word.png'), dpi)

    # 2. Render Oxi full
    print("  Rendering Oxi full...")
    run_renderer(docx_path, os.path.join(tmp, 'oxi_full'), dpi)

    # 3. Ablation: exclude each type
    types = ['text', 'border', 'shading', 'box', 'image', 'clip']

    for t in types:
        print(f"  Rendering Oxi without {t}...")
        run_renderer(docx_path, os.path.join(tmp, f'oxi_no_{t}'), dpi, exclude=t)

    # 4. Compare SSIMs
    print("\n=== Layer Ablation Results ===")
    for page_idx, word_path in enumerate(word_paths):
        word_img = np.array(Image.open(word_path).convert('L'))

        oxi_full_path = os.path.join(tmp, f'oxi_full_p{page_idx+1}.png')
        if not os.path.exists(oxi_full_path):
            continue
        oxi_full = np.array(Image.open(oxi_full_path).convert('L'))

        # Match sizes
        h = min(word_img.shape[0], oxi_full.shape[0])
        w = min(word_img.shape[1], oxi_full.shape[1])
        word_img = word_img[:h, :w]
        oxi_full = oxi_full[:h, :w]

        full_ssim = ssim(oxi_full, word_img)
        print(f"\nPage {page_idx + 1}: Full SSIM = {full_ssim:.4f}")
        print(f"  {'Type':<10} {'Without SSIM':>12} {'Delta':>8}  Impact")
        print(f"  {'-'*50}")

        results = []
        for t in types:
            ablated_path = os.path.join(tmp, f'oxi_no_{t}_p{page_idx+1}.png')
            if not os.path.exists(ablated_path):
                continue
            ablated = np.array(Image.open(ablated_path).convert('L'))[:h, :w]
            abl_ssim = ssim(ablated, word_img)
            delta = abl_ssim - full_ssim
            results.append((t, abl_ssim, delta))

        results.sort(key=lambda x: -x[2])
        for t, abl_ssim, delta in results:
            impact = "← HURTING" if delta > 0.01 else ("← HELPING" if delta < -0.01 else "  neutral")
            print(f"  {t:<10} {abl_ssim:>12.4f} {delta:>+8.4f}  {impact}")

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Usage: python layer_ablation.py <input.docx> [dpi]")
        sys.exit(1)
    dpi = int(sys.argv[2]) if len(sys.argv) > 2 else 150
    analyze(sys.argv[1], dpi)
