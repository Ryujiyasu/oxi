"""4-column comparison: Word | Oxi Before | Oxi After | Heatmap
Usage: python tools/compare4.py <docx_name> [page]

Renders current Oxi, stashes changes, renders baseline Oxi,
then creates a 4-column comparison image.
"""
import sys, os, subprocess, shutil
from PIL import Image, ImageDraw
import fitz

def render_oxi(docx_path, out_pdf):
    r = subprocess.run(
        ['cargo', 'run', '--bin', 'oxi', '--', 'docx-to-pdf', docx_path, out_pdf],
        capture_output=True, text=True, encoding='utf-8', errors='replace', timeout=60
    )
    return os.path.exists(out_pdf)

def pdf_to_png(pdf_path, page_idx=0):
    doc = fitz.open(pdf_path)
    if page_idx >= doc.page_count:
        doc.close()
        return None
    pix = doc[page_idx].get_pixmap(matrix=fitz.Matrix(150/72, 150/72))
    png_path = pdf_path.replace('.pdf', f'_p{page_idx+1}.png')
    pix.save(png_path)
    doc.close()
    return png_path

def make_heatmap(img1, img2):
    """Simple pixel difference heatmap"""
    import numpy as np
    a1 = np.array(img1.convert('L'), dtype=float)
    a2 = np.array(img2.convert('L'), dtype=float)
    # Resize if needed
    if a1.shape != a2.shape:
        h = min(a1.shape[0], a2.shape[0])
        w = min(a1.shape[1], a2.shape[1])
        a1 = a1[:h, :w]
        a2 = a2[:h, :w]
    diff = np.abs(a1 - a2)
    # Normalize to 0-255
    diff = (diff / diff.max() * 255).astype(np.uint8) if diff.max() > 0 else diff.astype(np.uint8)
    return Image.fromarray(diff)

def main():
    if len(sys.argv) < 2:
        print("Usage: python tools/compare4.py <docx_name> [page]")
        sys.exit(1)

    docx_name = sys.argv[1]
    page = int(sys.argv[2]) - 1 if len(sys.argv) > 2 else 0

    # Find docx
    docx_path = f'tools/golden-test/documents/docx/{docx_name}'
    if not docx_path.endswith('.docx'):
        docx_path += '.docx'
    if not os.path.exists(docx_path):
        print(f'Not found: {docx_path}')
        sys.exit(1)

    tmp = os.environ.get('TEMP', '/tmp')
    after_pdf = os.path.join(tmp, 'compare4_after.pdf')
    before_pdf = os.path.join(tmp, 'compare4_before.pdf')

    # 1. Render "after" (current code)
    print('Rendering Oxi (after)...')
    render_oxi(docx_path, after_pdf)
    after_png = pdf_to_png(after_pdf, page)

    # 2. Word reference
    word_dir = f'pipeline_data/word_png/{docx_name.replace(".docx", "")}'
    word_png = os.path.join(word_dir, f'page_{page+1:04d}.png')
    if not os.path.exists(word_png):
        print(f'Word PNG not found: {word_png}')
        sys.exit(1)

    # 3. Render "before" (stash current changes, build, render, unstash)
    print('Rendering Oxi (before - baseline)...')
    # Check if there are changes to stash
    status = subprocess.run(['git', 'diff', '--quiet'], capture_output=True)
    has_changes = status.returncode != 0

    if has_changes:
        subprocess.run(['git', 'stash'], capture_output=True)
        subprocess.run(['cargo', 'build', '--bin', 'oxi'], capture_output=True, timeout=120)
        render_oxi(docx_path, before_pdf)
        subprocess.run(['git', 'stash', 'pop'], capture_output=True)
        subprocess.run(['cargo', 'build', '--bin', 'oxi'], capture_output=True, timeout=120)
    else:
        # No changes - before = after
        shutil.copy(after_pdf, before_pdf)

    before_png = pdf_to_png(before_pdf, page)

    # 4. Create 4-column image
    word_img = Image.open(word_png)
    after_img = Image.open(after_png) if after_png else Image.new('RGB', word_img.size, (255,255,255))
    before_img = Image.open(before_png) if before_png else Image.new('RGB', word_img.size, (255,255,255))

    # Resize all to same size
    w, h = word_img.size
    after_img = after_img.resize((w, h), Image.LANCZOS)
    before_img = before_img.resize((w, h), Image.LANCZOS)

    # Heatmap (after vs word)
    heatmap = make_heatmap(word_img, after_img)
    heatmap = heatmap.resize((w, h), Image.LANCZOS)

    # Combine
    gap = 10
    label_h = 25
    total_w = w * 4 + gap * 3
    combined = Image.new('RGB', (total_w, h + label_h), (255, 255, 255))

    labels = ['Word', 'Oxi (before)', 'Oxi (after)', 'Diff heatmap']
    for i, (img, label) in enumerate(zip([word_img, before_img, after_img, heatmap], labels)):
        x = i * (w + gap)
        combined.paste(img if img.mode == 'RGB' else img.convert('RGB'), (x, label_h))
        draw = ImageDraw.Draw(combined)
        draw.text((x + w // 2 - len(label) * 3, 5), label, fill='black')

    out_path = os.path.join(tmp, 'compare4_result.png')
    combined.save(out_path)
    print(f'Saved: {out_path}')

if __name__ == '__main__':
    main()
