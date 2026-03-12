#!/usr/bin/env python3
"""Fast SSIM comparison using native Rust layout + Pillow rendering.
Skips WASM build and Puppeteer — ~15s total vs ~4min."""
import subprocess
import sys
import time
from pathlib import Path

import cv2
import numpy as np
from PIL import Image, ImageDraw, ImageFont
from skimage.metrics import structural_similarity as ssim

REPO_ROOT = Path(__file__).resolve().parent.parent.parent
EXAMPLE_BIN = REPO_ROOT / "target" / "release" / "examples" / "layout_json"
OUTPUT_DIR = Path(__file__).resolve().parent / "pixel_output"
WORD_DIR = OUTPUT_DIR / "word"
OXI_FAST_DIR = OUTPUT_DIR / "oxi_fast"
DPI = 150
SCALE = DPI / 72.0  # points to pixels


def build_example():
    """Build the layout_json example in release mode."""
    print("Building layout_json (release)...", flush=True)
    t0 = time.time()
    r = subprocess.run(
        ["cargo", "build", "--release", "--example", "layout_json"],
        cwd=str(REPO_ROOT),
        capture_output=True, text=True, encoding='utf-8', errors='replace'
    )
    if r.returncode != 0:
        print(f"Build failed:\n{r.stderr}", file=sys.stderr)
        sys.exit(1)
    print(f"  Built in {time.time()-t0:.1f}s", flush=True)


def render_layout(docx_path: Path, out_png: Path):
    """Run native layout_json and render with Pillow."""
    r = subprocess.run(
        [str(EXAMPLE_BIN), str(docx_path)],
        capture_output=True, text=True, timeout=30, encoding='utf-8', errors='replace'
    )
    if r.returncode != 0:
        return False

    lines = r.stdout.strip().split('\n')
    pages = []
    current_page = None
    pending_text = None

    for line in lines:
        parts = line.split('\t')
        cmd = parts[0]

        if cmd == 'PAGE':
            if current_page:
                pages.append(current_page)
            w, h = float(parts[2]), float(parts[3])
            current_page = {
                'width': int(w * SCALE),
                'height': int(h * SCALE),
                'elements': []
            }
        elif cmd == 'TEXT':
            pending_text = {
                'type': 'text',
                'x': float(parts[1]) * SCALE,
                'y': float(parts[2]) * SCALE,
                'width': float(parts[3]) * SCALE,
                'height': float(parts[4]) * SCALE,
                'font_size': float(parts[5]),
                'font_family': parts[6],
                'bold': parts[7] == '1',
                'italic': parts[8] == '1',
                'underline': parts[9] == '1',
                'baseline_offset': float(parts[11]) * SCALE,
                'color': parts[12],
                'highlight': parts[13] if len(parts) > 13 else '',
            }
        elif cmd == 'T' and pending_text:
            pending_text['text'] = '\t'.join(parts[1:])  # rejoin in case text had tabs
            if current_page:
                current_page['elements'].append(pending_text)
            pending_text = None
        elif cmd == 'BORDER':
            if current_page:
                current_page['elements'].append({
                    'type': 'border',
                    'x1': float(parts[1]) * SCALE,
                    'y1': float(parts[2]) * SCALE,
                    'x2': float(parts[3]) * SCALE,
                    'y2': float(parts[4]) * SCALE,
                    'color': parts[5],
                    'width': float(parts[6]) * SCALE,
                })
        elif cmd == 'BG':
            if current_page:
                current_page['elements'].append({
                    'type': 'bg',
                    'x': float(parts[1]) * SCALE,
                    'y': float(parts[2]) * SCALE,
                    'width': float(parts[3]) * SCALE,
                    'height': float(parts[4]) * SCALE,
                    'color': parts[5],
                })

    if current_page:
        pages.append(current_page)

    if not pages:
        return False

    # Render page 0 only (matching golden test behavior)
    page = pages[0]
    img = Image.new('RGB', (page['width'], page['height']), 'white')
    draw = ImageDraw.Draw(img)

    # Font file mapping: family name -> (filename, ttc_index)
    FONT_MAP = {
        'calibri': ('calibri.ttf', 0),
        'calibri bold': ('calibrib.ttf', 0),
        'arial': ('arial.ttf', 0),
        'arial bold': ('arialbd.ttf', 0),
        'times new roman': ('times.ttf', 0),
        'times new roman bold': ('timesbd.ttf', 0),
        'ms mincho': ('msmincho.ttc', 0),
        'ms gothic': ('msgothic.ttc', 0),
        'ms pmincho': ('msmincho.ttc', 1),
        'ms pgothic': ('msgothic.ttc', 1),
        'yu gothic regular': ('YuGothR.ttc', 0),
        'yu gothic bold': ('YuGothB.ttc', 0),
        'yu mincho regular': ('yumin.ttf', 0),
        'yu mincho demibold': ('yumindb.ttf', 0),
        'century': ('CENTURY.TTF', 0),
        'cambria': ('cambria.ttc', 0),
        'hggothicm': ('HGRGM.TTC', 0),
        'hgpgothicm': ('HGRGM.TTC', 1),
        'hgsgothicm': ('HGRGM.TTC', 2),
        'hggothice': ('HGRGE.TTC', 0),
        'hgpgothice': ('HGRGE.TTC', 1),
        'hgsgothice': ('HGRGE.TTC', 2),
        'hgmincholb': ('HGRMB.TTC', 0),
        'hgminchole': ('HGRME.TTC', 0),
        'hgkyokashotai': ('HGRKK.TTC', 0),
        'hggyoshotai': ('HGRGY.TTC', 0),
    }
    # Aliases for Japanese font names
    FONT_ALIASES = {
        'ＭＳ 明朝': 'ms mincho', 'ＭＳ明朝': 'ms mincho', 'MS 明朝': 'ms mincho',
        'ＭＳ ゴシック': 'ms gothic', 'ＭＳゴシック': 'ms gothic', 'MS ゴシック': 'ms gothic',
        'ＭＳ Ｐ明朝': 'ms pmincho', 'ＭＳ Ｐゴシック': 'ms pgothic',
        '游ゴシック': 'yu gothic regular', '游明朝': 'yu mincho regular',
        'Yu Gothic': 'yu gothic regular', 'Yu Mincho': 'yu mincho regular',
        'HGｺﾞｼｯｸM': 'hggothicm', 'HGﾞｺｼｯｸM': 'hggothicm',
        'HGPｺﾞｼｯｸM': 'hgpgothicm', 'HGSｺﾞｼｯｸM': 'hgsgothicm',
        'HGｺﾞｼｯｸE': 'hggothice', 'HG丸ｺﾞｼｯｸM-PRO': 'hgsgothicm',
        'HG明朝B': 'hgmincholb', 'HG明朝E': 'hgminchole',
        'HG教科書体': 'hgkyokashotai', 'HG行書体': 'hggyoshotai',
    }

    # CJK Unicode ranges for font fallback detection
    def has_cjk(text):
        for ch in text:
            cp = ord(ch)
            if (0x3000 <= cp <= 0x9FFF or 0xF900 <= cp <= 0xFAFF
                    or 0xFF00 <= cp <= 0xFFEF or 0x20000 <= cp <= 0x2FA1F
                    or 0x25A0 <= cp <= 0x25FF  # Geometric Shapes (○●◎△□■)
                    or 0x2600 <= cp <= 0x26FF  # Miscellaneous Symbols (★☆)
                    or 0x2460 <= cp <= 0x24FF  # Enclosed Alphanumerics (①②③)
                    ):
                return True
        return False

    # Latin-only fonts that need CJK fallback
    LATIN_ONLY = {'arial', 'arial bold', 'calibri', 'calibri bold',
                  'times new roman', 'times new roman bold', 'century', 'cambria'}

    font_cache = {}
    def get_font(family, size_pt, bold=False, italic=False, cjk_fallback=False):
        # Include cjk_fallback in cache key so CJK text doesn't reuse Latin font
        lookup = FONT_ALIASES.get(family, family.lower())
        needs_cjk_redirect = cjk_fallback and lookup in LATIN_ONLY
        key = (family, size_pt, bold, italic, needs_cjk_redirect)
        if key in font_cache:
            return font_cache[key]
        px = max(1, round(size_pt * SCALE))
        font_dir = Path("C:/Windows/Fonts")

        # CJK fallback: if the font is Latin-only, use MS Gothic for CJK text
        if needs_cjk_redirect:
            lookup = 'ms gothic'

        if bold and not lookup.endswith(' bold'):
            bold_lookup = lookup + ' bold'
            if bold_lookup in FONT_MAP:
                lookup = bold_lookup

        # Try direct map
        if lookup in FONT_MAP:
            fname, idx = FONT_MAP[lookup]
            fp = font_dir / fname
            if fp.exists():
                try:
                    font = ImageFont.truetype(str(fp), px, index=idx)
                    font_cache[key] = font
                    return font
                except Exception:
                    pass

        # Fallback: try common patterns
        font = None
        fam_clean = family.lower().replace(' ', '')
        candidates = [f"{fam_clean}.ttf", f"{fam_clean}.ttc"]
        if bold:
            candidates = [f"{fam_clean}b.ttf", f"{fam_clean}bd.ttf"] + candidates

        for c in candidates:
            fp = font_dir / c
            if fp.exists():
                try:
                    font = ImageFont.truetype(str(fp), px)
                    break
                except Exception:
                    pass
        if font is None:
            # Fallback: use MS Gothic for CJK text, Calibri for Latin
            fallback_file = "msgothic.ttc" if cjk_fallback else "calibri.ttf"
            try:
                font = ImageFont.truetype(str(font_dir / fallback_file), px)
            except Exception:
                try:
                    font = ImageFont.truetype(str(font_dir / "calibri.ttf"), px)
                except Exception:
                    font = ImageFont.load_default()
        font_cache[key] = font
        return font

    for elem in page['elements']:
        if elem['type'] == 'bg':
            x, y = int(elem['x']), int(elem['y'])
            w, h = int(elem['width']), int(elem['height'])
            try:
                draw.rectangle([x, y, x+w, y+h], fill=elem['color'])
            except Exception:
                pass
        elif elem['type'] == 'border':
            x1, y1 = int(elem['x1']), int(elem['y1'])
            x2, y2 = int(elem['x2']), int(elem['y2'])
            lw = max(1, int(elem['width']))
            try:
                col = elem['color'] if elem['color'].startswith('#') else '#000000'
                draw.line([x1, y1, x2, y2], fill=col, width=lw)
            except Exception:
                draw.line([x1, y1, x2, y2], fill='black', width=lw)
        elif elem['type'] == 'text':
            text = elem['text']
            if not text.strip():
                continue
            need_cjk = has_cjk(text)
            font = get_font(elem['font_family'], elem['font_size'], elem['bold'], elem['italic'], cjk_fallback=need_cjk)
            x = round(elem['x'])
            y = round(elem['y'] + elem['baseline_offset'])
            # Pillow draws from top of glyph, adjust for ascent
            try:
                ascent, descent = font.getmetrics()
                y_draw = y - ascent
            except Exception:
                y_draw = y
            col = elem['color'] if elem['color'].startswith('#') else '#000000'
            try:
                draw.text((x, y_draw), text, fill=col, font=font)
            except Exception:
                pass
            if elem['underline']:
                tw = int(elem['width'])
                draw.line([x, y+2, x+tw, y+2], fill=col, width=1)

    img.save(str(out_png))
    return True


def compare_ssim(oxi_png: Path, word_png: Path):
    img1 = cv2.imread(str(oxi_png), cv2.IMREAD_GRAYSCALE)
    img2 = cv2.imread(str(word_png), cv2.IMREAD_GRAYSCALE)
    if img1 is None or img2 is None:
        return None
    h = min(img1.shape[0], img2.shape[0])
    w = min(img1.shape[1], img2.shape[1])
    img1 = cv2.resize(img1, (w, h))
    img2 = cv2.resize(img2, (w, h))
    score, _ = ssim(img1, img2, full=True)
    return round(score, 4)


def main():
    build_example()

    OXI_FAST_DIR.mkdir(parents=True, exist_ok=True)

    # Collect test files that have Word renders
    fixtures_dir = REPO_ROOT / "tests" / "fixtures"
    golden_dir = Path(__file__).resolve().parent / "documents" / "docx"
    word_stems = {p.stem for p in WORD_DIR.glob("*.png")} if WORD_DIR.exists() else set()

    test_files = []
    for d in [fixtures_dir, golden_dir]:
        if d.exists():
            for f in sorted(d.glob("*.docx")):
                if f.stem in word_stems:
                    test_files.append(f)

    print(f"\n=== Rendering {len(test_files)} files (native) ===", flush=True)
    t0 = time.time()
    ok = 0
    for i, f in enumerate(test_files):
        out = OXI_FAST_DIR / f"{f.stem}.png"
        if render_layout(f, out):
            ok += 1
            print(f"  [{i+1}/{len(test_files)}] {f.stem[:50]}... OK", flush=True)
        else:
            print(f"  [{i+1}/{len(test_files)}] {f.stem[:50]}... FAIL", flush=True)
    print(f"  Rendered: {ok}/{len(test_files)} in {time.time()-t0:.1f}s", flush=True)

    print(f"\n=== Comparing Oxi (fast) vs Word (SSIM) ===", flush=True)
    scores = []
    for word_png in sorted(WORD_DIR.glob("*.png")):
        oxi_png = OXI_FAST_DIR / word_png.name
        if not oxi_png.exists():
            continue
        score = compare_ssim(oxi_png, word_png)
        if score is not None:
            scores.append((score, word_png.stem))

    scores.sort()
    for s, f in scores:
        print(f"  {s:.4f}  {f[:60]}")
    if scores:
        avg = sum(s for s, _ in scores) / len(scores)
        above_98 = sum(1 for s, _ in scores if s >= 0.98)
        above_95 = sum(1 for s, _ in scores if s >= 0.95)
        above_90 = sum(1 for s, _ in scores if s >= 0.90)
        print(f"\nAverage SSIM: {avg:.4f} ({len(scores)} files)")
        print(f"  >= 0.98: {above_98}/{len(scores)}")
        print(f"  >= 0.95: {above_95}/{len(scores)}")
        print(f"  >= 0.90: {above_90}/{len(scores)}")


if __name__ == "__main__":
    main()
