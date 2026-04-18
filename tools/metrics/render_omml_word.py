"""Render OMML fixture docx files via Word COM to PNG.

Uses CopyAsPicture + PlayEnhMetaFile flow (same as main pipeline) to
produce pixel-accurate Word renderings. Output PNGs saved alongside
Oxi outputs for visual/pixel comparison.

Output: pipeline_data/word_omml/{NAME}_p1.png
"""
import os, sys, time
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))
from pipeline.word_renderer import render_with_word

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

FIX_DIR = Path(__file__).resolve().parent.parent / "fixtures" / "omml_samples"
OUT_DIR = Path(__file__).resolve().parent.parent.parent / "pipeline_data" / "word_omml"
OUT_DIR.mkdir(parents=True, exist_ok=True)

# Override WORD_PNG_DIR from pipeline.config
from pipeline import config as pc
orig_dir = pc.WORD_PNG_DIR
pc.WORD_PNG_DIR = str(OUT_DIR)
# Also override Path-based constant
try:
    import pipeline.word_renderer as wr
    wr.WORD_PNG_DIR = str(OUT_DIR)
except Exception:
    pass

docx_paths = sorted(str(p) for p in FIX_DIR.glob("*.docx"))
print(f"Rendering {len(docx_paths)} fixture docx files via Word COM...")

results = render_with_word(docx_paths)
for docx, pngs in results.items():
    name = Path(docx).stem
    print(f"  {name}: {len(pngs)} page(s)")
    for p in pngs:
        print(f"    {p}")
