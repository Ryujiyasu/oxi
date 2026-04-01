"""Oxi Word SSIM Pipeline"""

import argparse
import glob
import os
from .word_renderer   import render_with_word
from .oxi_renderer    import render_with_oxi
from .ssim_calculator import calculate_ssim
from .reporter        import generate_report
from .config          import WORD_VERSION_TARGET


def run(docx_dir: str, limit: int = 0):
    print("=" * 60)
    print("Oxi Word SSIM Pipeline")
    print(f"Target: {WORD_VERSION_TARGET}")
    print(f"Source: {docx_dir}")
    print("=" * 60)

    docx_paths = sorted(glob.glob(os.path.join(docx_dir, "*.docx")))
    if limit > 0:
        docx_paths = docx_paths[:limit]

    print(f"\n{len(docx_paths)} docx files found")

    print("\n[1/3] Rendering with Word...")
    word_results = render_with_word(docx_paths)

    print("\n[2/3] Rendering with Oxi...")
    oxi_results = render_with_oxi(docx_paths)

    print("\n[3/3] Calculating SSIM + report...")
    ssim_scores = calculate_ssim(word_results, oxi_results)
    report_path = generate_report(ssim_scores)

    print("\n" + "=" * 60)
    print(f"Done. Report: {report_path}")
    print("=" * 60)

    os.startfile(report_path)


if __name__ == "__main__":
    default_dir = os.path.join(
        os.path.dirname(__file__), "..",
        "tools", "golden-test", "documents", "docx"
    )
    parser = argparse.ArgumentParser(description="Oxi Word SSIM Pipeline")
    parser.add_argument("--docx-dir", type=str, default=default_dir,
                        help="Directory containing .docx files")
    parser.add_argument("--limit", type=int, default=0,
                        help="Max number of files to process (0=all)")
    args = parser.parse_args()
    run(docx_dir=os.path.abspath(args.docx_dir), limit=args.limit)
