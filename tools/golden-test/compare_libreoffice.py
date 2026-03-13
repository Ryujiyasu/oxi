#!/usr/bin/env python3
"""
Compare Oxi vs LibreOffice parse success rate.
LibreOffice test: try to convert each file to PDF using --headless --convert-to pdf.
If conversion succeeds, count as "parsed successfully".
"""
import json
import os
import subprocess
import sys
import time
from pathlib import Path

SOFFICE = r"C:\Program Files\LibreOffice\program\soffice.exe"
OOXML_EXTENSIONS = {".docx", ".xlsx", ".pptx"}
TIMEOUT = 60  # seconds per file


def find_files(directory):
    files = []
    for root, dirs, filenames in os.walk(directory):
        for f in filenames:
            ext = Path(f).suffix.lower()
            if ext in OOXML_EXTENSIONS:
                files.append(Path(root) / f)
    return sorted(files)


def test_libreoffice(filepath, tmp_dir):
    """Try to open/convert a file with LibreOffice. Returns (success, error_msg, time_ms)."""
    start = time.time()
    try:
        result = subprocess.run(
            [SOFFICE, "--headless", "--convert-to", "pdf", "--outdir", str(tmp_dir), str(filepath)],
            capture_output=True, text=True, timeout=TIMEOUT,
            env={**os.environ, "HOME": str(tmp_dir)}  # Avoid profile lock
        )
        elapsed_ms = int((time.time() - start) * 1000)

        # Check if PDF was created
        pdf_name = filepath.stem + ".pdf"
        pdf_path = tmp_dir / pdf_name
        if pdf_path.exists():
            pdf_path.unlink()  # Clean up
            return True, None, elapsed_ms
        else:
            stderr = result.stderr.strip()[:200] if result.stderr else "No PDF output"
            return False, stderr, elapsed_ms
    except subprocess.TimeoutExpired:
        elapsed_ms = int((time.time() - start) * 1000)
        return False, "Timeout", elapsed_ms
    except Exception as e:
        elapsed_ms = int((time.time() - start) * 1000)
        return False, str(e)[:200], elapsed_ms


def main():
    doc_dir = Path("./documents")
    tmp_dir = Path("./lo_tmp")
    tmp_dir.mkdir(exist_ok=True)

    if not Path(SOFFICE).exists():
        print(f"LibreOffice not found at {SOFFICE}")
        sys.exit(1)

    # Load Oxi results
    oxi_report_path = doc_dir / "golden_test_report.json"
    oxi_report = {}
    if oxi_report_path.exists():
        oxi_report = json.loads(oxi_report_path.read_text())

    files = find_files(doc_dir)
    total = len(files)
    print(f"=== LibreOffice Parse Test ===")
    print(f"Files: {total}")
    print(f"LibreOffice: {SOFFICE}")
    print()

    results = []
    counts = {"docx": [0, 0], "xlsx": [0, 0], "pptx": [0, 0]}  # [success, total]

    for i, f in enumerate(files):
        ext = f.suffix.lower().lstrip('.')
        success, error, ms = test_libreoffice(f, tmp_dir)
        status = "OK" if success else "FAIL"
        err_msg = f" ({error})" if error else ""
        print(f"[{i+1:>4}/{total:>4}] {status:4} {ext} {f.name[:55]} ({ms}ms){err_msg}")

        results.append({
            "filename": f.name,
            "format": ext,
            "success": success,
            "error": error,
            "time_ms": ms,
        })

        if ext in counts:
            counts[ext][1] += 1
            if success:
                counts[ext][0] += 1

    lo_success = sum(c[0] for c in counts.values())
    lo_total = sum(c[1] for c in counts.values())
    lo_rate = (lo_success / lo_total * 100) if lo_total > 0 else 0

    print()
    print("=" * 55)
    print("  Comparison: Oxi vs LibreOffice")
    print("=" * 55)
    print(f"  {'':20} {'Oxi':>10} {'LibreOffice':>12}")
    print(f"  {'Overall':20} {oxi_report.get('success_rate', 0):>9.1f}% {lo_rate:>11.1f}%")

    for fmt in ["docx", "xlsx", "pptx"]:
        oxi_fmt = oxi_report.get("by_format", {}).get(fmt, {})
        oxi_r = oxi_fmt.get("success_rate", 0) if oxi_fmt.get("total", 0) > 0 else "-"
        lo_s, lo_t = counts[fmt]
        lo_r = f"{lo_s/lo_t*100:.1f}" if lo_t > 0 else "-"
        oxi_str = f"{oxi_r:.1f}%" if isinstance(oxi_r, (int, float)) else oxi_r
        lo_str = f"{lo_r}%" if lo_r != "-" else "-"
        print(f"  {fmt.upper():20} {oxi_str:>10} {lo_str:>12}")
    print("=" * 55)

    # Save comparison report
    comparison = {
        "total_files": lo_total,
        "oxi_success_rate": oxi_report.get("success_rate", 0),
        "libreoffice_success_rate": lo_rate,
        "libreoffice_results": results,
        "by_format": {
            fmt: {
                "oxi": oxi_report.get("by_format", {}).get(fmt, {}).get("success_rate", 0),
                "libreoffice": (counts[fmt][0] / counts[fmt][1] * 100) if counts[fmt][1] > 0 else 0,
            }
            for fmt in ["docx", "xlsx", "pptx"]
        }
    }
    report_path = doc_dir / "comparison_report.json"
    report_path.write_text(json.dumps(comparison, indent=2))
    print(f"\nReport saved: {report_path}")

    # Cleanup
    import shutil
    shutil.rmtree(tmp_dir, ignore_errors=True)


if __name__ == "__main__":
    main()
