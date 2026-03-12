#!/usr/bin/env python3
"""
Safe LibreOffice comparison - uses unique temp profiles per conversion
to avoid bootstrap.ini corruption and dialog blocking.
"""
import json, os, shutil, subprocess, sys, tempfile, time
from pathlib import Path

SOFFICE = r"C:\Program Files\LibreOffice\program\soffice.exe"
OOXML_EXTENSIONS = {".docx", ".xlsx", ".pptx"}
TIMEOUT = 45

def find_files(directory):
    files = []
    for root, dirs, filenames in os.walk(directory):
        for f in filenames:
            ext = Path(f).suffix.lower()
            if ext in OOXML_EXTENSIONS:
                files.append(Path(root) / f)
    return sorted(files)

def kill_soffice():
    """Kill any stuck soffice processes."""
    subprocess.run(["taskkill", "/f", "/im", "soffice.exe"],
                   capture_output=True, timeout=5)
    subprocess.run(["taskkill", "/f", "/im", "soffice.bin"],
                   capture_output=True, timeout=5)

def test_one(filepath):
    """Convert one file using a fresh temp profile."""
    start = time.time()
    tmp = tempfile.mkdtemp(prefix="lo_test_")
    profile_dir = os.path.join(tmp, "profile")
    out_dir = os.path.join(tmp, "out")
    os.makedirs(profile_dir)
    os.makedirs(out_dir)
    profile_url = "file:///" + profile_dir.replace("\\", "/")

    try:
        result = subprocess.run(
            [SOFFICE, "--headless", "--norestore", "--nologo",
             "--convert-to", "pdf",
             "--outdir", out_dir,
             f"-env:UserInstallation={profile_url}",
             str(filepath)],
            capture_output=True, text=True, timeout=TIMEOUT,
        )
        elapsed = int((time.time() - start) * 1000)
        pdf_path = os.path.join(out_dir, filepath.stem + ".pdf")
        if os.path.exists(pdf_path):
            return True, None, elapsed
        stderr = (result.stderr or "").strip()[:150]
        return False, stderr or "No PDF", elapsed
    except subprocess.TimeoutExpired:
        kill_soffice()
        return False, "Timeout", int((time.time() - start) * 1000)
    except Exception as e:
        return False, str(e)[:150], int((time.time() - start) * 1000)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

def main():
    doc_dir = Path("./documents")
    if not Path(SOFFICE).exists():
        print(f"LibreOffice not found: {SOFFICE}")
        sys.exit(1)

    oxi_report = {}
    rp = doc_dir / "golden_test_report.json"
    if rp.exists():
        oxi_report = json.loads(rp.read_text())

    files = find_files(doc_dir)
    total = len(files)
    print(f"=== LibreOffice Parse Test (Safe Mode) ===", flush=True)
    print(f"Files: {total}", flush=True)
    print(flush=True)

    counts = {"docx": [0, 0], "xlsx": [0, 0], "pptx": [0, 0]}
    results = []

    for i, f in enumerate(files):
        ext = f.suffix.lower().lstrip('.')
        ok, err, ms = test_one(f)
        status = "OK" if ok else "FAIL"
        err_msg = f" ({err})" if err else ""
        print(f"[{i+1:>4}/{total:>4}] {status:4} {ext} {f.name[:50]} ({ms}ms){err_msg}", flush=True)
        results.append({"filename": f.name, "format": ext, "success": ok, "error": err, "time_ms": ms})
        if ext in counts:
            counts[ext][1] += 1
            if ok:
                counts[ext][0] += 1
        # Small delay between conversions
        time.sleep(0.5)

    lo_success = sum(c[0] for c in counts.values())
    lo_total = sum(c[1] for c in counts.values())
    lo_rate = (lo_success / lo_total * 100) if lo_total > 0 else 0

    print(flush=True)
    print("=" * 60, flush=True)
    print("  Comparison: Oxi vs LibreOffice", flush=True)
    print("=" * 60, flush=True)
    print(f"  {'':20} {'Oxi':>10} {'LibreOffice':>13}", flush=True)
    print(f"  {'Overall':20} {oxi_report.get('success_rate', 0):>9.1f}% {lo_rate:>12.1f}%", flush=True)
    for fmt in ["docx", "xlsx", "pptx"]:
        oxi_fmt = oxi_report.get("by_format", {}).get(fmt, {})
        oxi_r = oxi_fmt.get("success_rate", 0) if oxi_fmt.get("total", 0) > 0 else 0
        s, t = counts[fmt]
        lo_r = (s / t * 100) if t > 0 else 0
        print(f"  {fmt.upper():20} {oxi_r:>9.1f}% {lo_r:>12.1f}%", flush=True)
    print("=" * 60, flush=True)
    print(f"  LibreOffice: {lo_success}/{lo_total} ({lo_rate:.1f}%)", flush=True)

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
    print(f"\nReport: {report_path}", flush=True)

if __name__ == "__main__":
    main()
