#!/usr/bin/env python3
"""
Compare Oxi vs LibreOffice vs OnlyOffice parse success rate.
Tests each file by attempting PDF conversion with each tool.
"""
import json
import os
import shutil
import subprocess
import sys
import time
from pathlib import Path

SOFFICE = r"C:\Program Files\LibreOffice\program\soffice.exe"
# OnlyOffice Desktop Editors paths (try both)
ONLYOFFICE_PATHS = [
    r"C:\Program Files\ONLYOFFICE\DesktopEditors\DesktopEditors.exe",
    r"C:\Program Files (x86)\ONLYOFFICE\DesktopEditors\DesktopEditors.exe",
    os.path.expandvars(r"%LOCALAPPDATA%\ONLYOFFICE\DesktopEditors\DesktopEditors.exe"),
]
# OnlyOffice also has DocumentBuilder for conversion
ONLYOFFICE_BUILDER_PATHS = [
    r"C:\Program Files\ONLYOFFICE\DocumentBuilder\docbuilder.exe",
    r"C:\Program Files (x86)\ONLYOFFICE\DocumentBuilder\docbuilder.exe",
]

OOXML_EXTENSIONS = {".docx", ".xlsx", ".pptx"}
TIMEOUT = 60


def find_files(directory):
    files = []
    for root, dirs, filenames in os.walk(directory):
        for f in filenames:
            ext = Path(f).suffix.lower()
            if ext in OOXML_EXTENSIONS:
                files.append(Path(root) / f)
    return sorted(files)


def find_executable(paths):
    for p in paths:
        expanded = os.path.expandvars(p)
        if Path(expanded).exists():
            return expanded
    return None


def test_libreoffice(filepath, tmp_dir):
    """Convert with LibreOffice headless."""
    start = time.time()
    try:
        profile_dir = tmp_dir / "lo_profile"
        profile_dir.mkdir(exist_ok=True)
        out_dir = tmp_dir / "lo_out"
        out_dir.mkdir(exist_ok=True)
        profile_url = "file:///" + str(profile_dir).replace("\\", "/")
        result = subprocess.run(
            [SOFFICE, "--headless", "--convert-to", "pdf",
             "--outdir", str(out_dir),
             f"-env:UserInstallation={profile_url}",
             str(filepath)],
            capture_output=True, text=True, timeout=TIMEOUT,
        )
        elapsed_ms = int((time.time() - start) * 1000)
        pdf_path = out_dir / (filepath.stem + ".pdf")
        if pdf_path.exists():
            pdf_path.unlink()
            return True, None, elapsed_ms
        else:
            stderr = result.stderr.strip()[:200] if result.stderr else "No PDF output"
            return False, stderr, elapsed_ms
    except subprocess.TimeoutExpired:
        return False, "Timeout", int((time.time() - start) * 1000)
    except Exception as e:
        return False, str(e)[:200], int((time.time() - start) * 1000)


def test_onlyoffice(filepath, tmp_dir, exe_path):
    """Convert with OnlyOffice Desktop Editors."""
    start = time.time()
    try:
        out_dir = tmp_dir / "oo_out"
        out_dir.mkdir(exist_ok=True)
        pdf_path = out_dir / (filepath.stem + ".pdf")

        # OnlyOffice Desktop Editors conversion:
        # DesktopEditors.exe --convert-to:pdf --output-dir:<dir> <file>
        result = subprocess.run(
            [exe_path,
             f"--convert-to:pdf",
             f"--output-dir:{out_dir}",
             str(filepath)],
            capture_output=True, text=True, timeout=TIMEOUT,
        )
        elapsed_ms = int((time.time() - start) * 1000)

        if pdf_path.exists():
            pdf_path.unlink()
            return True, None, elapsed_ms

        # Also check for output with different naming
        pdfs = list(out_dir.glob("*.pdf"))
        if pdfs:
            for p in pdfs:
                p.unlink()
            return True, None, elapsed_ms

        stderr = result.stderr.strip()[:200] if result.stderr else "No PDF output"
        return False, stderr, elapsed_ms
    except subprocess.TimeoutExpired:
        return False, "Timeout", int((time.time() - start) * 1000)
    except Exception as e:
        return False, str(e)[:200], int((time.time() - start) * 1000)


def test_onlyoffice_builder(filepath, tmp_dir, builder_path):
    """Convert with OnlyOffice DocumentBuilder."""
    start = time.time()
    try:
        out_dir = tmp_dir / "oob_out"
        out_dir.mkdir(exist_ok=True)
        pdf_path = out_dir / (filepath.stem + ".pdf")

        # Create a builder script for conversion
        script = tmp_dir / "convert.docbuilder"
        abs_input = str(filepath).replace("\\", "/")
        abs_output = str(pdf_path).replace("\\", "/")
        script.write_text(
            f'builder.OpenFile("{abs_input}");\n'
            f'builder.SaveFile("pdf", "{abs_output}");\n'
            f'builder.CloseFile();\n'
        )

        result = subprocess.run(
            [builder_path, str(script)],
            capture_output=True, text=True, timeout=TIMEOUT,
        )
        elapsed_ms = int((time.time() - start) * 1000)

        if pdf_path.exists():
            pdf_path.unlink()
            return True, None, elapsed_ms
        stderr = result.stderr.strip()[:200] if result.stderr else "No PDF output"
        return False, stderr, elapsed_ms
    except subprocess.TimeoutExpired:
        return False, "Timeout", int((time.time() - start) * 1000)
    except Exception as e:
        return False, str(e)[:200], int((time.time() - start) * 1000)


def main():
    doc_dir = Path("./documents")
    tmp_dir = Path("./compare_tmp")
    tmp_dir.mkdir(exist_ok=True)

    # Detect available tools
    has_lo = Path(SOFFICE).exists()
    oo_exe = find_executable(ONLYOFFICE_PATHS)
    oob_exe = find_executable(ONLYOFFICE_BUILDER_PATHS)

    # Load Oxi results
    oxi_report_path = doc_dir / "golden_test_report.json"
    oxi_report = {}
    if oxi_report_path.exists():
        oxi_report = json.loads(oxi_report_path.read_text())

    files = find_files(doc_dir)
    total = len(files)

    print("=" * 65)
    print("  Oxi vs LibreOffice vs OnlyOffice - Parse Comparison")
    print("=" * 65)
    print(f"  Files: {total}")
    print(f"  LibreOffice:  {'FOUND' if has_lo else 'NOT FOUND'} ({SOFFICE})")
    print(f"  OnlyOffice:   {'FOUND' if oo_exe else 'NOT FOUND'}")
    if oob_exe:
        print(f"  OO Builder:   FOUND ({oob_exe})")
    print(f"  Oxi:          {oxi_report.get('success_rate', 'N/A')}% (from golden test)")
    print("=" * 65)
    print()

    if not has_lo and not oo_exe and not oob_exe:
        print("No comparison tools found. Install LibreOffice and/or OnlyOffice.")
        print()
        print("  LibreOffice: https://www.libreoffice.org/download/")
        print("  OnlyOffice:  https://www.onlyoffice.com/download-desktop.aspx")
        sys.exit(1)

    # Run tests
    lo_counts = {"docx": [0, 0], "xlsx": [0, 0], "pptx": [0, 0]}
    oo_counts = {"docx": [0, 0], "xlsx": [0, 0], "pptx": [0, 0]}
    lo_results = []
    oo_results = []

    for i, f in enumerate(files):
        ext = f.suffix.lower().lstrip('.')
        line = f"[{i+1:>4}/{total:>4}] {ext:4} {f.name[:45]:45}"

        # LibreOffice test
        lo_status = "-"
        if has_lo:
            lo_ok, lo_err, lo_ms = test_libreoffice(f, tmp_dir)
            lo_status = "OK" if lo_ok else "FAIL"
            lo_results.append({"filename": f.name, "format": ext, "success": lo_ok, "error": lo_err, "time_ms": lo_ms})
            if ext in lo_counts:
                lo_counts[ext][1] += 1
                if lo_ok:
                    lo_counts[ext][0] += 1

        # OnlyOffice test
        oo_status = "-"
        if oo_exe:
            oo_ok, oo_err, oo_ms = test_onlyoffice(f, tmp_dir, oo_exe)
            oo_status = "OK" if oo_ok else "FAIL"
            oo_results.append({"filename": f.name, "format": ext, "success": oo_ok, "error": oo_err, "time_ms": oo_ms})
            if ext in oo_counts:
                oo_counts[ext][1] += 1
                if oo_ok:
                    oo_counts[ext][0] += 1
        elif oob_exe:
            oo_ok, oo_err, oo_ms = test_onlyoffice_builder(f, tmp_dir, oob_exe)
            oo_status = "OK" if oo_ok else "FAIL"
            oo_results.append({"filename": f.name, "format": ext, "success": oo_ok, "error": oo_err, "time_ms": oo_ms})
            if ext in oo_counts:
                oo_counts[ext][1] += 1
                if oo_ok:
                    oo_counts[ext][0] += 1

        print(f"{line}  LO:{lo_status:4}  OO:{oo_status:4}")

    # Summary
    def calc_rate(counts):
        total = sum(c[1] for c in counts.values())
        success = sum(c[0] for c in counts.values())
        return (success / total * 100) if total > 0 else 0

    lo_rate = calc_rate(lo_counts) if has_lo else None
    oo_rate = calc_rate(oo_counts) if (oo_exe or oob_exe) else None

    print()
    print("=" * 65)
    print("  RESULTS")
    print("=" * 65)

    header = f"  {'':12}"
    header += f" {'Oxi':>10}"
    if has_lo:
        header += f" {'LibreOffice':>13}"
    if oo_exe or oob_exe:
        header += f" {'OnlyOffice':>12}"
    print(header)

    # Overall
    line = f"  {'Overall':12}"
    line += f" {oxi_report.get('success_rate', 0):>9.1f}%"
    if has_lo:
        line += f" {lo_rate:>12.1f}%"
    if oo_exe or oob_exe:
        line += f" {oo_rate:>11.1f}%"
    print(line)

    # Per format
    for fmt in ["docx", "xlsx", "pptx"]:
        oxi_fmt = oxi_report.get("by_format", {}).get(fmt, {})
        oxi_r = oxi_fmt.get("success_rate", 0) if oxi_fmt.get("total", 0) > 0 else 0
        line = f"  {fmt.upper():12} {oxi_r:>9.1f}%"
        if has_lo:
            s, t = lo_counts[fmt]
            r = (s / t * 100) if t > 0 else 0
            line += f" {r:>12.1f}%"
        if oo_exe or oob_exe:
            s, t = oo_counts[fmt]
            r = (s / t * 100) if t > 0 else 0
            line += f" {r:>11.1f}%"
        print(line)

    print("=" * 65)

    # Save report
    comparison = {
        "total_files": total,
        "oxi": {
            "success_rate": oxi_report.get("success_rate", 0),
            "by_format": {
                fmt: oxi_report.get("by_format", {}).get(fmt, {}).get("success_rate", 0)
                for fmt in ["docx", "xlsx", "pptx"]
            }
        },
    }
    if has_lo:
        comparison["libreoffice"] = {
            "success_rate": lo_rate,
            "by_format": {
                fmt: (lo_counts[fmt][0] / lo_counts[fmt][1] * 100) if lo_counts[fmt][1] > 0 else 0
                for fmt in ["docx", "xlsx", "pptx"]
            },
            "results": lo_results,
        }
    if oo_exe or oob_exe:
        comparison["onlyoffice"] = {
            "success_rate": oo_rate,
            "by_format": {
                fmt: (oo_counts[fmt][0] / oo_counts[fmt][1] * 100) if oo_counts[fmt][1] > 0 else 0
                for fmt in ["docx", "xlsx", "pptx"]
            },
            "results": oo_results,
        }

    report_path = doc_dir / "comparison_report.json"
    report_path.write_text(json.dumps(comparison, indent=2))
    print(f"\nReport saved: {report_path}")

    shutil.rmtree(tmp_dir, ignore_errors=True)


if __name__ == "__main__":
    main()
