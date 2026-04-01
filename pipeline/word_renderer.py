"""Word COM API -> PNG renderer (Windows only)"""

import os
import sys
import time
import subprocess
from pathlib import Path
from .config import WORD_PNG_DIR, RENDER_DPI

# Max seconds per file before killing Word
RENDER_TIMEOUT = 30


def _kill_word():
    """残存 Word プロセスを強制終了（COM競合防止）"""
    subprocess.run(["taskkill", "/F", "/IM", "WINWORD.EXE"],
                   capture_output=True, timeout=10)
    time.sleep(1)


def render_with_word(docx_paths: list[str]) -> dict[str, list[str]]:
    """
    Word COM API で各.docxをページごとにPNG化する。
    ファイルごとにサブプロセスで実行し、タイムアウトで安定性確保。
    """
    # 開始前に残存Wordを掃除（Not Responding防止）
    _kill_word()

    results = {}

    for i, docx_path in enumerate(docx_paths):
        doc_id = Path(docx_path).stem
        out_dir = Path(WORD_PNG_DIR) / doc_id
        out_dir.mkdir(parents=True, exist_ok=True)

        # Skip if already rendered
        existing = sorted(out_dir.glob("page_*.png"))
        if existing:
            results[docx_path] = [str(p) for p in existing]
            continue

        try:
            result = subprocess.run(
                [sys.executable, "-c", _RENDER_SCRIPT,
                 os.path.abspath(docx_path), str(out_dir), str(RENDER_DPI)],
                capture_output=True, text=True, encoding="utf-8", errors="replace",
                timeout=RENDER_TIMEOUT,
            )
            if result.returncode == 0:
                pngs = sorted(out_dir.glob("page_*.png"))
                results[docx_path] = [str(p) for p in pngs]
                print(f"  Word: {doc_id} ({len(pngs)} pages) [{i+1}/{len(docx_paths)}]")
            else:
                print(f"[NG] Word error ({doc_id}): {result.stderr[:200]}")
                results[docx_path] = []
        except subprocess.TimeoutExpired:
            print(f"[NG] Word timeout ({doc_id}): >{RENDER_TIMEOUT}s")
            results[docx_path] = []
            subprocess.run(["taskkill", "/F", "/IM", "WINWORD.EXE"],
                           capture_output=True, timeout=10)
            time.sleep(1)
        except Exception as e:
            print(f"[NG] Word error ({doc_id}): {e}")
            results[docx_path] = []

    ok = sum(1 for v in results.values() if v)
    print(f"[OK] Word rendering done: {ok}/{len(results)} succeeded")
    return results


_RENDER_SCRIPT = r'''
import sys, os
docx_path, out_dir, dpi = sys.argv[1], sys.argv[2], int(sys.argv[3])

import win32com.client
import pythoncom
pythoncom.CoInitialize()
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False  # wdAlertsNone
word.AutomationSecurity = 3  # msoAutomationSecurityForceDisable (マクロ無効)
word.Options.UpdateLinksAtOpen = False  # リンク更新ダイアログ抑制
try:
    doc = word.Documents.Open(docx_path, ReadOnly=True, AddToRecentFiles=False, ConfirmConversions=False)
    try:
        page_count = doc.ComputeStatistics(2)
        for page_num in range(1, page_count + 1):
            pdf_path = os.path.join(out_dir, f"page_{page_num:04d}.pdf")
            png_path = os.path.join(out_dir, f"page_{page_num:04d}.png")
            doc.ExportAsFixedFormat(
                OutputFileName=pdf_path,
                ExportFormat=17, OpenAfterExport=False,
                OptimizeFor=0, Range=3, From=page_num, To=page_num,
            )
            import fitz
            d = fitz.open(pdf_path)
            zoom = dpi / 72
            pix = d[0].get_pixmap(matrix=fitz.Matrix(zoom, zoom))
            pix.save(png_path)
            d.close()
            os.unlink(pdf_path)
    finally:
        doc.Close(SaveChanges=False)
finally:
    word.Quit()
    pythoncom.CoUninitialize()
'''
