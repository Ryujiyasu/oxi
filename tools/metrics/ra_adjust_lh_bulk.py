"""
Ra: 全テスト文書の adjustLineHeightInTable を XML と COM で比較
XML要素の有無 vs COM Compatibility(12) の対応を確定する
"""
import os
import json
import glob
import zipfile
import re
import subprocess
import sys
import time

DOCX_DIR = os.path.abspath(os.path.join(
    os.path.dirname(__file__), "..", "..",
    "tools", "golden-test", "documents", "docx"
))
OUT = os.path.abspath(os.path.join(
    os.path.dirname(__file__), "..", "..",
    "pipeline_data", "com_measurements", "adjust_lh_bulk.json"
))


def check_xml(docx_path):
    """XML内のadjustLineHeightInTableとcompatibilityModeを確認"""
    try:
        z = zipfile.ZipFile(docx_path)
        settings = z.read("word/settings.xml").decode("utf-8")

        has_adjust = "adjustLineHeightInTable" in settings

        # compatibilityMode
        compat_mode = None
        m = re.search(r'compatibilityMode.*?w:val="(\d+)"', settings)
        if m:
            compat_mode = int(m.group(1))

        # Check if adjustLineHeightInTable has explicit val attribute
        adjust_val = None
        m2 = re.search(r'adjustLineHeightInTable[^/]*?(?:w:val="([^"]*)")?', settings)
        if has_adjust:
            m3 = re.search(r'adjustLineHeightInTable[^>]*w:val="([^"]*)"', settings)
            if m3:
                adjust_val = m3.group(1)
            else:
                adjust_val = "implicit_true"  # present without val

        z.close()
        return has_adjust, adjust_val, compat_mode
    except Exception as e:
        return None, None, None


# Subprocess script for COM measurement
_COM_SCRIPT = r'''
import sys, json, win32com.client, pythoncom
pythoncom.CoInitialize()
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = 0
word.AutomationSecurity = 3
results = []
try:
    for path in sys.argv[1:]:
        try:
            doc = word.Documents.Open(path, ReadOnly=True, AddToRecentFiles=False, ConfirmConversions=False)
            try:
                adj = doc.Compatibility(12)
                results.append({"path": path, "com_adjust": adj})
            finally:
                doc.Close(SaveChanges=False)
        except Exception as e:
            results.append({"path": path, "error": str(e)})
finally:
    word.Quit()
    pythoncom.CoUninitialize()
print(json.dumps(results))
'''


def main():
    docx_paths = sorted(glob.glob(os.path.join(DOCX_DIR, "*.docx")))
    print(f"Scanning {len(docx_paths)} documents...")

    # Step 1: XML analysis (fast)
    xml_data = {}
    for p in docx_paths:
        doc_id = os.path.splitext(os.path.basename(p))[0]
        has_adjust, adjust_val, compat_mode = check_xml(p)
        xml_data[doc_id] = {
            "path": os.path.abspath(p),
            "xml_has_adjust": has_adjust,
            "xml_adjust_val": adjust_val,
            "xml_compat_mode": compat_mode,
        }

    # Summarize XML findings
    by_pattern = {}
    for doc_id, d in xml_data.items():
        key = (d["xml_has_adjust"], d["xml_adjust_val"], d["xml_compat_mode"])
        by_pattern.setdefault(key, []).append(doc_id)

    print("\n=== XML Patterns ===")
    for (has, val, mode), docs in sorted(by_pattern.items(), key=lambda x: -len(x[1])):
        print(f"  has={has} val={val} compat={mode}: {len(docs)} docs")
        for d in docs[:3]:
            print(f"    {d}")
        if len(docs) > 3:
            print(f"    ... +{len(docs)-3} more")

    # Step 2: COM measurement (batch, one per pattern)
    # Pick one representative doc from each XML pattern
    representatives = {}
    for (has, val, mode), docs in by_pattern.items():
        representatives[(has, val, mode)] = xml_data[docs[0]]["path"]

    print(f"\n=== COM Measurement ({len(representatives)} representative docs) ===")
    rep_paths = list(representatives.values())

    # Run COM in subprocess (batch)
    try:
        result = subprocess.run(
            [sys.executable, "-c", _COM_SCRIPT] + rep_paths,
            capture_output=True, text=True, encoding="utf-8", errors="replace",
            timeout=120,
        )
        if result.returncode == 0:
            com_results = json.loads(result.stdout)
        else:
            print(f"COM error: {result.stderr[:300]}")
            com_results = []
    except subprocess.TimeoutExpired:
        print("COM timeout!")
        com_results = []

    # Map COM results back
    com_map = {}
    for r in com_results:
        com_map[r["path"]] = r.get("com_adjust", r.get("error", "unknown"))

    print("\n=== XML vs COM Comparison ===")
    final = []
    for (has, val, mode), path in representatives.items():
        com_val = com_map.get(path, "no_data")
        count = len(by_pattern[(has, val, mode)])
        print(f"  XML: has={has} val={val} compat={mode} -> COM: {com_val}  ({count} docs)")
        final.append({
            "xml_has_adjust": has,
            "xml_adjust_val": val,
            "xml_compat_mode": mode,
            "com_adjust": com_val,
            "doc_count": count,
            "example": os.path.basename(path),
        })

    # Save
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump({"patterns": final, "xml_data": xml_data}, f, indent=2, ensure_ascii=False)
    print(f"\nSaved to {OUT}")


if __name__ == "__main__":
    main()
