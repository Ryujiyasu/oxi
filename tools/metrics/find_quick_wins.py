"""Find non-linesAndChars docs with low SSIM for quick wins."""
import json, zipfile, re, glob, os, sys
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

b = json.load(open("pipeline_data/ssim_baseline.json"))
results = []
for docx in sorted(glob.glob("tools/golden-test/documents/docx/*.docx")):
    name = os.path.splitext(os.path.basename(docx))[0]
    if name not in b: continue
    avg = sum(b[name].values()) / len(b[name])
    if avg > 0.85: continue
    try:
        with zipfile.ZipFile(docx) as z:
            doc = z.read("word/document.xml").decode("utf-8")
        g = re.search(r'docGrid[^/]*type="([^"]+)"', doc)
        grid_type = g.group(1) if g else "none"
    except:
        grid_type = "err"
    results.append((name, avg, len(b[name]), grid_type))

results.sort(key=lambda x: x[1])
print(f"Docs with SSIM < 0.85 ({len(results)} total):")
for name, avg, n, gt in results[:25]:
    print(f"  {avg:.4f} ({n:>2}p) [{gt:>15}] {name[:55]}")
