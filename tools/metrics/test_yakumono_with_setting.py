"""Create a test docx with characterSpacingControl=compressPunctuation
and verify Word renders it with yakumono compression, then verify
Oxi matches.
"""
import win32com.client
import os
import time
import sys
import subprocess

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False

# Create new doc and set the setting
doc = word.Documents.Add()
time.sleep(0.5)

# We will patch settings.xml after save

# Use a sample text with multiple consecutive punct
text = "テスト：「漢字」（読み）と「説明」（注意）が含まれます。"
rng = doc.Range()
rng.InsertAfter(text)
rng = doc.Range()
rng.Font.Name = "ＭＳ 明朝"
rng.Font.Size = 11.0
doc.Paragraphs(1).Alignment = 0

# Save with the setting
out_path = os.path.abspath("pipeline_data/docx_test/test_yakumono_compress.docx")
os.makedirs(os.path.dirname(out_path), exist_ok=True)
doc.SaveAs2(out_path, FileFormat=12)  # docx
doc.Close(SaveChanges=False)
word.Quit()

# Check the saved settings.xml
import zipfile
with zipfile.ZipFile(out_path) as z:
    s = z.read("word/settings.xml").decode("utf-8")
import re
m = re.search(r'characterSpacingControl[^/]*', s)
print("characterSpacingControl:", m.group(0) if m else "NOT FOUND")

# Now manually patch the XML to set compressPunctuation
import shutil
patched_path = os.path.abspath("pipeline_data/docx_test/test_yakumono_compressed.docx")
shutil.copy(out_path, patched_path)
# Reopen and modify settings.xml
import os
tmp_dir = "C:/tmp/yakumono_test_unzip"
if os.path.exists(tmp_dir):
    import shutil as sh
    sh.rmtree(tmp_dir)
os.makedirs(tmp_dir)

with zipfile.ZipFile(patched_path) as z:
    z.extractall(tmp_dir)

settings_path = os.path.join(tmp_dir, "word/settings.xml")
with open(settings_path, encoding="utf-8") as f:
    s = f.read()

# Replace doNotCompress with compressPunctuation
new_s = re.sub(
    r'characterSpacingControl w:val="[^"]*"',
    'characterSpacingControl w:val="compressPunctuation"',
    s
)
if new_s == s:
    # Insert before </w:settings>
    new_s = s.replace(
        '</w:settings>',
        '<w:characterSpacingControl w:val="compressPunctuation"/></w:settings>'
    )

with open(settings_path, "w", encoding="utf-8") as f:
    f.write(new_s)

# Re-zip
os.remove(patched_path)
with zipfile.ZipFile(patched_path, 'w', zipfile.ZIP_DEFLATED) as z:
    for root, dirs, files in os.walk(tmp_dir):
        for f in files:
            full = os.path.join(root, f)
            arc = os.path.relpath(full, tmp_dir).replace("\\", "/")
            z.write(full, arc)

print(f"Created: {patched_path}")

# Verify the setting
with zipfile.ZipFile(patched_path) as z:
    s = z.read("word/settings.xml").decode("utf-8")
m = re.search(r'characterSpacingControl[^/]*', s)
print("After patch:", m.group(0) if m else "NOT FOUND")
