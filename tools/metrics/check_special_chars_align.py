"""Check the actual alignment Word uses for special_chars_spacing_01.docx."""
import win32com.client
import os
import sys

sys.stdout.reconfigure(encoding="utf-8")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False

DOCX = os.path.abspath("pipeline_data/docx/special_chars_spacing_01.docx")
doc = word.Documents.Open(DOCX, ReadOnly=True)
para = doc.Paragraphs(1)
print(f"Alignment value (raw): {para.Alignment}")
# wdAlignParagraphLeft=0, Center=1, Right=2, Justify=3, Distribute=4
# wdAlignParagraphJustifyMed=5, JustifyHi=7, JustifyLow=8
ALIGN_NAMES = {0:"Left", 1:"Center", 2:"Right", 3:"Justify", 4:"Distribute",
               5:"JustifyMed", 7:"JustifyHi", 8:"JustifyLow"}
print(f"Alignment name: {ALIGN_NAMES.get(para.Alignment, '?')}")
# Check Format properties
fmt = para.Format
print(f"FirstLineIndent: {fmt.FirstLineIndent}")
print(f"LeftIndent: {fmt.LeftIndent}")
print(f"RightIndent: {fmt.RightIndent}")
# CJK-specific
try:
    print(f"AutoAdjustRightIndent: {fmt.AutoAdjustRightIndent}")
except: pass
try:
    print(f"DisableLineHeightGrid: {fmt.DisableLineHeightGrid}")
except: pass
try:
    print(f"HalfWidthPunctuationOnTopOfLine: {fmt.HalfWidthPunctuationOnTopOfLine}")
except: pass
try:
    print(f"HangingPunctuation: {fmt.HangingPunctuation}")
except: pass
try:
    print(f"WordWrap: {fmt.WordWrap}")
except: pass
try:
    print(f"FarEastLineBreakControl: {fmt.FarEastLineBreakControl}")
except: pass
try:
    print(f"CharacterUnitFirstLineIndent: {fmt.CharacterUnitFirstLineIndent}")
except: pass

doc.Close(SaveChanges=False)
word.Quit()
