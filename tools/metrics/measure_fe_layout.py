"""Test if useFELayout / enableOpenTypeFeatures cause the 0.5pt step pattern."""
import win32com.client
import os
import sys
import tempfile
import zipfile
import re

sys.stdout.reconfigure(encoding="utf-8")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False


def make_and_edit(text, font, sz_halfpt, settings_inject):
    doc = word.Documents.Add()
    ps = doc.PageSetup
    ps.PageWidth = 612.0
    ps.PageHeight = 792.0
    ps.LeftMargin = 90.0
    ps.RightMargin = 90.0
    ps.TopMargin = 72.0
    ps.BottomMargin = 72.0
    rng = doc.Range()
    rng.InsertAfter(text)
    rng = doc.Range()
    rng.Font.Name = font
    rng.Font.Size = sz_halfpt / 2.0
    doc.Paragraphs(1).Alignment = 0
    tmp = os.path.join(tempfile.gettempdir(), "fe_test.docx")
    if os.path.exists(tmp):
        os.remove(tmp)
    doc.SaveAs2(tmp, FileFormat=12)
    doc.Close(SaveChanges=False)

    tmp2 = tmp + ".edited.docx"
    with zipfile.ZipFile(tmp, "r") as zin:
        with zipfile.ZipFile(tmp2, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.namelist():
                data = zin.read(item)
                if item == "word/settings.xml":
                    s = data.decode("utf-8")
                    # Force doNotCompress
                    if "characterSpacingControl" in s:
                        s = re.sub(
                            r'<w:characterSpacingControl[^/]*/>',
                            '<w:characterSpacingControl w:val="doNotCompress"/>',
                            s,
                        )
                    else:
                        s = s.replace(
                            "</w:settings>",
                            '<w:characterSpacingControl w:val="doNotCompress"/></w:settings>',
                        )
                    # Inject extra settings inside <w:compat>
                    if settings_inject:
                        if "<w:compat>" in s:
                            s = s.replace("<w:compat>", "<w:compat>" + settings_inject)
                        elif "<w:compat/>" in s:
                            s = s.replace("<w:compat/>", "<w:compat>" + settings_inject + "</w:compat>")
                        else:
                            s = s.replace(
                                "</w:settings>",
                                "<w:compat>" + settings_inject + "</w:compat></w:settings>",
                            )
                    data = s.encode("utf-8")
                zout.writestr(item, data)
    return tmp, tmp2


def measure(text, font, sz_halfpt, settings_inject=""):
    tmp, tmp2 = make_and_edit(text, font, sz_halfpt, settings_inject)
    doc = word.Documents.Open(tmp2, ReadOnly=True)
    chars = doc.Range().Characters
    out = []
    prev_x = None
    prev_line = None
    for ci in range(1, chars.Count + 1):
        try:
            c = chars(ci)
            ch = c.Text
            if ch in ("\r", "\x07"):
                continue
            ln = c.Information(10)
            x = c.Information(5)
            dx = (x - prev_x) if (prev_x is not None and ln == prev_line) else None
            out.append((ch, ln, x, dx))
            prev_x = x
            prev_line = ln
        except Exception:
            pass
    doc.Close(SaveChanges=False)
    try:
        os.remove(tmp); os.remove(tmp2)
    except Exception:
        pass
    return out


def report(label, text, font, sz_halfpt, settings_inject=""):
    data = measure(text, font, sz_halfpt, settings_inject)
    line1 = [(ch, x, dx) for ch, ln, x, dx in data if ln == 1]
    from collections import Counter
    hist = Counter(round(dx, 2) for _, _, dx in line1 if dx is not None)
    width = (line1[-1][1] - line1[0][1] + (list(hist.keys())[0] if hist else 0)) if line1 else 0
    print(f"{label:50s}  L1={len(line1):2d}  width={width:6.2f}  dx={dict(hist)}")


# Baseline: no compat injection
TEXT = "漢" * 50
print("=== 漢×50 メイリオ 11pt ===")
report("baseline (no compat)", TEXT, "メイリオ", 22, "")
report("+useFELayout",         TEXT, "メイリオ", 22, "<w:useFELayout/>")
report("+useFELayout +OT",     TEXT, "メイリオ", 22,
       '<w:useFELayout/><w:compatSetting w:name="enableOpenTypeFeatures" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>')
report("+useFELayout +mode14", TEXT, "メイリオ", 22,
       '<w:useFELayout/><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/>')
report("+full special_chars compat", TEXT, "メイリオ", 22,
       '<w:useFELayout/>'
       '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/>'
       '<w:compatSetting w:name="overrideTableStyleFontSizeAndJustification" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>'
       '<w:compatSetting w:name="enableOpenTypeFeatures" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>'
       '<w:compatSetting w:name="doNotFlipMirrorIndents" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>')

print("\n=== special_chars text メイリオ 11pt ===")
SPECIAL = "特殊文字：①②③④⑤⑥⑦⑧⑨⑩　記号：★☆◆◇■□●○　単位：㎡㎏㎝　括弧：【】『』〈〉《》"
report("baseline", SPECIAL, "メイリオ", 22, "")
report("+full special_chars compat", SPECIAL, "メイリオ", 22,
       '<w:useFELayout/>'
       '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/>'
       '<w:compatSetting w:name="overrideTableStyleFontSizeAndJustification" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>'
       '<w:compatSetting w:name="enableOpenTypeFeatures" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>'
       '<w:compatSetting w:name="doNotFlipMirrorIndents" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>')

word.Quit()
