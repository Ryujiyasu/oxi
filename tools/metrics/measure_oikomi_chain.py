"""Measure oikomi (push-down) chain behavior with doNotCompress forced.

Tests how many chars Word backs up when multiple line-start-prohibited
chars cluster at the natural break point, and when the candidate fallback
char is itself problematic (e.g. line-end-prohibited).
"""
import win32com.client
import time
import sys
import os
import tempfile

sys.stdout.reconfigure(encoding="utf-8")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False

WD_FIRST_LINE = 10


def measure(text):
    """Create a doc with doNotCompress + ＭＳ 明朝 10.5pt + 432pt content,
    return list of (char, line_no)."""
    doc = word.Documents.Add()
    time.sleep(0.1)
    ps = doc.PageSetup
    ps.PageWidth = 612.0
    ps.PageHeight = 792.0
    ps.LeftMargin = 90.0
    ps.RightMargin = 90.0
    ps.TopMargin = 72.0
    ps.BottomMargin = 72.0
    # (forcing doNotCompress is done via the file round-trip path below)
    rng = doc.Range()
    rng.InsertAfter(text)
    rng = doc.Range()
    rng.Font.Name = "ＭＳ 明朝"
    rng.Font.Size = 10.5
    doc.Paragraphs(1).Alignment = 0
    # Disable kerning/compression on the document level
    # Word stores characterSpacingControl in document settings, not per-range.
    # We override via Document.Compatibility property where possible.
    # As a robust workaround, save→edit settings.xml→reopen.
    time.sleep(0.05)
    chars = doc.Range().Characters
    out = []
    for ci in range(1, chars.Count + 1):
        try:
            c = chars(ci)
            ch = c.Text
            if ch in ("\r", "\x07"):
                continue
            ln = c.Information(WD_FIRST_LINE)
            x = c.Information(5)
            out.append((ch, ln, x))
        except Exception:
            continue
    doc.Close(SaveChanges=False)
    return out


def measure_via_file(text):
    """Round-trip via file: create docx, edit settings.xml to set doNotCompress,
    reopen via COM, then measure."""
    import zipfile, shutil, re
    # Create base doc and save
    doc = word.Documents.Add()
    time.sleep(0.1)
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
    rng.Font.Name = "ＭＳ 明朝"
    rng.Font.Size = 10.5
    doc.Paragraphs(1).Alignment = 0
    tmp = os.path.join(tempfile.gettempdir(), "oikomi_test.docx")
    if os.path.exists(tmp):
        os.remove(tmp)
    doc.SaveAs2(tmp, FileFormat=12)  # wdFormatXMLDocument
    doc.Close(SaveChanges=False)

    # Edit settings.xml to inject doNotCompress
    tmp2 = tmp + ".edited.docx"
    with zipfile.ZipFile(tmp, "r") as zin:
        with zipfile.ZipFile(tmp2, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.namelist():
                data = zin.read(item)
                if item == "word/settings.xml":
                    s = data.decode("utf-8")
                    if "characterSpacingControl" in s:
                        s = re.sub(
                            r'<w:characterSpacingControl[^/]*/>',
                            '<w:characterSpacingControl w:val="doNotCompress"/>',
                            s,
                        )
                    else:
                        # Insert before </w:settings>
                        s = s.replace(
                            "</w:settings>",
                            '<w:characterSpacingControl w:val="doNotCompress"/></w:settings>',
                        )
                    data = s.encode("utf-8")
                zout.writestr(item, data)

    # Reopen and measure
    doc = word.Documents.Open(tmp2, ReadOnly=True)
    chars = doc.Range().Characters
    out = []
    for ci in range(1, chars.Count + 1):
        try:
            c = chars(ci)
            ch = c.Text
            if ch in ("\r", "\x07"):
                continue
            ln = c.Information(WD_FIRST_LINE)
            x = c.Information(5)
            out.append((ch, ln, x))
        except Exception:
            continue
    doc.Close(SaveChanges=False)
    try:
        os.remove(tmp)
        os.remove(tmp2)
    except Exception:
        pass
    return out


def line_lens(per_char):
    d = {}
    for _, ln, _ in per_char:
        d[ln] = d.get(ln, 0) + 1
    return [d[k] for k in sorted(d.keys())]


# Test cases — fillers chosen so the natural overflow lands on the char of interest.
# 41 normal CJK chars = 430.5pt (fits in 432). char 42 starts at 520.5, ends 531 (overflow).
F41 = "漢" * 41
F40 = "漢" * 40

PATTERNS = [
    # Single prohibited at pos 42 (single push-down expected)
    ("kuten_42",          F41 + "。漢漢漢"),
    ("touten_42",         F41 + "、漢漢漢"),
    ("close_paren_42",    F41 + "）漢漢漢"),
    ("close_brkt_42",     F41 + "」漢漢漢"),
    # Two prohibited at pos 42, 43 — does Word back up further?
    ("close2_at_42",      F41 + "）」漢漢"),
    ("kuten_close_42",    F41 + "。）漢漢"),
    # Three prohibited at pos 42-44
    ("close3_at_42",      F41 + "）」』漢"),
    # Prohibited at pos 41 (last fittable char), 42 (overflow) — line-end issue
    # If pos 41 is OK and pos 42 is prohibited → 1-char back to pos 40.
    # If pos 40 is also prohibited (line-end-proh like opening paren), what then?
    ("open_at_40",        "漢" * 39 + "（漢）漢漢"),  # 40=（ open(line-end-proh), 42=)
    # The fallback char (pos 41) being itself line-start-prohibited
    ("kuten_at_41_42",    F40 + "。。漢漢"),  # 41=。, 42=。 → can't put either at L2 start
    # Mixed: pos 41 is line-end-proh AND pos 42 is line-start-proh
    ("open41_close42",    F40 + "「）漢漢"),  # 41=「(line-end-proh) 42=)(line-start-proh)
    # The reference pattern from ruby_text_lineheight_11 first 50 chars
    ("ruby_lh11_actual",
     "漢字にルビを付けた文章：「専門用語（せんもんようご）」「技術革新（ぎじゅつかくしん）」「情報処理（じょうほうしょり）」が含まれる段落で、行間の自動拡張動作を検証します。"),
]

print(f"{'pattern':22s}  L1  L2  L3  L4  L5   notes (last 5 chars of L1)")
print("-" * 80)
for name, text in PATTERNS:
    try:
        per_char = measure_via_file(text)
        L = line_lens(per_char)
        L_str = "  ".join(f"{x:2d}" for x in L[:5]) + "  " * (5 - min(len(L), 5))
        # last 5 chars of L1
        l1_len = L[0] if L else 0
        last5 = "".join(c for c, ln, _ in per_char[max(0, l1_len - 5):l1_len])
        next2 = "".join(c for c, ln, _ in per_char[l1_len:l1_len + 2])
        print(f"{name:22s}  {L_str}  L1.tail={last5!r} L2.head={next2!r}")
    except Exception as e:
        print(f"{name:22s}  ERR: {e}")

word.Quit()
