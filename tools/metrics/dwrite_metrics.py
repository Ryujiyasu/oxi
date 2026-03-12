"""Get DirectWrite font metrics via COM interfaces."""
import ctypes
from ctypes import wintypes, POINTER, HRESULT, Structure, c_float, c_uint16, c_int16, c_uint32
import comtypes
from comtypes import GUID

# DirectWrite COM interfaces
class DWRITE_FONT_METRICS(Structure):
    _fields_ = [
        ("designUnitsPerEm", c_uint16),
        ("ascent", c_uint16),
        ("descent", c_uint16),
        ("lineGap", c_int16),
        ("capHeight", c_uint16),
        ("xHeight", c_uint16),
        ("underlinePosition", c_int16),
        ("underlineThickness", c_uint16),
        ("strikethroughPosition", c_int16),
        ("strikethroughThickness", c_uint16),
    ]

class DWRITE_FONT_METRICS1(Structure):
    _fields_ = [
        ("designUnitsPerEm", c_uint16),
        ("ascent", c_uint16),
        ("descent", c_uint16),
        ("lineGap", c_int16),
        ("capHeight", c_uint16),
        ("xHeight", c_uint16),
        ("underlinePosition", c_int16),
        ("underlineThickness", c_uint16),
        ("strikethroughPosition", c_int16),
        ("strikethroughThickness", c_uint16),
        # DWRITE_FONT_METRICS1 extensions
        ("glyphBoxLeft", c_int16),
        ("glyphBoxTop", c_int16),
        ("glyphBoxRight", c_int16),
        ("glyphBoxBottom", c_int16),
        ("subscriptPositionX", c_int16),
        ("subscriptPositionY", c_int16),
        ("subscriptSizeX", c_int16),
        ("subscriptSizeY", c_int16),
        ("superscriptPositionX", c_int16),
        ("superscriptPositionY", c_int16),
        ("superscriptSizeX", c_int16),
        ("superscriptSizeY", c_int16),
        ("hasTypographicMetrics", ctypes.c_bool),
    ]

class DWRITE_GDI_INTEROP_METRICS(Structure):
    """Result from GetGdiCompatibleMetrics"""
    _fields_ = DWRITE_FONT_METRICS._fields_

# Use comtypes to access DirectWrite
try:
    import comtypes.client
    dwrite = ctypes.windll.LoadLibrary("dwrite.dll")
except Exception as e:
    print(f"Failed to load dwrite.dll: {e}")
    exit(1)

# Create DWrite factory
from comtypes import IUnknown, COMMETHOD

IID_IDWriteFactory = GUID("{b859ee5a-d838-4b5b-a2e8-1adc7d93db48}")
CLSID_DWriteFactory = GUID("{b859ee5a-d838-4b5b-a2e8-1adc7d93db48}")

# Simpler approach: use ctypes to call DWriteCreateFactory
DWriteCreateFactory = dwrite.DWriteCreateFactory
DWriteCreateFactory.restype = HRESULT

# Actually, let's use a simpler approach with fontTools + manual calculation
# to verify what DWrite would compute
from fontTools.ttLib import TTFont
import os

fonts = {
    'Yu Gothic': ('C:/Windows/Fonts/YuGothR.ttc', 0),
    'Yu Mincho': ('C:/Windows/Fonts/YuMincho.ttc', 0),
    'Meiryo': ('C:/Windows/Fonts/meiryo.ttc', 0),
    'MS Gothic': ('C:/Windows/Fonts/msgothic.ttc', 0),
    'MS Mincho': ('C:/Windows/Fonts/msgothic.ttc', 2),
    'Calibri': ('C:/Windows/Fonts/calibri.ttf', None),
    'Arial': ('C:/Windows/Fonts/arial.ttf', None),
    'Times New Roman': ('C:/Windows/Fonts/times.ttf', None),
}

def muldiv(a, b, c):
    return (a * b + c // 2) // c

print("Font metric analysis for K derivation")
print("=" * 90)

for name, (path, font_num) in fonts.items():
    if not os.path.exists(path):
        print(f"{name}: NOT FOUND")
        continue

    try:
        if font_num is not None:
            tt = TTFont(path, fontNumber=font_num)
        else:
            tt = TTFont(path)
    except Exception as e:
        print(f"{name}: ERROR {e}")
        continue

    os2 = tt.get('OS/2')
    hhea = tt.get('hhea')
    head = tt.get('head')
    vhea = tt.get('vhea')

    UPM = head.unitsPerEm
    wA = os2.usWinAscent
    wD = os2.usWinDescent
    hA = abs(hhea.ascent)
    hD = abs(hhea.descent)
    hG = hhea.lineGap
    tA = os2.sTypoAscender
    tD = abs(os2.sTypoDescender)
    tG = os2.sTypoLineGap
    useTypo = bool(os2.fsSelection & 0x80)

    # DWrite logic:
    # If hasTypographicMetrics (USE_TYPO_METRICS, fsSelection bit 7):
    #   ascent = typoAscender, descent = |typoDescender|, lineGap = typoLineGap
    # Else:
    #   ascent = winAscent, descent = winDescent, lineGap = 0
    #   (DWrite treats win metrics as having no line gap)
    #
    # But DWrite also has GetGdiCompatibleMetrics which adds GDI-compatible rounding

    if useTypo:
        dw_asc = tA
        dw_desc = tD
        dw_gap = tG
    else:
        dw_asc = wA
        dw_desc = wD
        dw_gap = 0  # DWrite uses 0 for gap when using win metrics

    dw_sum = dw_asc + dw_desc + dw_gap

    # Computed line height: (asc + desc + gap) / UPM * fontSize
    # For 10.5pt: dw_sum / UPM * 10.5 (in pt) = dw_sum / UPM * 210 (in twips)
    lh_twips_10_5 = muldiv(dw_sum, 210, UPM)

    # Also compute with hhea metrics
    hhea_sum = hA + hD + hG
    lh_hhea_10_5 = muldiv(hhea_sum, 210, UPM)

    # Also compute with typo metrics
    typo_sum = tA + tD + tG
    lh_typo_10_5 = muldiv(typo_sum, 210, UPM)

    # What about max(hhea_sum, win_sum(+lineGap=0), typo_sum)?
    win_sum = wA + wD
    max_sum = max(hhea_sum, win_sum, typo_sum)
    lh_max_10_5 = muldiv(max_sum, 210, UPM)

    # For GDI-compatible: DWrite adds the external leading separately
    # GetGdiCompatibleMetrics adjusts metrics to match GDI at a given ppem

    vhea_sum = (abs(vhea.ascent) + abs(vhea.descent) + vhea.lineGap) if vhea else 0

    print(f"{name}:")
    print(f"  UPM={UPM} useTypo={useTypo}")
    print(f"  win: {wA}+{wD}={win_sum}")
    print(f"  hhea: {hA}+{hD}+{hG}={hhea_sum}")
    print(f"  typo: {tA}+{tD}+{tG}={typo_sum}")
    if vhea:
        print(f"  vhea: {abs(vhea.ascent)}+{abs(vhea.descent)}+{vhea.lineGap}={vhea_sum}")
    print(f"  DWrite(10.5pt): sum={dw_sum} -> {lh_twips_10_5}tw={lh_twips_10_5/20:.2f}pt")
    print(f"  hhea(10.5pt): sum={hhea_sum} -> {lh_hhea_10_5}tw={lh_hhea_10_5/20:.2f}pt")
    print(f"  typo(10.5pt): sum={typo_sum} -> {lh_typo_10_5}tw={lh_typo_10_5/20:.2f}pt")
    print(f"  max(10.5pt): sum={max_sum} -> {lh_max_10_5}tw={lh_max_10_5/20:.2f}pt")

    # Check if hhea-based DWrite (ascent=hA, descent=hD, gap=hG) is used
    # Some DWrite implementations fall back to hhea when USE_TYPO_METRICS is false
    # and the font doesn't have "good" win metrics
    print(f"  Possible K candidates:")
    for label, K_val in [
        ("win_sum", win_sum),
        ("hhea_sum", hhea_sum),
        ("typo_sum", typo_sum),
        (f"win_sum * 83/64", win_sum * 83 // 64),
        (f"hhea_sum * 83/64", hhea_sum * 83 // 64),
        (f"typo_sum * 83/64", typo_sum * 83 // 64),
        (f"max + typoG", max(win_sum, hhea_sum) + tG),
        (f"win + typoG", win_sum + tG),
        (f"win + hheaG", win_sum + hG),
        (f"hhea + typoG-hheaG", hhea_sum + tG - hG),
        (f"typo + (winA-tA)", typo_sum + (wA - tA)),
        (f"typo + (winD-tD)", typo_sum + (wD - tD)),
        (f"typo + (wA-tA)+(wD-tD)", typo_sum + (wA - tA) + (wD - tD)),
    ]:
        lh = muldiv(K_val, 210, UPM)
        print(f"    {label:<25} K={K_val:>5} -> {lh}tw = {lh/20:.2f}pt")

    print()
    tt.close()
