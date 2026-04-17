"""Further bisection: replace parts of S1 to find trigger.

S1 had: only para 28 body, but kept styles.xml / theme.xml / sectPr intact.
Now replace each of those one at a time.
"""
import os, re, zipfile

SRC = os.path.abspath(r"pipeline_data/d77a_S1_only_para28.docx")
OUT_DIR = os.path.abspath(r"pipeline_data")

MINIMAL_STYLES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>
'''

MINIMAL_THEME = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office">
<a:themeElements>
<a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme>
<a:fontScheme name="Office"><a:majorFont><a:latin typeface="Cambria"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>
<a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>
</a:themeElements>
</a:theme>
'''


def rewrite(src, dst, transforms):
    with zipfile.ZipFile(src, 'r') as zin:
        with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.namelist():
                data = zin.read(item)
                if item in transforms:
                    data = transforms[item](data)
                zout.writestr(item, data)


def main():
    # S6: S1 + minimal styles.xml
    S6 = os.path.join(OUT_DIR, "d77a_S6_minimal_styles.docx")
    rewrite(SRC, S6, {"word/styles.xml": lambda _: MINIMAL_STYLES.encode("utf-8")})
    print(f"[S6] {S6}")

    # S7: S1 + minimal theme.xml
    S7 = os.path.join(OUT_DIR, "d77a_S7_minimal_theme.docx")
    rewrite(SRC, S7, {"word/theme/theme1.xml": lambda _: MINIMAL_THEME.encode("utf-8")})
    print(f"[S7] {S7}")

    # S8: S1 + strip docGrid from sectPr
    def strip_docgrid(data):
        xml = data.decode("utf-8")
        xml = re.sub(r'<w:docGrid[^/]*/>', '', xml)
        return xml.encode("utf-8")
    S8 = os.path.join(OUT_DIR, "d77a_S8_no_docgrid.docx")
    rewrite(SRC, S8, {"word/document.xml": strip_docgrid})
    print(f"[S8] {S8}")

    # S9: S1 + remove w:hint="eastAsia" from rFonts
    def strip_hint(data):
        xml = data.decode("utf-8")
        xml = re.sub(r' w:hint="eastAsia"', '', xml)
        return xml.encode("utf-8")
    S9 = os.path.join(OUT_DIR, "d77a_S9_no_hint.docx")
    rewrite(SRC, S9, {"word/document.xml": strip_hint})
    print(f"[S9] {S9}")


if __name__ == "__main__":
    main()
