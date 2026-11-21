# -*- coding: utf-8 -*-
"""When the glyph after a fitting one is line-start prohibited (々), does Word pull it in (oikomi, squeezing the
line's marks) or push the pair down? Sweep the overflow o of the prohibited glyph past the margin for k = 0..3
marks (、) on the line; control: a plain glyph (本) in the same place. ＭＳ 明朝 10.5, compat 15,
compressPunctuation, jc both, docGrid lines, balance (the blind-A ja/creative__14af1ca6 settings)."""
import os, sys, zipfile
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _pb_tblstyle_layer_gen import RELS, pdf, oxi
from _pb_vert_band_gen import W, CT, DR
ST = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + W + '><w:characterSpacingControl w:val="compressPunctuation"/><w:compat><w:balanceSingleByteDoubleByteWidth/><w:useFELayout/><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
STY = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + W + '><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:kern w:val="2"/><w:sz w:val="21"/><w:szCs w:val="22"/><w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
       '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr></w:style></w:styles>')
CT2 = CT.replace('</Types>', '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>')
DR2 = DR.replace('</Relationships>', '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>')
def line(k):
    base = list("国家公務員給与法案参院本会議可決成立物価高騰賃金実質春闘前進生活")[:29 - k]
    for i in range(k):
        base.insert(6 + i * 7, "、")
    return "".join(base)
def gen(out):
    os.makedirs(out, exist_ok=True)
    import os as _o
    ext = bool(_o.environ.get("EXT"))
    for k in (range(1, 5) if ext else range(4)):
        for ctl, pair in ((("ki", "日々"),) if ext else (("ki", "日々"), ("pl", "日本"))):
            for o10 in (range(100, 230, 5) if ext else range(0, 110, 10)):  # overflow in 0.1pt
                o = o10 / 10.0
                avail = 31 * 10.5 - o
                right = round((595.3 - 56.7 - avail) * 20)
                text = line(k) + pair + "の後に続く文章です"
                doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + W + '><w:body><w:p><w:r><w:t>%s</w:t></w:r></w:p>' % text +
                       '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="%d" w:bottom="1134" w:left="1134" w:header="851" w:footer="992" w:gutter="0"/>'
                       '<w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr></w:body></w:document>' % right)
                with zipfile.ZipFile(os.path.join(out, "k%d_%s_o%03d.docx" % (k, ctl, o10)), "w", zipfile.ZIP_DEFLATED) as z:
                    z.writestr("[Content_Types].xml", CT2); z.writestr("_rels/.rels", RELS); z.writestr("word/document.xml", doc)
                    z.writestr("word/_rels/document.xml.rels", DR2); z.writestr("word/settings.xml", ST); z.writestr("word/styles.xml", STY)
def first_line(p):
    import fitz, warnings; warnings.filterwarnings("ignore")
    rows = {}
    for b in fitz.open(p)[0].get_text("rawdict")["blocks"]:
        for l in b.get("lines", []):
            for s in l["spans"]:
                for c in s["chars"]:
                    rows.setdefault(round(c["origin"][1]), []).append(c["c"])
    return "".join(rows[min(rows)]) if rows else ""
def table(out):
    import collections
    res = collections.defaultdict(dict)
    for f in sorted(os.listdir(out)):
        if f.endswith(".docx"):
            b = os.path.join(out, f[:-5]); k, ctl, o = f[:-5].split("_")
            w = first_line(b + ".pdf"); ox = first_line(b + "_oxi.pdf") if os.path.exists(b + "_oxi.pdf") else ""
            last = lambda t: "2" if t.rstrip().endswith(("々", "本")) else ("1" if t.rstrip().endswith("日") else "0")
            res[(k, ctl)][int(o[1:])] = last(w) + ("/" + last(ox) if ox else "")
    for key in sorted(res):
        print(key, " ".join("%4.1f:%s" % (o / 10, v) for o, v in sorted(res[key].items())))
if __name__ == "__main__":
    {"gen": lambda: gen(sys.argv[2]), "pdf": lambda: pdf(sys.argv[2]), "oxi": lambda: oxi(sys.argv[2], sys.argv[3]), "table": lambda: table(sys.argv[2])}[sys.argv[1]]()
