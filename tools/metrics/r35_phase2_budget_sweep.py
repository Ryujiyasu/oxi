"""R35 Phase 2: focused Mech 2 budget characterization.

Phase 0 found:
- 3 standalone B yak: 2pt overflow accepted (V2), 7pt overflow refused (V3)
- Per-yak budget cap somewhere between 0.67 and 2.33pt

Phase 2 narrows the boundary + checks variants.

Probes (20 cells):
  P1 (3 standalone B yak, jc=both): cw sweep 252→242 in 1pt steps (11 cw points)
       to find exact compression budget cap
  P2 (3 FINAL-RULE B yak: 漢「」、漢): cw sweep 252→230 (testing Mech 1+2 stack)
       Use text where 」、 forms B→B FINAL RULE pair
  P3 (3 standalone A yak): cw sweep, A-class budget
  P4 (jc=left, same text): verify no Mech 2 fire even with negative slack
"""
import os
import time
import json
import zipfile
import shutil
import sys
import tempfile
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath("pipeline_data/r35_phase2_docs")
RESULT_PATH = os.path.abspath("pipeline_data/r35_phase2_2026-05-02.json")

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>
"""
RELS_ROOT = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""
WORD_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>"""

STYLES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr>
<w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/>
<w:kern w:val="2"/><w:sz w:val="24"/>
<w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/>
</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:qFormat/></w:style>
</w:styles>"""

SETTINGS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:characterSpacingControl w:val="compressPunctuation"/>
<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/></w:compat>
</w:settings>"""


def gen_doc(text, content_w_tw, jc):
    pgw = 11906
    margin_lr = (pgw - content_w_tw) // 2
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:body><w:p>'
            f'<w:pPr><w:jc w:val="{jc}"/></w:pPr>'
            '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>'
            f'<w:t>{text}</w:t></w:r></w:p>'
            f'<w:sectPr><w:pgSz w:w="{pgw}" w:h="16838"/>'
            f'<w:pgMar w:top="1440" w:right="{margin_lr}" w:bottom="1440" '
            f'w:left="{margin_lr}" w:header="720" w:footer="720" w:gutter="0"/>'
            '<w:cols w:space="720"/><w:docGrid w:type="default" w:linePitch="360"/>'
            '</w:sectPr></w:body></w:document>')


def write_docx(path, text, content_w_tw, jc):
    tmp = tempfile.mkdtemp(prefix="r35p2_")
    try:
        os.makedirs(os.path.join(tmp, "_rels"), exist_ok=True)
        os.makedirs(os.path.join(tmp, "word", "_rels"), exist_ok=True)
        files = [
            ("[Content_Types].xml", CONTENT_TYPES),
            ("_rels/.rels", RELS_ROOT),
            ("word/_rels/document.xml.rels", WORD_RELS),
            ("word/styles.xml", STYLES_XML),
            ("word/settings.xml", SETTINGS_XML),
            ("word/document.xml", gen_doc(text, content_w_tw, jc)),
        ]
        for relpath, content in files:
            full = os.path.join(tmp, relpath.replace("/", os.sep))
            os.makedirs(os.path.dirname(full), exist_ok=True)
            with open(full, "w", encoding="utf-8") as f:
                f.write(content)
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
            for root, _, names in os.walk(tmp):
                for fn in names:
                    full = os.path.join(root, fn)
                    arc = os.path.relpath(full, tmp).replace("\\", "/")
                    z.write(full, arc)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


YAKUMONO_A = set("（「『【〔｛〈《［")
YAKUMONO_B = set("）」』】〕｝〉》］、。，．—")


def cls(ch):
    if ch in YAKUMONO_A: return "A"
    if ch in YAKUMONO_B: return "B"
    return "X"


def measure(word, path):
    d = word.Documents.Open(path, ReadOnly=True)
    time.sleep(0.4)
    chars = d.Range().Characters
    rows = []
    for ci in range(1, chars.Count + 1):
        try:
            c = chars(ci)
            t = c.Text
            if t in ("\r", "\x07"): continue
            rows.append({"ch": t, "x": float(c.Information(5)),
                          "y": float(c.Information(6)),
                          "size": float(c.Font.Size)})
        except Exception:
            continue
    d.Close(SaveChanges=False)
    lines, cur, last_y = [], [], None
    for r in rows:
        if last_y is not None and abs(r["y"] - last_y) > 1.0:
            lines.append(cur); cur = []
        cur.append(r); last_y = r["y"]
    if cur: lines.append(cur)
    out_lines = []
    for li, line in enumerate(lines):
        n = len(line)
        advs = []
        nat_sum = obs_sum = 0
        n_yak = n_comp = 0
        for i, r in enumerate(line):
            sz = r["size"]
            adv = (line[i+1]["x"] - r["x"]) if i+1 < len(line) else sz
            ratio = adv / sz if sz else None
            cl = cls(r["ch"])
            advs.append({"ch": r["ch"], "adv": round(adv, 4),
                          "r": round(ratio, 4) if ratio else None, "cls": cl})
            obs_sum += adv
            if cl in ("A", "B"):
                nat_sum += sz; n_yak += 1
                if ratio is not None and ratio < 0.85: n_comp += 1
            else:
                nat_sum += adv
        out_lines.append({"line": li+1, "n": n,
                           "obs": round(obs_sum, 2),
                           "nat": round(nat_sum, 2),
                           "n_yak": n_yak, "n_comp": n_comp,
                           "advs": advs})
    return out_lines


# Texts (each 12pt, 21 chars total, but yak position differs)
# P1: 3 standalone B yak (no FINAL RULE pairs)
T_STANDALONE_B = "漢漢漢漢漢「漢漢漢漢漢」漢漢漢漢漢、漢漢漢"
# P2: 3 yak forming B→B FINAL RULE pair (」、) — total 21 chars
T_FINAL_RULE_B = "漢漢漢漢漢漢漢漢漢「漢漢漢漢漢漢」、漢漢漢"
# P3: 3 standalone A yak
T_STANDALONE_A = "漢漢漢漢漢（漢漢漢漢漢「漢漢漢漢漢『漢漢漢"
# P4 same as P1, jc=left


def cw(pt):
    return int(round(pt * 20))


VARIANTS = []
# P1: 3 standalone B yak — narrow scan around budget cap
for w in [252, 251, 250, 249, 248, 247, 246]:
    VARIANTS.append((f"P1_B_{w}_jcboth", T_STANDALONE_B, cw(w), "both"))

# P2: FINAL RULE B→B yak — much wider compression expected
for w in [252, 248, 244, 240, 235, 230]:
    VARIANTS.append((f"P2_FB_{w}_jcboth", T_FINAL_RULE_B, cw(w), "both"))

# P3: 3 standalone A yak
for w in [252, 250, 248, 246]:
    VARIANTS.append((f"P3_A_{w}_jcboth", T_STANDALONE_A, cw(w), "both"))

# P4: jc=left no compression check
for w in [252, 250, 248]:
    VARIANTS.append((f"P4_B_{w}_jcleft", T_STANDALONE_B, cw(w), "left"))


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    results = {}
    for label, text, cw_tw, jc in VARIANTS:
        path = os.path.join(OUT_DIR, f"{label}.docx")
        write_docx(path, text, cw_tw, jc)
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False; word.DisplayAlerts = False
        try:
            try:
                lines = measure(word, path)
            except Exception as e:
                results[label] = {"error": str(e)}
                print(f"[{label}] ERR: {e}", flush=True)
                continue
            cw_pt = cw_tw / 20
            results[label] = {"cw_pt": cw_pt, "jc": jc, "n_lines": len(lines),
                               "lines": lines}
            for ls in lines:
                comp_total = round(ls["nat"] - ls["obs"], 2)
                slack_nat = round(cw_pt - ls["nat"], 2)
                slack_obs = round(cw_pt - ls["obs"], 2)
                marker = "*" if ls["n_comp"] > 0 else " "
                print(f"{marker}[{label}] L{ls['line']} n={ls['n']:>2} "
                      f"obs={ls['obs']:>6.2f} nat={ls['nat']:>6.2f} "
                      f"slack_nat={slack_nat:>+6.2f} comp={comp_total:>4.1f} "
                      f"yak={ls['n_yak']} comp_yak={ls['n_comp']}",
                      flush=True)
                if ls["n_comp"] > 0:
                    for a in ls["advs"]:
                        if a["cls"] in ("A","B") and (a["r"] or 1) < 0.99:
                            print(f"      {a['ch']} adv={a['adv']:.2f} "
                                  f"r={a['r']:.4f}", flush=True)
        finally:
            try: word.Quit()
            except: pass
            time.sleep(1.0)

    with open(RESULT_PATH, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f"\nWrote {RESULT_PATH}", flush=True)


if __name__ == "__main__":
    main()
