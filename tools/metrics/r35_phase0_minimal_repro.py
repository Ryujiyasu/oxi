"""R35 Phase 0: minimal repro of Mech 2 wrap-budget mechanism.

Goal: build a synthetic OOXML with kern + compressPunctuation + jc=both
that shows Word fitting N+1 chars/line by Mech 2 compression on overflow.
Verify by comparing Word's per-char measurement vs natural width.

Test design:
  Page width: 11906 twips (A4 wide), margin L=R=1700 → content_w = 8506tw = 425.3pt
  For 12pt MS Mincho yak:
    Each yak natural = 12pt
    With kern + compressPunctuation, yak can compress ~6pt

  Test text: 漢漢漢漢漢「漢漢漢漢漢」漢漢漢漢漢、漢漢漢 (21 chars, 3 yak)
    21 × 12pt = 252pt natural
    Each yak max compress ~6pt (50% rule per Phase 1 max ratio 0.5)
    3 yak × 6pt = 18pt max compress
    So natural 252pt fits in any cw ≥ (252 - 18) = 234pt with full compress

  Variants:
    V1: cw = 252pt (natural exactly fits) → Word fits 21 chars no compression
    V2: cw = 250pt (natural overflows 2pt) → Word fits 21 chars, 2pt distributed
    V3: cw = 240pt (natural overflows 12pt) → Word fits 21 chars, 12pt distributed
    V4: cw = 234pt (natural overflows 18pt = max savings) → Word fits 21 chars, max compress
    V5: cw = 233pt (overflow exceeds max savings) → Word wraps to 20 chars
    V6: cw = 254pt (under-fill) → Word fits 21 chars natural

  Also try with `LongerSet` (60 chars / 8 yak) on cw = 720pt baseline

Expected outcome (per Phase 1 Word measurement model):
  V1-V4: 21 chars/line, comp_total = max(0, 252 - cw)
  V5: 20 chars/line (wrap), no compression of standalone yak

This will be the ground truth Oxi must reproduce.
Comparison to Oxi (R34 strict) on same input is Phase 5.
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

OUT_DIR = os.path.abspath("pipeline_data/r35_phase0_docs")
RESULT_PATH = os.path.abspath("pipeline_data/r35_phase0_2026-05-02.json")

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml"
   ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml"
   ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml"
   ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>
"""

RELS_ROOT = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""

WORD_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>
"""

STYLES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults><w:rPrDefault><w:rPr>
    <w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/>
    <w:kern w:val="2"/>
    <w:sz w:val="24"/>
    <w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/>
  </w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="a">
    <w:name w:val="Normal"/><w:qFormat/>
  </w:style>
</w:styles>
"""

SETTINGS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:characterSpacingControl w:val="compressPunctuation"/>
  <w:compat>
    <w:compatSetting w:name="compatibilityMode"
     w:uri="http://schemas.microsoft.com/office/word" w:val="14"/>
  </w:compat>
</w:settings>
"""


def gen_doc(text, content_w_tw, jc="both"):
    # Page width 11906 (A4w), margins L=R = (11906 - content_w_tw)/2 each
    pgw = 11906
    margin_lr = (pgw - content_w_tw) // 2
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:body><w:p>'
            f'<w:pPr><w:jc w:val="{jc}"/></w:pPr>'
            '<w:r>'
            '<w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>'
            f'<w:t>{text}</w:t>'
            '</w:r></w:p>'
            '<w:sectPr>'
            f'<w:pgSz w:w="{pgw}" w:h="16838"/>'
            f'<w:pgMar w:top="1440" w:right="{margin_lr}" w:bottom="1440" '
            f'w:left="{margin_lr}" w:header="720" w:footer="720" w:gutter="0"/>'
            '<w:cols w:space="720"/>'
            '<w:docGrid w:type="default" w:linePitch="360"/>'
            '</w:sectPr></w:body></w:document>')


def write_docx(path, text, content_w_tw, jc="both"):
    tmp = tempfile.mkdtemp(prefix="r35p0_")
    try:
        os.makedirs(os.path.join(tmp, "_rels"), exist_ok=True)
        os.makedirs(os.path.join(tmp, "word", "_rels"), exist_ok=True)
        with open(os.path.join(tmp, "[Content_Types].xml"), "w",
                  encoding="utf-8") as f:
            f.write(CONTENT_TYPES)
        with open(os.path.join(tmp, "_rels", ".rels"), "w",
                  encoding="utf-8") as f:
            f.write(RELS_ROOT)
        with open(os.path.join(tmp, "word", "_rels", "document.xml.rels"),
                  "w", encoding="utf-8") as f:
            f.write(WORD_RELS)
        with open(os.path.join(tmp, "word", "styles.xml"), "w",
                  encoding="utf-8") as f:
            f.write(STYLES_XML)
        with open(os.path.join(tmp, "word", "settings.xml"), "w",
                  encoding="utf-8") as f:
            f.write(SETTINGS_XML)
        with open(os.path.join(tmp, "word", "document.xml"), "w",
                  encoding="utf-8") as f:
            f.write(gen_doc(text, content_w_tw, jc))
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
            for root, _, files in os.walk(tmp):
                for fn in files:
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


def measure(word, path, expected_text):
    d = word.Documents.Open(path, ReadOnly=True)
    time.sleep(0.5)
    chars = d.Range().Characters
    rows = []
    for ci in range(1, chars.Count + 1):
        try:
            c = chars(ci)
            t = c.Text
            if t in ("\r", "\x07"):
                continue
            rows.append({
                "ch": t, "x": float(c.Information(5)),
                "y": float(c.Information(6)),
                "page": int(c.Information(3)),
                "size": float(c.Font.Size),
            })
        except Exception:
            continue
    d.Close(SaveChanges=False)
    # Compute advance + group by line
    lines = []
    cur_line = []
    last_y = None
    for r in rows:
        if last_y is not None and abs(r["y"] - last_y) > 1.0:
            lines.append(cur_line)
            cur_line = []
        cur_line.append(r)
        last_y = r["y"]
    if cur_line:
        lines.append(cur_line)
    line_summaries = []
    for li, line in enumerate(lines):
        n = len(line)
        x_start = line[0]["x"]
        x_end = line[-1]["x"] + line[-1]["size"]  # estimated
        cw_obs = x_end - x_start
        # advances
        advs = []
        nat_sum = 0
        obs_sum = 0
        n_yak = 0
        n_comp = 0
        for i, r in enumerate(line):
            sz = r["size"]
            adv = (line[i + 1]["x"] - r["x"]) if i + 1 < len(line) else sz
            ratio = adv / sz if sz else None
            cl = cls(r["ch"])
            advs.append({"ch": r["ch"], "adv": round(adv, 4), "sz": sz,
                          "ratio": round(ratio, 4) if ratio else None,
                          "cls": cl})
            obs_sum += adv
            if cl in ("A", "B"):
                nat_sum += sz
                n_yak += 1
                if ratio < 0.85:
                    n_comp += 1
            else:
                nat_sum += adv
        line_summaries.append({
            "line": li + 1, "n": n, "x_start": x_start,
            "obs_sum": round(obs_sum, 2), "nat_sum": round(nat_sum, 2),
            "n_yak": n_yak, "n_compressed": n_comp,
            "advances": advs,
        })
    return line_summaries


# 21 chars, 3 yak (B), 12pt → 252pt natural
TEXT_21_3 = "漢漢漢漢漢「漢漢漢漢漢」漢漢漢漢漢、漢漢漢"

# Margin L = R = (pgW - content_w_tw)/2; content_w_tw = content_w_pt * 20
def cw_pt_to_tw(cw_pt):
    return int(round(cw_pt * 20))


VARIANTS = [
    ("V1_252pt_jcboth", TEXT_21_3, cw_pt_to_tw(252.0), "both"),
    ("V2_250pt_jcboth", TEXT_21_3, cw_pt_to_tw(250.0), "both"),
    ("V3_245pt_jcboth", TEXT_21_3, cw_pt_to_tw(245.0), "both"),
    ("V4_240pt_jcboth", TEXT_21_3, cw_pt_to_tw(240.0), "both"),
    ("V5_236pt_jcboth", TEXT_21_3, cw_pt_to_tw(236.0), "both"),
    ("V6_254pt_jcboth", TEXT_21_3, cw_pt_to_tw(254.0), "both"),
    ("V7_252pt_jcleft", TEXT_21_3, cw_pt_to_tw(252.0), "left"),
    ("V8_245pt_jcleft", TEXT_21_3, cw_pt_to_tw(245.0), "left"),
    ("V9_240pt_jcleft", TEXT_21_3, cw_pt_to_tw(240.0), "left"),
]


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    results = {}
    for label, text, cw_tw, jc in VARIANTS:
        path = os.path.join(OUT_DIR, f"{label}.docx")
        write_docx(path, text, cw_tw, jc)
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False
        try:
            try:
                line_summaries = measure(word, path, text)
            except Exception as e:
                results[label] = {"error": str(e)}
                print(f"[{label}] ERR: {e}", flush=True)
                continue
            cw_pt = cw_tw / 20
            results[label] = {
                "cw_tw": cw_tw, "cw_pt": cw_pt, "jc": jc,
                "n_lines": len(line_summaries),
                "lines": line_summaries,
            }
            print(f"\n[{label}] cw={cw_pt:.1f}pt jc={jc} → {len(line_summaries)} lines:",
                  flush=True)
            for ls in line_summaries:
                print(f"   line {ls['line']}: n={ls['n']:>2} "
                      f"obs={ls['obs_sum']:>6.1f} nat={ls['nat_sum']:>6.1f} "
                      f"slack_obs={cw_pt - ls['obs_sum']:>+6.2f} "
                      f"slack_nat={cw_pt - ls['nat_sum']:>+6.2f} "
                      f"yak={ls['n_yak']} comp={ls['n_compressed']}",
                      flush=True)
                for a in ls["advances"]:
                    if a["cls"] in ("A", "B"):
                        flag = " *" if a.get("ratio", 1.0) < 0.85 else "  "
                        print(f"     {flag} {a['ch']} adv={a['adv']:.2f} "
                              f"r={a['ratio']:.4f}", flush=True)
        finally:
            try:
                word.Quit()
            except Exception:
                pass
            time.sleep(1.0)

    with open(RESULT_PATH, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f"\nWrote {RESULT_PATH}", flush=True)


if __name__ == "__main__":
    main()
