"""§4.7c Mech 3 — alignment gate full coverage test.

Round 1 confirmed:
  cSC=compressPunctuation + jc=left  → fires
  cSC=compressPunctuation + jc=both  → fires
  cSC=doNotCompress      + jc=*     → does not fire

Untested:
  cSC=compressPunctuation + jc=center
  cSC=compressPunctuation + jc=right
  cSC=compressPunctuation + jc=distribute (= wdAlignParagraphDistribute, val="distribute")

This script clones 7f272a and tests all 5 alignments × cSC ∈ {compress, doNotCompress}.
"""
import json, os, sys, re, time, zipfile, shutil, tempfile
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

SRC_REAL = os.path.abspath(
    r"C:\Users\ryuji\oxi-1\tools\golden-test\documents\docx\7f272a2dfd3b_index-21.docx")
OUT_DIR = os.path.abspath(r"C:\Users\ryuji\oxi-1\pipeline_data\mech3_align_full_docs")
RESULT = os.path.abspath(r"C:\Users\ryuji\oxi-1\pipeline_data\mech3_align_full.json")
os.makedirs(OUT_DIR, exist_ok=True)

YAKUMONO = set("（「『【〔｛〈《［）」』】〕｝〉》］、。，．—")

# 7f272a P13 actual text
PROBE_TEXT = (
    "卸売市場法第６条第１項（第14条において準用する同法第６条第１項）"
    "の規定により、中央卸売市場（地方卸売市場）に係る認定事項の変更について"
    "認定を受けたいので、次のとおり関係書類を添えて申請します。"
)


def make_doc_xml(text, jc):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
            ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<w:body><w:p>'
            f'<w:pPr><w:jc w:val="{jc}"/>'
            '<w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr></w:pPr>'
            '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>'
            f'<w:t>{text}</w:t></w:r></w:p>'
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838" w:code="9"/>'
            '<w:pgMar w:top="1134" w:right="1304" w:bottom="1134" w:left="1304"'
            ' w:header="851" w:footer="992" w:gutter="0"/>'
            '<w:cols w:space="720"/>'
            '<w:docGrid w:type="lines" w:linePitch="360"/>'
            '</w:sectPr></w:body></w:document>')


def replace_csc(s, value):
    if value is None:
        return re.sub(r'<w:characterSpacingControl[^/]*/>', '', s)
    return re.sub(
        r'<w:characterSpacingControl[^/]*?w:val="[^"]*"/>',
        f'<w:characterSpacingControl w:val="{value}"/>', s
    )


def make_variant(label, csc_value, jc):
    out_path = os.path.join(OUT_DIR, f"{label}.docx")
    tmp = tempfile.mkdtemp(prefix="m3a_")
    try:
        with zipfile.ZipFile(SRC_REAL) as z:
            z.extractall(tmp)
        with open(os.path.join(tmp, "word", "document.xml"), "w", encoding="utf-8") as f:
            f.write(make_doc_xml(PROBE_TEXT, jc))
        sp = os.path.join(tmp, "word", "settings.xml")
        with open(sp, "r", encoding="utf-8") as f:
            s = f.read()
        s = replace_csc(s, csc_value)
        with open(sp, "w", encoding="utf-8") as f:
            f.write(s)
        with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
            for root, _, files in os.walk(tmp):
                for fn in files:
                    full = os.path.join(root, fn)
                    arc = os.path.relpath(full, tmp).replace("\\", "/")
                    z.write(full, arc)
        return out_path
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def measure(word, path):
    d = word.Documents.Open(path, ReadOnly=True)
    time.sleep(0.2)
    try:
        chars = d.Range().Characters
        xs = []
        for ci in range(1, chars.Count + 1):
            try:
                c = chars(ci)
                t = c.Text
                if t in ("\r", "\x07"):
                    continue
                xs.append((t, float(c.Information(5)), float(c.Information(6)),
                           float(c.Font.Size if c.Font.Size else 0)))
            except Exception:
                continue
        # Word's reported alignment
        try:
            actual_align = d.Paragraphs(1).Alignment
        except Exception:
            actual_align = None
    finally:
        try: d.Close(SaveChanges=False)
        except: pass
    if not xs: return {"error": "no chars"}
    lines_b = {}
    for t, x, y, sz in xs:
        ykey = round(y, 0)
        lines_b.setdefault(ykey, []).append((t, x, y, sz))
    n_yak_comp = 0
    n_yak_total = 0
    detail = []
    for ykey in sorted(lines_b.keys()):
        items = sorted(lines_b[ykey], key=lambda v: v[1])
        for i in range(len(items) - 1):
            t = items[i][0]
            a = round(items[i+1][1] - items[i][1], 3)
            sz = items[i][3]
            if t in YAKUMONO:
                n_yak_total += 1
                if sz > 0 and a < sz * 0.99:
                    n_yak_comp += 1
                    if len(detail) < 8:
                        detail.append((t, round(a, 2), round(sz, 1)))
    return {
        "actual_alignment_int": actual_align,
        "n_yak_total": n_yak_total,
        "n_yak_compressed": n_yak_comp,
        "fires": n_yak_comp > 0,
        "comp_detail_first8": detail,
    }


VARIANTS = []
for csc in ["compressPunctuation", "doNotCompress"]:
    csc_short = "comp" if csc == "compressPunctuation" else "noComp"
    for jc in ["left", "both", "center", "right", "distribute"]:
        VARIANTS.append((f"AL_{csc_short}_{jc}", csc, jc))


def kill_word():
    import subprocess
    try:
        subprocess.run(['taskkill','/F','/IM','WINWORD.EXE'], capture_output=True)
    except Exception: pass
    time.sleep(3)


def measure_with_restart(label, p):
    """Try measurement, restarting Word on RPC failure."""
    for attempt in range(3):
        try:
            word = w32.Dispatch("Word.Application")
            word.Visible = False
            word.DisplayAlerts = False
        except Exception as e:
            kill_word()
            continue
        try:
            r = measure(word, p)
            try: word.Quit()
            except: pass
            return r
        except Exception as e:
            try: word.Quit()
            except: pass
            if "RPC" in str(e) or "サーバー" in str(e) or "呼び出し" in str(e):
                kill_word()
                continue
            return {"measure_error": str(e)}
    return {"measure_error": "max retries"}


def main():
    out = {}
    for label, csc, jc in VARIANTS:
        try:
            p = make_variant(label, csc, jc)
        except Exception as e:
            out[label] = {"build_error": str(e)}
            print(f"[{label}] BUILD ERR: {e}")
            continue
        # Always restart Word per variant for reliability
        kill_word()
        r = measure_with_restart(label, p)
        out[label] = {"csc": csc, "jc": jc, **r}
        fire = "FIRE" if r.get("fires") else "no  "
        detail = r.get("comp_detail_first8", [])
        detail_str = ", ".join(f"{t}={a}/{s}" for t, a, s in detail[:4])
        print(f"[{label:<28s}] csc={csc:<22s} jc={jc:<11s} word_align={r.get('actual_alignment_int','?')}  {fire}  comp={r.get('n_yak_compressed','?')}/{r.get('n_yak_total','?')}  ex={detail_str}")

    os.makedirs(os.path.dirname(RESULT), exist_ok=True)
    with open(RESULT, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)

    print("\n=== Alignment × cSC matrix ===")
    print(f"{'jc':>12} | {'compressPunctuation':>22} | {'doNotCompress':>22}")
    print("-" * 64)
    for jc in ["left", "both", "center", "right", "distribute"]:
        comp_label = f"AL_comp_{jc}"
        no_label = f"AL_noComp_{jc}"
        c_fire = "FIRES" if out.get(comp_label, {}).get("fires") else "  no "
        n_fire = "FIRES" if out.get(no_label, {}).get("fires") else "  no "
        c_n = out.get(comp_label, {}).get("n_yak_compressed", "?")
        n_n = out.get(no_label, {}).get("n_yak_compressed", "?")
        print(f"{jc:>12} | {c_fire:>4} (n={c_n:>2})         | {n_fire:>4} (n={n_n:>2})")


if __name__ == "__main__":
    main()
