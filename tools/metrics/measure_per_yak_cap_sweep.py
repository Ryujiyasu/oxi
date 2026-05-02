"""§4.7b — per-yak cap formula regression for N ∈ {2, 3, 4, 5, 7}.

Known:
  N=1: cap = 4.0pt = fontSize/3 (at 12pt font)
  N=9: cap = 2.5pt = fontSize×5/24 (at 12pt font)

Hypotheses:
  (i) Stepped: N=1 → fontSize/3, N≥2 → fontSize×5/24
  (ii) Linear: cap = a + b×N
  (iii) 1/N decay: cap = c/N (rejected: 1×4=4 ≠ 9×2.5=22.5)

Test: for each N, find drop threshold by sweeping content_w around
expected boundary. Drop threshold / N = per-yak cap.

Probe: 24-char line, N yakumono evenly distributed (not at line-start
position 1).
"""
import json, os, sys, time, zipfile, shutil, tempfile, subprocess
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath(r"C:\Users\ryuji\oxi-1\tools\metrics\per_yak_cap_repro")
RESULT = os.path.abspath(r"C:\Users\ryuji\oxi-1\pipeline_data\per_yak_cap_sweep.json")
os.makedirs(OUT_DIR, exist_ok=True)

YAKUMONO = set("「")
PROBE_LEN = 24    # 24 CJK chars total
FONT = "ＭＳ 明朝"
SIZE_HALF = 24    # w:sz val=24 → 12pt
NATURAL = PROBE_LEN * 12.0  # 288pt


def make_probe(n_yak: int) -> str:
    """24-char probe with n_yak yakumono ('「') evenly spaced (not at pos 1)."""
    chars = ["漢"] * PROBE_LEN
    if n_yak == 0:
        return "".join(chars)
    # Place yak at positions: divide 24-char line into n_yak+1 segments,
    # place yak at end of each segment (skip line-start)
    # E.g. N=2 → positions 8 and 16
    # E.g. N=3 → positions 6, 12, 18
    # Make sure first yak NOT at position 1
    step = (PROBE_LEN - 2) // n_yak   # leave space at edges
    positions = []
    for i in range(n_yak):
        pos = (i + 1) * step + 1  # 1-indexed; first yak at pos step+1 ≥ 2
        if pos < PROBE_LEN:
            positions.append(pos - 1)  # convert to 0-index
    for p in positions[:n_yak]:
        chars[p] = "「"
    return "".join(chars)


def make_doc_xml(probe, jc, page_w_tw, margin_tw):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:body><w:p>'
            f'<w:pPr><w:jc w:val="{jc}"/>'
            '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            '</w:pPr>'
            '<w:r><w:rPr>'
            f'<w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:eastAsia="{FONT}" w:hint="eastAsia"/>'
            f'<w:sz w:val="{SIZE_HALF}"/></w:rPr>'
            f'<w:t>{probe}</w:t></w:r></w:p>'
            f'<w:sectPr><w:pgSz w:w="{page_w_tw}" w:h="16838"/>'
            f'<w:pgMar w:top="1134" w:right="{margin_tw}" w:bottom="1134" w:left="{margin_tw}"'
            ' w:header="720" w:footer="720" w:gutter="0"/>'
            '<w:cols w:space="720"/>'
            '<w:docGrid w:linePitch="360"/>'
            '</w:sectPr></w:body></w:document>')


def make_settings_xml():
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:characterSpacingControl w:val="compressPunctuation"/>'
            '</w:settings>')


def make_docx(label, probe, content_w_pt, jc="both"):
    """Direct-zip docx with explicit cSC=compressPunctuation (no Word.Documents.Add)."""
    out_path = os.path.join(OUT_DIR, f"{label}.docx")
    page_w_tw = int((content_w_pt + 170) * 20)
    margin_tw = 170 * 10  # 85pt each side
    doc_xml = make_doc_xml(probe, jc, page_w_tw, margin_tw)
    settings_xml = make_settings_xml()
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml"'
        ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '<Override PartName="/word/settings.xml"'
        ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
        '</Types>'
    )
    pkg_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"'
        ' Target="word/document.xml"/>'
        '</Relationships>'
    )
    doc_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"'
        ' Target="settings.xml"/>'
        '</Relationships>'
    )
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", pkg_rels)
        z.writestr("word/_rels/document.xml.rels", doc_rels)
        z.writestr("word/document.xml", doc_xml)
        z.writestr("word/settings.xml", settings_xml)
    return out_path


def kill_word():
    subprocess.run(['taskkill','/F','/IM','WINWORD.EXE'], capture_output=True)
    time.sleep(2)


def measure_one(path):
    word = w32.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    try:
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
        finally:
            try: d.Close(SaveChanges=False)
            except: pass
        if not xs: return {"error": "no chars"}
        # Group line 1
        y0 = xs[0][2]
        line1 = sorted([(t, x, sz) for t, x, y, sz in xs if abs(y - y0) < 0.5],
                       key=lambda v: v[1])
        n_line1 = len(line1)
        # Compute total compression on line 1
        total_comp = 0.0
        n_yak_total = 0
        n_yak_compressed = 0
        for i in range(len(line1) - 1):
            t = line1[i][0]
            adv = round(line1[i+1][1] - line1[i][1], 3)
            sz = line1[i][2]
            if t in YAKUMONO:
                n_yak_total += 1
                if sz > 0 and adv < sz * 0.99:
                    n_yak_compressed += 1
                    total_comp += (sz - adv)
        return {
            "n_chars_line1": n_line1,
            "total_compression_pt": round(total_comp, 3),
            "n_yak_total_line1": n_yak_total,
            "n_yak_compressed": n_yak_compressed,
        }
    finally:
        try: word.Quit()
        except: pass


def sweep_for_drop(n_yak, jc="both"):
    probe = make_probe(n_yak)
    # Sweep content_w from natural to natural - 30pt
    cw_values = [NATURAL - k for k in [-1, 0, 1, 2, 3, 4, 5, 6, 7, 8, 10, 12, 14, 16, 20, 25]]
    results = []
    for cw in cw_values:
        label = f"N{n_yak}_cw{cw:.0f}_jc{jc}"
        try:
            p = make_docx(label, probe, cw, jc)
        except Exception as e:
            results.append({"cw": cw, "build_error": str(e)})
            continue
        kill_word()
        try:
            r = measure_one(p)
        except Exception as e:
            r = {"measure_error": str(e)}
            kill_word()
        slack = NATURAL - cw
        results.append({"cw": cw, "slack": slack, **r})
        n = r.get("n_chars_line1", "?")
        comp = r.get("total_compression_pt", "?")
        print(f"  N={n_yak} cw={cw:.1f} slack={slack:+.1f} n_line1={n} total_comp={comp}")
    return {
        "n_yak": n_yak,
        "probe": probe,
        "natural": NATURAL,
        "sweep": results,
    }


def main():
    out = {}
    for n_yak in [2, 3, 4, 5, 7]:
        print(f"\n=== N={n_yak} ===")
        out[f"N{n_yak}"] = sweep_for_drop(n_yak, "both")

    os.makedirs(os.path.dirname(RESULT), exist_ok=True)
    with open(RESULT, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)

    print("\n=== Summary: drop threshold per N ===")
    print(f"{'N':>3} | {'natural':>8} | {'max_n_chars_24':>14} | {'max_total_comp':>14} | {'per_yak_cap':>11}")
    for n_yak in [2, 3, 4, 5, 7]:
        sweep = out[f"N{n_yak}"]["sweep"]
        # Find max slack where line still has all 24 chars
        # AND get the corresponding total_compression
        max_comp_at_24 = 0
        for r in sweep:
            if r.get("n_chars_line1") == 24:
                tc = r.get("total_compression_pt", 0)
                if tc > max_comp_at_24:
                    max_comp_at_24 = tc
        per_yak = max_comp_at_24 / n_yak if n_yak > 0 else 0
        print(f"{n_yak:>3} | {NATURAL:>8.1f} | {'24':>14} | {max_comp_at_24:>14.3f} | {per_yak:>11.4f}pt")


if __name__ == "__main__":
    main()
