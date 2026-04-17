"""Detailed OMML COM measurement: fraction geometry + italic-math substitution.

Part 1: Fraction bar position and numerator/denominator offsets.
  Measures per-char x/y for 01_frac.docx to derive:
  - numerator baseline y vs math paragraph top
  - denominator baseline y
  - implicit bar y (midway)
  - horizontal centering (num vs den x)

Part 2: Italic-math Latin substitution enumeration.
  Tests each ASCII letter [A-Za-z] in a math run; records the
  substituted character (U+1D...) to build a lookup table.

Output:
  tools/metrics/output/omml_fraction_geometry.json
  tools/metrics/output/omml_italic_math_table.json
"""
import json, os, sys, time, zipfile
from pathlib import Path
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT1 = Path(__file__).with_name("output") / "omml_fraction_geometry.json"
OUT2 = Path(__file__).with_name("output") / "omml_italic_math_table.json"
OUT1.parent.mkdir(parents=True, exist_ok=True)

TMP = Path("pipeline_data") / "_omml_detail_tmp.docx"
TMP.parent.mkdir(parents=True, exist_ok=True)

NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'
CT = '<?xml version="1.0"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
RELS = '<?xml version="1.0"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'


def build_docx(body_xml: str):
    xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:document {NS}>'
        f'<w:body>{body_xml}'
        '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>'
        '</w:body></w:document>'
    )
    with zipfile.ZipFile(TMP, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/document.xml", xml)


def measure_fraction(word):
    """Measure 01_frac (a/b) per-char x,y to extract bar + num/den geometry.
    We build the fraction with distinct num+den chars to spot them easily.
    """
    body = (
        '<w:p><w:pPr><w:rPr><w:sz w:val="24"/></w:rPr></w:pPr>'
        '<w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>BEFORE</w:t></w:r></w:p>'
        '<w:p><w:pPr><w:rPr><w:sz w:val="24"/></w:rPr></w:pPr>'
        '<m:oMathPara><m:oMathParaPr><m:jc m:val="center"/></m:oMathParaPr><m:oMath>'
        '<m:f><m:num><m:r><m:t>A</m:t></m:r></m:num>'
        '<m:den><m:r><m:t>B</m:t></m:r></m:den></m:f>'
        '</m:oMath></m:oMathPara></w:p>'
        '<w:p><w:pPr><w:rPr><w:sz w:val="24"/></w:rPr></w:pPr>'
        '<w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>AFTER</w:t></w:r></w:p>'
    )
    build_docx(body)
    try:
        doc = word.Documents.Open(str(TMP.resolve()), ReadOnly=True)
        time.sleep(0.5)
        # Paragraphs: BEFORE (1), math (2), AFTER (3)
        sel = word.Selection
        rng_before = doc.Paragraphs(1).Range
        rng_math = doc.Paragraphs(2).Range
        rng_after = doc.Paragraphs(3).Range

        y_before = rng_before.Information(6)
        y_math = rng_math.Information(6)
        y_after = rng_after.Information(6)

        # Enumerate each character of the math paragraph
        chars = []
        for ci in range(rng_math.Start, rng_math.End):
            sel.SetRange(ci, ci + 1)
            try:
                y = sel.Information(6)
                x = sel.Information(5)
                ch = sel.Text
            except Exception:
                continue
            if ch.strip():
                # Word may return surrogate pairs for U+1D4xx math italic chars.
                # Use single codepoint if len==1 OR interpret as surrogate pair.
                cp = None
                if len(ch) == 1:
                    cp = ord(ch)
                elif len(ch) == 2 and 0xD800 <= ord(ch[0]) <= 0xDBFF:
                    # high surrogate + low surrogate → full codepoint
                    cp = 0x10000 + (ord(ch[0]) - 0xD800) * 0x400 + (ord(ch[1]) - 0xDC00)
                chars.append({"ci": ci - rng_math.Start, "ch": ch, "x": round(x, 2), "y": round(y, 2), "cp": cp})

        doc.Close(False)
        result = {
            "y_before": round(y_before, 2),
            "y_math_top": round(y_math, 2),
            "y_after": round(y_after, 2),
            "math_height": round(y_after - y_math, 2),
            "chars": chars,
        }
        # Compute num/den y offsets
        if len(chars) >= 2:
            a = chars[0]  # num (A)
            b = chars[1]  # den (B)
            result["num_y"] = a["y"]
            result["den_y"] = b["y"]
            result["num_x"] = a["x"]
            result["den_x"] = b["x"]
            result["num_den_dy"] = round(b["y"] - a["y"], 2)
            result["num_den_dx"] = round(b["x"] - a["x"], 2)
            result["bar_y_estimated"] = round((a["y"] + b["y"]) / 2, 2)  # rough
        return result
    except Exception as e:
        return {"error": str(e)}


def measure_italic_math(word):
    """Test each ASCII letter in math run; record Word's substituted char."""
    results = {}
    letters = list("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789")
    # Build one doc with all letters in one math run
    math_content = ''.join(f'<m:r><m:t>{c}</m:t></m:r>' for c in letters)
    body = (
        '<w:p><w:pPr><w:rPr><w:sz w:val="24"/></w:rPr></w:pPr>'
        '<m:oMathPara><m:oMath>'
        f'{math_content}'
        '</m:oMath></m:oMathPara></w:p>'
    )
    build_docx(body)
    try:
        doc = word.Documents.Open(str(TMP.resolve()), ReadOnly=True)
        time.sleep(0.5)
        rng = doc.Paragraphs(1).Range
        # Iterate each char; Range.Text contains the rendered substituted chars
        text = rng.Text
        # Character-level via selection (paired up input vs output)
        sel = word.Selection
        rendered_chars = []
        # Step through char positions, combining surrogate pairs.
        ci = rng.Start
        while ci < rng.End:
            sel.SetRange(ci, ci + 1)
            ch = sel.Text
            if not ch or ch in ('\r', '\x07', '\n'):
                ci += 1
                continue
            if 0xD800 <= ord(ch[0]) <= 0xDBFF:
                # High surrogate — grab next position as low surrogate
                sel.SetRange(ci + 1, ci + 2)
                low = sel.Text
                if low and 0xDC00 <= ord(low[0]) <= 0xDFFF:
                    cp = 0x10000 + (ord(ch[0]) - 0xD800) * 0x400 + (ord(low[0]) - 0xDC00)
                    rendered_chars.append({"ch": ch + low, "cp": cp})
                    ci += 2
                    continue
            cp = ord(ch[0])
            rendered_chars.append({"ch": ch, "cp": cp})
            ci += 1

        doc.Close(False)
        # Pair each input letter to rendered char (by order)
        if len(rendered_chars) == len(letters):
            for i, c in enumerate(letters):
                r = rendered_chars[i]
                results[c] = {
                    "input_cp": ord(c),
                    "rendered_ch": r["ch"],
                    "rendered_cp": r["cp"],
                    "substituted": r["cp"] != ord(c),
                }
        else:
            results["_mismatch"] = {
                "n_letters": len(letters),
                "n_rendered": len(rendered_chars),
                "rendered": rendered_chars,
            }
        return results
    except Exception as e:
        return {"error": str(e)}


def main():
    word = None
    for attempt in range(5):
        try:
            word = win32com.client.Dispatch("Word.Application")
            time.sleep(2.0)
            word.Visible = False
            word.DisplayAlerts = False
            break
        except Exception as e:
            print(f"  Word COM attempt {attempt+1} err: {e}")
            time.sleep(10 * (attempt + 1))
    if word is None:
        print("Word COM unavailable after 5 retries. Bailing.")
        return

    try:
        # Part 1: fraction geometry
        print("=== Part 1: fraction (A/B) geometry ===")
        frac = measure_fraction(word)
        if "error" in frac:
            print(f"  ERR: {frac['error']}")
        else:
            print(f"  y_math_top: {frac['y_math_top']}")
            print(f"  math_height: {frac['math_height']}")
            cp0 = frac['chars'][0]['cp']
            cp1 = frac['chars'][1]['cp']
            print(f"  num: x={frac['num_x']} y={frac['num_y']} ch={frac['chars'][0]['ch']!r} cp={f'U+{cp0:04X}' if cp0 else 'None'}")
            print(f"  den: x={frac['den_x']} y={frac['den_y']} ch={frac['chars'][1]['ch']!r} cp={f'U+{cp1:04X}' if cp1 else 'None'}")
            print(f"  num→den dy={frac['num_den_dy']} dx={frac['num_den_dx']}")
            print(f"  bar y (est): {frac['bar_y_estimated']}")
        with open(OUT1, "w", encoding="utf-8") as f:
            json.dump(frac, f, ensure_ascii=True, indent=2)
        print(f"  Saved → {OUT1}")

        # Part 2: italic-math substitution table
        print("\n=== Part 2: italic-math Latin substitution ===")
        subs = measure_italic_math(word)
        if "error" in subs:
            print(f"  ERR: {subs['error']}")
        elif "_mismatch" in subs:
            print(f"  MISMATCH: {subs['_mismatch']}")
        else:
            # Group by substituted or not
            substituted = [(c, d) for c, d in subs.items() if d.get("substituted")]
            unchanged = [(c, d) for c, d in subs.items() if not d.get("substituted")]
            print(f"  Substituted: {len(substituted)} chars")
            print(f"  Unchanged:   {len(unchanged)} chars")
            print(f"\n  Substitution table (input → rendered):")
            for c, d in sorted(subs.items()):
                if "substituted" not in d: continue
                mark = "→" if d["substituted"] else " "
                print(f"    '{c}' (U+{d['input_cp']:04X}) {mark} '{d['rendered_ch']}' (U+{d['rendered_cp']:04X})")
        with open(OUT2, "w", encoding="utf-8") as f:
            json.dump(subs, f, ensure_ascii=True, indent=2)
        print(f"  Saved → {OUT2}")
    finally:
        try: word.Quit()
        except: pass
        try: os.remove(TMP)
        except: pass


if __name__ == "__main__":
    main()
