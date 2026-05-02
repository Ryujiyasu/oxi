"""
§11.2 nextPage section break refinement.

Build fixtures via Word COM (avoids python-docx XML quirks). Each fixture has:
  - Section 1: a few body paragraphs
  - Section 2 (nextPage break): different settings per fixture

Variables to sweep:
  - s2 topMargin (default 72 vs 108 vs 144)
  - s2 leftMargin (default 72 vs 144)
  - s2 pageSize (Letter vs A4)
  - s2 columns (1 vs 2)
  - s2 headerDistance (default 36 vs 72)
  - s2 different first-page header

Measure: s2 first body paragraph (x, y) on page 2, and confirm:
  s2.first.y == s2.topMargin (no header overflow)
  s2.first.x == s2.leftMargin
  Etc.
"""
import os
import sys
import time
import json

import win32com.client
import pythoncom

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.join(os.path.dirname(__file__), "output")
FIX_DIR = os.path.join(OUT_DIR, "nextpage_section_fixtures")
os.makedirs(FIX_DIR, exist_ok=True)
OUT_JSON = os.path.join(OUT_DIR, "ra2_nextpage_section.json")


WD_LINE_SPACE_SINGLE = 0
WD_SECTION_BREAK_NEXTPAGE = 2  # wdSectionBreakNextPage

RPC_REJECTED_CODES = {-2147418111, -2147023174, -2147023170}

def retry(fn, *args, retries=15, delay=0.3, **kwargs):
    last_exc = None
    for i in range(retries):
        try:
            return fn(*args, **kwargs)
        except Exception as e:
            last_exc = e
            code = e.args[0] if hasattr(e, "args") and len(e.args) >= 1 else None
            if code in RPC_REJECTED_CODES or "rejected" in str(e).lower():
                pythoncom.PumpWaitingMessages()
                time.sleep(delay * (1.3 ** i))
                continue
            raise
    raise last_exc


def add_section1(wdoc, n_body_paras=3):
    sec1 = wdoc.Sections(1)
    ps = sec1.PageSetup
    ps.LeftMargin = 72
    ps.RightMargin = 72
    ps.TopMargin = 72
    ps.BottomMargin = 72
    ps.HeaderDistance = 36
    ps.FooterDistance = 36

    wdoc.Content.Text = "\r".join(f"S1_B{i+1}" for i in range(n_body_paras))
    for i in range(1, n_body_paras + 1):
        p = wdoc.Paragraphs(i)
        p.Range.Font.Name = "Calibri"
        p.Range.Font.Size = 11
        p.Format.LineSpacingRule = WD_LINE_SPACE_SINGLE
        p.Format.SpaceBefore = 0
        p.Format.SpaceAfter = 0


def append_nextpage_break(wdoc):
    end = wdoc.Content.End - 1
    rng = wdoc.Range(end, end)
    rng.InsertBreak(WD_SECTION_BREAK_NEXTPAGE)


def append_section2_body(wdoc, n_body_paras=3):
    end = wdoc.Content.End - 1
    rng = wdoc.Range(end, end)
    rng.InsertAfter("\r".join(f"S2_B{i+1}" for i in range(n_body_paras)))
    # Reformat all paragraphs to ensure consistent style
    for i in range(1, wdoc.Paragraphs.Count + 1):
        p = wdoc.Paragraphs(i)
        p.Range.Font.Name = "Calibri"
        p.Range.Font.Size = 11
        p.Format.LineSpacingRule = WD_LINE_SPACE_SINGLE
        p.Format.SpaceBefore = 0
        p.Format.SpaceAfter = 0


def configure_section2(wdoc, *, top_margin=72, left_margin=72,
                        right_margin=72, header_distance=36,
                        page_width=None, page_height=None,
                        n_columns=1):
    sec2 = retry(lambda: wdoc.Sections(2))
    ps = sec2.PageSetup
    ps.LeftMargin = left_margin
    ps.RightMargin = right_margin
    ps.TopMargin = top_margin
    ps.BottomMargin = 72
    ps.HeaderDistance = header_distance
    ps.FooterDistance = 36
    if page_width is not None:
        ps.PageWidth = page_width
    if page_height is not None:
        ps.PageHeight = page_height
    if n_columns > 1:
        ps.TextColumns.SetCount(n_columns)


def build_and_measure(word, name, **s2_kwargs):
    wdoc = retry(word.Documents.Add)
    add_section1(wdoc)
    append_nextpage_break(wdoc)
    append_section2_body(wdoc)
    configure_section2(wdoc, **s2_kwargs)
    wdoc.Repaginate()
    time.sleep(0.05)

    # Measure
    paras = []
    for i in range(1, wdoc.Paragraphs.Count + 1):
        p = wdoc.Paragraphs(i)
        try:
            paras.append({
                "i": i,
                "y": round(p.Range.Information(6), 4),
                "x": round(p.Range.Information(5), 4),
                "page": p.Range.Information(3),
                "text": p.Range.Text.strip()[:20],
            })
        except Exception:
            pass

    sec_count = wdoc.Sections.Count
    sec_props = []
    for s in range(1, sec_count + 1):
        sec = wdoc.Sections(s)
        sec_props.append({
            "i": s,
            "topMargin": round(sec.PageSetup.TopMargin, 4),
            "leftMargin": round(sec.PageSetup.LeftMargin, 4),
            "headerDistance": round(sec.PageSetup.HeaderDistance, 4),
            "pageWidth": round(sec.PageSetup.PageWidth, 4),
            "pageHeight": round(sec.PageSetup.PageHeight, 4),
            "n_columns": sec.PageSetup.TextColumns.Count,
        })

    rec = {
        "name": name,
        "s2_kwargs": s2_kwargs,
        "section_count": sec_count,
        "sections": sec_props,
        "paragraphs": paras,
    }
    wdoc.Close(False)
    return rec


def main():
    word = win32com.client.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    time.sleep(2.0)
    try:
        retry(lambda: word.Documents.Count)
    except Exception:
        pass

    cases = [
        ("baseline_default", {}),
        ("s2_topmargin_108", {"top_margin": 108}),
        ("s2_topmargin_144", {"top_margin": 144}),
        ("s2_topmargin_36",  {"top_margin": 36}),
        ("s2_leftmargin_144", {"left_margin": 144}),
        ("s2_headerdist_72", {"header_distance": 72}),
        ("s2_pagesize_A4",   {"page_width": 595, "page_height": 842}),  # A4 portrait
        ("s2_2col",          {"n_columns": 2}),
        ("s2_topmargin_108_leftmargin_144", {"top_margin": 108, "left_margin": 144}),
    ]

    results = []
    try:
        for name, kw in cases:
            try:
                r = build_and_measure(word, name, **kw)
                results.append(r)
                print(f"\n=== {name} ===")
                print(f"  s2_kwargs: {kw}")
                for sp in r["sections"]:
                    print(f"  sec{sp['i']}: tm={sp['topMargin']} lm={sp['leftMargin']} hd={sp['headerDistance']} cols={sp['n_columns']} pgW={sp['pageWidth']}x{sp['pageHeight']}")
                # Find first S2 paragraph
                s2_first = None
                for p in r["paragraphs"]:
                    if p.get("text", "").startswith("S2_"):
                        s2_first = p
                        break
                if s2_first:
                    print(f"  S2 first para: P{s2_first['i']}@page{s2_first['page']} y={s2_first['y']} x={s2_first['x']} '{s2_first['text']}'")
            except Exception as e:
                print(f"  ERR {name}: {e}")
    finally:
        with open(OUT_JSON, "w", encoding="utf-8") as f:
            json.dump(results, f, indent=2, ensure_ascii=False)
        print(f"\nSaved {len(results)} records to {OUT_JSON}")
        try:
            word.Quit()
        except Exception:
            pass


if __name__ == "__main__":
    main()
