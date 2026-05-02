"""
§11.5 evenPage / oddPage section break behavior.

Variants:
  - wdSectionBreakEvenPage = 4  → s2 starts on next EVEN page (blank if necessary)
  - wdSectionBreakOddPage  = 5  → s2 starts on next ODD page (blank if necessary)

For each variant, test:
  - s1 ends on odd page 1 → break to even (blank page 2 if oddPage break? no)
  - s1 ends on even page → ...
  - Header/footer on inserted blank page

Tests:
  T1: s1 fills 1 page (page 1, odd), oddPage break → s2 starts page 3 (skip 2)
  T2: s1 fills 1 page (page 1, odd), evenPage break → s2 starts page 2
  T3: s1 fills 2 pages (page 1+2), oddPage break → s2 starts page 3
  T4: s1 fills 2 pages (page 1+2), evenPage break → s2 starts page 4 (skip 3)
"""
import os
import sys
import time
import json

import win32com.client
import pythoncom

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.join(os.path.dirname(__file__), "output")
OUT_JSON = os.path.join(OUT_DIR, "ra2_even_odd_page_break.json")
os.makedirs(OUT_DIR, exist_ok=True)


WD_LINE_SPACE_SINGLE = 0
WD_SECTION_BREAK_EVEN_PAGE = 4
WD_SECTION_BREAK_ODD_PAGE = 5
WD_SECTION_BREAK_NEXT_PAGE = 2

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


def build_doc(word, *, n_s1_paras, break_type, n_s2_paras=3):
    wdoc = retry(word.Documents.Add)
    sec1 = retry(lambda: wdoc.Sections(1))
    ps = sec1.PageSetup
    ps.LeftMargin = 72
    ps.RightMargin = 72
    ps.TopMargin = 72
    ps.BottomMargin = 72
    ps.HeaderDistance = 36
    ps.FooterDistance = 36

    # Section 1 body — enough lines per paragraph or many paragraphs
    s1_text = "\r".join(f"S1_B{i+1}" for i in range(n_s1_paras))
    wdoc.Content.Text = s1_text
    for i in range(1, wdoc.Paragraphs.Count + 1):
        p = wdoc.Paragraphs(i)
        p.Range.Font.Name = "Calibri"
        p.Range.Font.Size = 11
        p.Format.LineSpacingRule = WD_LINE_SPACE_SINGLE
        p.Format.SpaceBefore = 0
        p.Format.SpaceAfter = 0

    # Insert section break
    end = wdoc.Content.End - 1
    wdoc.Range(end, end).InsertBreak(break_type)

    # Section 2 body
    end = wdoc.Content.End - 1
    s2_text = "\r".join(f"S2_B{i+1}" for i in range(n_s2_paras))
    wdoc.Range(end, end).InsertAfter(s2_text)
    for i in range(1, wdoc.Paragraphs.Count + 1):
        p = wdoc.Paragraphs(i)
        p.Range.Font.Name = "Calibri"
        p.Range.Font.Size = 11
        p.Format.LineSpacingRule = WD_LINE_SPACE_SINGLE
        p.Format.SpaceBefore = 0
        p.Format.SpaceAfter = 0

    wdoc.Repaginate()
    time.sleep(0.05)

    # Measure: pages each paragraph lands on, total page count
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

    # Total page count — find max page in paragraphs (Word doesn't expose total directly)
    n_pages = max((p["page"] for p in paras), default=0)
    rec = {
        "break_type": break_type,
        "n_s1_paras": n_s1_paras,
        "n_s2_paras": n_s2_paras,
        "n_pages_found": n_pages,
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
        ("nextPage_s1_3paras",  3,  WD_SECTION_BREAK_NEXT_PAGE),
        ("nextPage_s1_50paras", 50, WD_SECTION_BREAK_NEXT_PAGE),
        ("oddPage_s1_3paras",   3,  WD_SECTION_BREAK_ODD_PAGE),
        ("oddPage_s1_50paras",  50, WD_SECTION_BREAK_ODD_PAGE),
        ("evenPage_s1_3paras",  3,  WD_SECTION_BREAK_EVEN_PAGE),
        ("evenPage_s1_50paras", 50, WD_SECTION_BREAK_EVEN_PAGE),
    ]

    results = []
    try:
        for name, n_s1, btype in cases:
            try:
                r = build_doc(word, n_s1_paras=n_s1, break_type=btype)
                r["name"] = name
                results.append(r)
                print(f"\n=== {name} (n_s1={n_s1}, break_type={btype}) ===")
                print(f"  Total pages: {r['n_pages_found']}")
                # Show first/last s1 + first s2
                s1_paras = [p for p in r["paragraphs"] if p["text"].startswith("S1_")]
                s2_paras = [p for p in r["paragraphs"] if p["text"].startswith("S2_")]
                if s1_paras:
                    print(f"  S1 first: P{s1_paras[0]['i']:3} @page{s1_paras[0]['page']} y={s1_paras[0]['y']}")
                    print(f"  S1 last:  P{s1_paras[-1]['i']:3} @page{s1_paras[-1]['page']} y={s1_paras[-1]['y']}")
                if s2_paras:
                    print(f"  S2 first: P{s2_paras[0]['i']:3} @page{s2_paras[0]['page']} y={s2_paras[0]['y']}")
                    print(f"  S2 last:  P{s2_paras[-1]['i']:3} @page{s2_paras[-1]['page']} y={s2_paras[-1]['y']}")
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
