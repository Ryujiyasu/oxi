"""COM-measure Word b837 p.4 body bottom and footnote area top.

Goal: derive Word's reserve algorithm.
- body_bot: last body paragraph Y + line height
- fn_area_top: first footnote paragraph Y (if we can reach them via Word.Footnotes)
- gap = fn_area_top - body_bot

Compare with Oxi Step 1 partial (current main, body_bot=468.5pt, fn area at 656.5pt)
and Oxi fixed-point FALSIFIED (body_bot=558.5pt).
"""
import os, sys, time, json
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOCX = os.path.abspath(
    r"tools\golden-test\documents\docx\b837808d0555_20240705_resources_data_guideline_02.docx"
)

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
try:
    doc = word.Documents.Open(DOCX, ReadOnly=True); time.sleep(0.5)

    # Page 4 body paragraphs
    print("=== Word b837 p.4 body paragraphs ===")
    p4_paras = []
    for i, p in enumerate(doc.Paragraphs, 1):
        try:
            if p.Range.Information(3) == 4:
                y = p.Range.Information(6)  # vertical pos relative to page
                text = p.Range.Text[:60].replace('\r','').replace('\n',' ').replace('\x07','')
                # Find fn refs inside this paragraph
                refs = []
                for r in p.Range.Footnotes:
                    refs.append(r.Index)
                p4_paras.append({"idx": i, "y": round(y, 2), "refs": refs, "text": text})
        except Exception as e:
            pass

    for p in p4_paras:
        refs_s = f" refs={p['refs']}" if p['refs'] else ""
        print(f"  idx={p['idx']:<4} y={p['y']:>7.2f}{refs_s}  {p['text'][:50]!r}")

    # Approx body bottom = last para y + line height (assume ~18pt)
    if p4_paras:
        last = p4_paras[-1]
        body_bot_approx = last['y'] + 18.0
        print(f"\n=== Body bottom approx (last para y=idx{last['idx']} + 18pt) ===")
        print(f"  body_bot ≈ {body_bot_approx:.1f}pt")

    # Try to find footnote paragraphs on p4 via doc.Footnotes
    print("\n=== Word b837 footnotes (all) ===")
    fn_on_p4 = []
    for fn in doc.Footnotes:
        try:
            ref_pg = fn.Reference.Information(3)
            # The footnote body text is in fn.Range; find its page
            body_pg = fn.Range.Information(3)
            body_y = fn.Range.Information(6)
            ref_y = fn.Reference.Information(6)
            first_text = fn.Range.Text[:40].replace('\r','').replace('\n',' ').replace('\x07','')
            row = {
                "seq": fn.Index,
                "ref_pg": ref_pg,
                "ref_y": round(ref_y, 2),
                "body_pg": body_pg,
                "body_y": round(body_y, 2),
                "text": first_text,
            }
            if ref_pg == 4 or body_pg == 4:
                fn_on_p4.append(row)
        except Exception as e:
            pass

    for fn in fn_on_p4:
        print(f"  seq={fn['seq']:<3} ref_pg={fn['ref_pg']} ref_y={fn['ref_y']:>7.2f}  body_pg={fn['body_pg']} body_y={fn['body_y']:>7.2f}  {fn['text'][:40]!r}")

    if fn_on_p4:
        p4_fn_body = sorted([fn['body_y'] for fn in fn_on_p4 if fn['body_pg'] == 4])
        if p4_fn_body:
            fn_area_top = p4_fn_body[0]
            fn_area_bot = p4_fn_body[-1]
            print(f"\n=== Word p4 fn area ===")
            print(f"  first fn body y = {fn_area_top:.2f}pt")
            print(f"  last  fn body y = {fn_area_bot:.2f}pt")
            if p4_paras:
                gap = fn_area_top - (p4_paras[-1]['y'] + 18.0)
                print(f"  gap between body_bot(≈) and fn_area_top = {gap:.2f}pt")

    # Summary
    print("\n=== Comparison (Word vs Oxi) ===")
    print(f"  Word p4 body last para y    = {p4_paras[-1]['y']:.2f} (idx={p4_paras[-1]['idx']})")
    print(f"  Oxi main p4 body_bot         = 468.5 (Step 1 partial)")
    print(f"  Oxi fixed-point body_bot     = 558.5 (FALSIFIED)")
    print(f"  Oxi delta vs Word para last  = {468.5 - p4_paras[-1]['y']:+.2f} (Oxi main vs Word idx{p4_paras[-1]['idx']})")

    doc.Close(False)
finally:
    word.Quit()
