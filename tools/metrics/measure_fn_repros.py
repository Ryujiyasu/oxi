"""COM-measure the self-authored fn repro scenarios.

For each RA/RB/RC/RD docx, record per-page:
  - last body para y, y+line_h (body_bot)
  - fn body y per fn (sorted)
  - gap = fn_first_y - body_bot
  - which paragraphs contain fn refs (by idx) and which page they rendered on
"""
import os, sys, time, json, glob
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPRO_DIR = r"tools\metrics\fn_reserve_repro"
LINE_H_DEFAULT = 18.0

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
try:
    paths = sorted(glob.glob(os.path.join(REPRO_DIR, "R*.docx")))
    results = []
    for p in paths:
        name = os.path.basename(p)
        print(f"\n=== {name} ===")
        doc = word.Documents.Open(os.path.abspath(p), ReadOnly=True)
        time.sleep(0.3)
        try:
            total_pg = int(doc.ComputeStatistics(2))
            paras_by_page = {}
            for i, pp in enumerate(doc.Paragraphs, 1):
                try:
                    pg = pp.Range.Information(3)
                    y = pp.Range.Information(6)
                    refs = [r.Index for r in pp.Range.Footnotes]
                    paras_by_page.setdefault(pg, []).append({
                        "idx": i, "y": round(y, 1), "refs": refs,
                        "text": pp.Range.Text[:30].replace('\r','').replace('\n',' ').replace('\x07',''),
                    })
                except Exception:
                    pass
            fns_by_page = {}
            for fn in doc.Footnotes:
                try:
                    rp = fn.Reference.Information(3)
                    bp = fn.Range.Information(3)
                    ry = fn.Reference.Information(6)
                    by = fn.Range.Information(6)
                    fns_by_page.setdefault(bp, []).append({
                        "seq": fn.Index, "ref_pg": rp, "ref_y": round(ry, 1),
                        "body_y": round(by, 1),
                    })
                except Exception:
                    pass
            print(f"  total_pg={total_pg}, fn_total={doc.Footnotes.Count}")
            for pg in sorted(paras_by_page.keys()):
                body = sorted(paras_by_page[pg], key=lambda x: x["y"])
                if not body:
                    continue
                last = body[-1]
                fns = sorted(fns_by_page.get(pg, []), key=lambda x: x["body_y"])
                body_bot = last["y"] + LINE_H_DEFAULT
                fn_top = fns[0]["body_y"] if fns else None
                gap = (fn_top - body_bot) if fn_top is not None else None
                print(f"  p{pg}: n_paras={len(body)} last_idx={last['idx']} last_y={last['y']} last_refs={last['refs']} body_bot~{body_bot} fn_top={fn_top} gap={gap}")
                for fn in fns:
                    print(f"         fn seq={fn['seq']} ref_pg={fn['ref_pg']} ref_y={fn['ref_y']} body_y={fn['body_y']}")
            results.append({"doc": name, "paras_by_page": paras_by_page, "fns_by_page": fns_by_page})
        finally:
            doc.Close(False)
    # Save
    out = os.path.join("pipeline_data", "fn_repro_measurements.json")
    with open(out, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f"\nSaved {out}")
finally:
    word.Quit()
