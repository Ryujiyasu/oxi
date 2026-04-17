"""Measure Word's last paragraph on d77a p.2 and first paragraph on p.3.

Iterate body paragraphs; get their page via Range.Information(3).
Report the boundary so we can compare with Oxi dump.
"""
import os, sys, time, json
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
# IMPORTANT: don't lose encoding in JSON — keep raw text without replace

docx = os.path.abspath(
    r"tools\golden-test\documents\docx\d77a58485f16_20240705_resources_data_outline_08.docx"
)

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
try:
    doc = word.Documents.Open(docx, ReadOnly=True); time.sleep(0.5)
    paras = list(doc.Paragraphs)
    print(f"Total body paras: {len(paras)}\n")

    # Collect page+y per para
    data = []
    for i, p in enumerate(paras, 1):
        try:
            pg = p.Range.Information(3)
            y = p.Range.Information(6)
            txt = p.Range.Text[:50].replace('\r', '').replace('\n', ' ').replace('\x07', '')
            # Page break before?
            pb = False
            try:
                pb = p.PageBreakBefore
            except Exception:
                pass
            # Space before/after
            sb = p.Format.SpaceBefore
            sa = p.Format.SpaceAfter
            ln = p.Format.LineSpacing
            lr = p.Format.LineSpacingRule
            data.append({
                "idx": i, "page": int(pg), "y": round(y, 1),
                "text": txt, "pb": pb,
                "sb": round(sb, 1), "sa": round(sa, 1),
                "ls": round(ln, 1), "lr": lr,
            })
        except Exception as e:
            data.append({"idx": i, "error": str(e)})

    # Print paras around p2/p3 boundary
    print("Paragraphs on p1-p4:")
    for d in data:
        if d.get("page") in (1, 2, 3) and "error" not in d:
            print(f"  p{d['page']} idx={d['idx']:<3} y={d['y']:<7} sb={d['sb']:<5} sa={d['sa']:<5} ls={d['ls']} lr={d['lr']} pb={d['pb']} {d['text'][:50]!r}")

    # Save
    with open("pipeline_data/d77a_word_p1_p3_paras.json", "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    doc.Close(False)
finally:
    word.Quit()
