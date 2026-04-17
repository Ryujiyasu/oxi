"""Measure Word footnote text line pitch on b837.

Iterate footnote 3 (long, multi-line) and get char Y positions.
"""
import os, sys, time, json
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

docx = os.path.abspath(
    r"tools\golden-test\documents\docx\b837808d0555_20240705_resources_data_guideline_02.docx"
)

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
try:
    doc = word.Documents.Open(docx, ReadOnly=True); time.sleep(0.5)
    fns = doc.Footnotes
    print(f"Total footnotes: {fns.Count}")

    results = []
    for i in range(1, min(fns.Count, 10) + 1):
        try:
            fn = fns(i)
            rng = fn.Range
            # Step every char
            n_chars = rng.Characters.Count
            if n_chars < 1: continue
            # Sample multiple chars to detect line boundaries
            ys = []
            step = max(1, n_chars // 50)
            for ci in range(1, n_chars + 1, step):
                try:
                    y = rng.Characters(ci).Information(6)
                    ys.append(round(y, 1))
                except: pass
            unique = sorted(set(ys))
            # Line pitch = median consecutive diff
            diffs = [unique[j+1] - unique[j] for j in range(len(unique)-1)]
            diffs = [d for d in diffs if d > 2]  # filter same-line
            txt = rng.Text[:30].replace('\r','').replace('\n',' ').replace('\x07','')
            print(f"FN{i}: n_chars={n_chars} unique_y={len(unique)} diffs={diffs[:6]} text={txt!r}")
            results.append({"idx": i, "n_chars": n_chars, "unique_y": unique, "pitches": diffs, "text": txt})
        except Exception as e:
            print(f"FN{i} ERROR: {e}")

    with open("pipeline_data/b837_word_footnote_pitch.json", "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)

    doc.Close(False)
finally:
    word.Quit()
