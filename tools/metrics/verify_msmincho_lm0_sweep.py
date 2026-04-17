"""Sweep MS Mincho LM0 lineRule=auto gap for sizes that actual docs use."""
import win32com.client, time, sys, json
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False

results = {}
for size in [9.0, 9.5, 10.0, 10.5, 11.0, 12.0]:
    doc = word.Documents.Add()
    time.sleep(0.2)
    ps = doc.PageSetup
    ps.PageWidth = 595.3; ps.PageHeight = 841.9
    ps.LeftMargin = 99.25; ps.RightMargin = 99.25
    ps.TopMargin = 113.4; ps.BottomMargin = 113.4
    try:
        ps.LayoutMode = 0
    except Exception:
        pass
    rng = doc.Range()
    rng.InsertAfter("あいうえお\nかきくけこ\nさしすせそ")
    for p in [doc.Paragraphs(1), doc.Paragraphs(2), doc.Paragraphs(3)]:
        p.Range.Font.Name = "ＭＳ 明朝"
        p.Range.Font.Size = size
        p.LineSpacingRule = 0
        p.SpaceBefore = 0; p.SpaceAfter = 0
    time.sleep(0.2)
    y1 = doc.Paragraphs(1).Range.Information(6)
    y2 = doc.Paragraphs(2).Range.Information(6)
    y3 = doc.Paragraphs(3).Range.Information(6)
    g12 = y2 - y1
    g23 = y3 - y2
    results[f"{size:.1f}"] = {"P1-P2": round(g12,3), "P2-P3": round(g23,3)}
    print(f"size={size:<5}  P1_y={y1:.2f}  P2_y={y2:.2f}  P3_y={y3:.2f}  g12={g12:.2f}  g23={g23:.2f}")
    doc.Close(SaveChanges=False)

# Print existing table values for comparison
existing = {
    "9.0": 9.0, "9.5": 10.0, "10.0": 11.0, "10.5": 13.5, "11.0": 13.0, "12.0": 15.5
}
print("\nComparison (P2-P3 is the 'consecutive' single-spacing value):")
print(f"{'size':<6} {'table':<8} {'measured':<10} {'diff':>8}")
for s, r in results.items():
    t = existing.get(s, None)
    m = r["P2-P3"]
    d = (m - t) if t else 0
    mark = "!" if abs(d) > 0.1 else " "
    print(f"{mark} {s:<5} {t:<8} {m:<10} {d:+.2f}")

word.Quit()
