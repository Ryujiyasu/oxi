"""
Grid snap order systematic test.
Determines: snap(max(run, default)) vs max(snap(default), run)
"""
import win32com.client
import os, time, math

word = win32com.client.Dispatch('Word.Application')
word.Visible = False

for grid_lines in [30, 36, 40, 44]:
    doc = word.Documents.Add()
    time.sleep(0.5)
    
    sec = doc.Sections(1)
    pg = sec.PageSetup
    pg.LinesPage = grid_lines
    grid_twips = pg.LinePitch
    grid_pt = grid_twips / 20.0
    
    style_normal = doc.Styles("Normal")
    style_normal.Font.Name = "Calibri"
    style_normal.Font.Size = 11
    
    # Build: Calibri 11 -> Meiryo 20 -> Calibri 11
    rng = doc.Range(0, 0)
    rng.Text = "Default Calibri 11pt\r"
    rng.Font.Name = "Calibri"
    rng.Font.Size = 11
    
    rng2 = doc.Range(rng.End, rng.End)
    rng2.Text = "Meiryo 20pt test\r"
    rng2.Font.Name = "Meiryo"
    rng2.Font.Size = 20
    
    rng3 = doc.Range(rng2.End, rng2.End)
    rng3.Text = "Back to Calibri 11pt\r"
    rng3.Font.Name = "Calibri"
    rng3.Font.Size = 11
    
    doc.Repaginate()
    time.sleep(0.5)
    
    print(f"\nGrid: {grid_lines} lines/page = {grid_pt:.2f}pt pitch")
    
    positions = []
    for i in range(1, min(doc.Paragraphs.Count + 1, 5)):
        para = doc.Paragraphs(i)
        r = para.Range
        top = r.Information(4)  # wdVerticalPositionRelativeToPage
        ls = para.LineSpacing
        fn = r.Font.Name if r.Font.Name else "?"
        fs = r.Font.Size if r.Font.Size else 0
        positions.append((fn, fs, top, ls))
        print(f"  P{i}: {fn} {fs:.0f}pt top={top:.2f} ls={ls:.2f}")
    
    if len(positions) >= 2:
        gap = positions[1][2] - positions[0][2]
        default_ls = positions[0][3]
        
        # snap(default)
        snap_def = round((default_ls + grid_pt/2) / grid_pt) * grid_pt
        
        # Meiryo 20pt natural height (GDI): (2171+901)/2048*20*15/20 ≈ 22.49pt
        # Wait, let me use actual winAscent/winDescent from metrics
        # Meiryo: UPM=2048, winA=2171, winD=901
        # GDI tmHeight = ceil((2171+901)/2048 * 20) = ceil(30.0) = 30
        # tmExternalLeading = max(0, ceil((2171+901+76)/2048*20) - ceil((2171+901)/2048*20)) = max(0, ceil(30.74)-30) = max(0,31-30) = 1
        # gdi_height = (30+1)*15/20 = 23.25pt ... hmm
        # Actually with lineGap=0 fix: lineGap=0 → artificial gap = 76/256*UPM  
        # No wait, Meiryo has lineGap=0? Let me check
        # Meiryo: hhea lineGap might not be 0
        # Let me just measure from Word directly
        
        meiryo_ls = positions[1][3] if len(positions) >= 2 else 0
        
        # Formula A: snap(max(meiryo_natural, default_natural))
        # The line spacing Word reports should be the snapped value
        formula_a = round((max(meiryo_ls, default_ls) + grid_pt/2) / grid_pt) * grid_pt
        # But if meiryo_ls is already the final value, we can't separate snap from max
        
        # Better: use gap between P1 top and P2 top = effective line height of P1
        # P1 is Calibri 11 but it determines the spacing to P2
        # Actually in Word, each paragraph's line height is its own - the gap IS line 1's height
        print(f"  Gap P1→P2 = {gap:.2f}pt")
        print(f"  snap(default_ls={default_ls:.2f}) = {snap_def:.2f}")
        print(f"  Meiryo ls = {meiryo_ls:.2f}")
    
    doc.Close(0)

word.Quit()
print("\nDone.")
