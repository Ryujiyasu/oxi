"""Measure Word's actual rendered width of 人的管理措置 at MS Mincho 10.5pt."""
import sys, os, win32com.client
sys.stdout.reconfigure(encoding='utf-8')

DOC = r"C:\Users\ryuji\oxi-main\tools\golden-test\documents\docx\b35123fe8efc_tokumei_08_01.docx"
word = win32com.client.Dispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = False
wdoc = word.Documents.Open(os.path.abspath(DOC), ReadOnly=True)
try:
    wdoc.Repaginate()
    tbl = wdoc.Tables(1)
    cell = tbl.Cell(6, 1)  # 人的管理措置
    rng = cell.Range
    # Get each char's x position using Information(2) - wdHorizontalPositionRelativeToPage
    print(f"Col 1 cell width = {cell.Width:.3f}pt")
    text = rng.Text.strip()
    print(f"Cell text: {text!r}")
    xs = []
    for i in range(1, rng.Characters.Count + 1):
        try:
            ch = rng.Characters(i)
            x = ch.Information(2)
            y = ch.Information(6)
            char = ch.Text
            xs.append((i, char, x, y))
        except Exception as e:
            break
    for i, c, x, y in xs:
        print(f"  {i}: {c!r} x={x} y={y}")
    if len(xs) >= 2:
        advances = [xs[i+1][2] - xs[i][2] for i in range(len(xs)-1) if xs[i+1][3] == xs[i][3]]
        print(f"Advances: {advances}")
        if advances:
            print(f"Mean advance: {sum(advances)/len(advances):.3f}pt")
finally:
    wdoc.Close(False)
    word.Quit()
