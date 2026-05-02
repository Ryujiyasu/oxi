"""Debug: enumerate all chars in cell probe at slack=3 to find the 2 extra chars."""
import sys, time, subprocess
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
subprocess.run(['taskkill','/F','/IM','WINWORD.EXE'], capture_output=True)
time.sleep(2)

word = w32.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False
path = r"C:\Users\ryuji\oxi-1\tools\metrics\mech2_cell_repro\cell_sl3.0.docx"
d = word.Documents.Open(path, ReadOnly=True)
chars = d.Range().Characters
print(f"Total characters: {chars.Count}")
for ci in range(1, chars.Count + 1):
    try:
        c = chars(ci)
        t = c.Text
        x = float(c.Information(5))
        y = float(c.Information(6))
        print(f"  [{ci}] ord={ord(t[0]) if t else '?'} repr={t!r} x={x} y={y}")
    except Exception as e:
        print(f"  [{ci}] error: {e}")
d.Close(SaveChanges=False)
word.Quit()
