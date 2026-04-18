"""Run a single docx test, kill Word first if needed.

Usage: python single_test_runner.py <docx_path>
"""
import os, sys, time, subprocess
import win32com.client

def kill_word():
    try:
        subprocess.run(['taskkill', '/F', '/IM', 'WINWORD.EXE'], capture_output=True, timeout=10)
        time.sleep(2)
    except Exception:
        pass

def measure(path):
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(path, ReadOnly=True); time.sleep(0.5)
        results = {}
        for p in doc.Paragraphs:
            text = p.Range.Text
            if '（' not in text: continue
            rng = p.Range
            for ci in range(1, rng.Characters.Count + 1):
                c = rng.Characters(ci)
                if c.Text == '（':
                    try:
                        x1 = c.Information(5); y1 = c.Information(6)
                        nxt = rng.Characters(ci + 1)
                        x2 = nxt.Information(5); y2 = nxt.Information(6)
                        if abs(y1 - y2) > 2: continue
                        fs = round(c.Font.Size, 1)
                        if fs not in results:
                            results[fs] = round(x2 - x1, 2)
                    except: pass
            if len(results) >= 3: break
        doc.Close(False)
        return results
    finally:
        try: word.Quit()
        except: pass

if __name__ == '__main__':
    kill_word()
    path = sys.argv[1]
    r = measure(path)
    print(r)
