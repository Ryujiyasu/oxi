"""Re-measure kinsoku_mech2_repro/ with per-variant Word restart."""
import json, os, sys, time, subprocess
from pathlib import Path
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPRO = Path(r"C:\Users\ryuji\oxi-1\tools\metrics\kinsoku_mech2_repro")
RESULT = Path(r"C:\Users\ryuji\oxi-1\pipeline_data\kinsoku_mech2.json")
YAKUMONO_B = set("」")


def kill_word():
    subprocess.run(['taskkill','/F','/IM','WINWORD.EXE'], capture_output=True)
    time.sleep(3)


def measure(path):
    word = w32.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    try:
        d = word.Documents.Open(str(path), ReadOnly=True)
        time.sleep(0.3)
        try:
            chars = d.Range().Characters
            xs = []
            for ci in range(1, chars.Count + 1):
                try:
                    c = chars(ci)
                    t = c.Text
                    if t in ("\r","\x07"): continue
                    xs.append((t, float(c.Information(5)), float(c.Information(6)),
                               float(c.Font.Size if c.Font.Size else 0)))
                except: continue
        finally:
            try: d.Close(SaveChanges=False)
            except: pass
        if not xs: return {"error": "no chars"}
        lines_b = {}
        for t, x, y, sz in xs:
            ykey = round(y, 0)
            lines_b.setdefault(ykey, []).append((t, x, y, sz))
        line_data = []
        for ykey in sorted(lines_b.keys()):
            items = sorted(lines_b[ykey], key=lambda v: v[1])
            chars_text = "".join(it[0] for it in items)
            advs = []
            for i in range(len(items) - 1):
                advs.append((items[i][0], round(items[i+1][1] - items[i][1], 3),
                             items[i][3]))
            yak_in_line = None
            for i, (ch, adv, sz) in enumerate(advs):
                if ch == "」":
                    yak_in_line = {"line_pos": i+1, "adv": adv, "sz": sz}
                    break
            # If yak is last char in line (no advance), check items
            if not yak_in_line:
                for i, item in enumerate(items):
                    ch, _, _, sz = item
                    if ch == "」":
                        yak_in_line = {"line_pos": i+1, "adv": None, "sz": sz}
                        break
            line_width = (items[-1][1] - items[0][1]) + items[-1][3] if items else 0
            line_data.append({
                "y": ykey,
                "n_chars": len(items),
                "text_summary": chars_text[:50],
                "line_width": round(line_width, 2),
                "yak_in_line": yak_in_line,
            })
        return {"n_lines": len(line_data), "lines": line_data}
    finally:
        try: word.Quit()
        except: pass


def main():
    docs = sorted(REPRO.glob("K_*.docx"))
    out = {}
    for d in docs:
        name = d.stem
        cw = int(name.replace("K_cw", ""))
        kill_word()
        try:
            r = measure(d)
        except Exception as e:
            r = {"measure_error": str(e)}
        out[name] = {"content_w": cw, "natural": 600.0, **r}
        print(f"\n[{name}] cw={cw}")
        if "lines" in r:
            for li, ln in enumerate(r["lines"], start=1):
                yak_str = ""
                if ln.get("yak_in_line"):
                    ya = ln["yak_in_line"]
                    adv_str = f"{ya['adv']:.2f}" if ya.get('adv') else "(line-end)"
                    yak_str = f"  [」 at L{li} pos {ya['line_pos']} adv={adv_str}]"
                print(f"  L{li}: n={ln['n_chars']:>3} w={ln['line_width']:>6.2f} (cw={cw}) diff={ln['line_width']-cw:+5.2f}{yak_str}")
        else:
            print(f"  ERR: {r.get('measure_error', 'unknown')[:100]}")
        with open(RESULT, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)


if __name__ == "__main__":
    main()
