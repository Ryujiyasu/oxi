"""
Ra: 番号リスト(numbering)のindent・タブ・番号位置をCOM計測で確定
- numFmt別の番号テキスト位置
- hanging indent の関係
- タブ位置 (番号とテキスト間)
- ネストレベル別のインデント
- 番号リスト + カスタムタブの相互作用
"""
import win32com.client, json, os

word = win32com.client.Dispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = False

results = []


def test_basic_numbered_list():
    """Basic numbered list (1. 2. 3.) - measure positions."""
    wdoc = word.Documents.Add()
    try:
        sec = wdoc.Sections(1)
        sec.PageSetup.LeftMargin = 72
        sec.PageSetup.RightMargin = 72

        wdoc.Content.Text = ""

        # Add 5 paragraphs with numbered list
        for i in range(5):
            if i > 0:
                rng = wdoc.Range(wdoc.Content.End - 1, wdoc.Content.End - 1)
                rng.InsertParagraphAfter()
            para = wdoc.Paragraphs(wdoc.Paragraphs.Count)
            para.Range.Text = f"List item number {i+1} text content here."
            para.Range.Font.Name = "Calibri"
            para.Range.Font.Size = 11

        # Apply numbered list via COM
        rng = wdoc.Range(wdoc.Paragraphs(1).Range.Start, wdoc.Paragraphs(5).Range.End)
        rng.ListFormat.ApplyNumberDefault()

        wdoc.Repaginate()

        data = {"scenario": "basic_numbered", "paragraphs": []}
        for i in range(1, 6):
            para = wdoc.Paragraphs(i)
            prng = para.Range
            fmt = para.Format

            # Get list info
            list_str = prng.ListFormat.ListString
            list_level = prng.ListFormat.ListLevelNumber

            pd = {
                "index": i,
                "x_pt": round(prng.Information(5), 4),
                "y_pt": round(prng.Information(6), 4),
                "left_indent": round(fmt.LeftIndent, 4),
                "first_line_indent": round(fmt.FirstLineIndent, 4),
                "list_string": list_str,
                "list_level": list_level,
                "text": prng.Text.strip()[:40],
            }

            # Measure character positions
            segments = []
            cur = {"start_x": None, "chars": ""}
            for ci in range(prng.Start, min(prng.End, prng.Start + 80)):
                cr = wdoc.Range(ci, ci + 1)
                ch = cr.Text
                x = cr.Information(5)
                if ord(ch) == 9:
                    if cur["start_x"] is not None:
                        segments.append(cur)
                    cur = {"start_x": None, "chars": ""}
                elif ord(ch) in (13, 7):
                    pass
                else:
                    if cur["start_x"] is None:
                        cur["start_x"] = round(x, 4)
                    cur["chars"] += ch
            if cur["start_x"] is not None:
                segments.append(cur)
            pd["segments"] = segments

            data["paragraphs"].append(pd)

        return data
    finally:
        wdoc.Close(False)


def test_bullet_list():
    """Bullet list - measure positions."""
    wdoc = word.Documents.Add()
    try:
        sec = wdoc.Sections(1)
        sec.PageSetup.LeftMargin = 72

        wdoc.Content.Text = ""
        for i in range(5):
            if i > 0:
                rng = wdoc.Range(wdoc.Content.End - 1, wdoc.Content.End - 1)
                rng.InsertParagraphAfter()
            para = wdoc.Paragraphs(wdoc.Paragraphs.Count)
            para.Range.Text = f"Bullet item {i+1} text."
            para.Range.Font.Name = "Calibri"
            para.Range.Font.Size = 11

        rng = wdoc.Range(wdoc.Paragraphs(1).Range.Start, wdoc.Paragraphs(5).Range.End)
        rng.ListFormat.ApplyBulletDefault()

        wdoc.Repaginate()

        data = {"scenario": "bullet_list", "paragraphs": []}
        for i in range(1, 6):
            para = wdoc.Paragraphs(i)
            prng = para.Range
            fmt = para.Format
            data["paragraphs"].append({
                "index": i,
                "x_pt": round(prng.Information(5), 4),
                "left_indent": round(fmt.LeftIndent, 4),
                "first_line_indent": round(fmt.FirstLineIndent, 4),
                "list_string": prng.ListFormat.ListString,
                "text": prng.Text.strip()[:40],
            })

        return data
    finally:
        wdoc.Close(False)


def test_nested_list():
    """Multi-level nested list - measure indent per level."""
    wdoc = word.Documents.Add()
    try:
        sec = wdoc.Sections(1)
        sec.PageSetup.LeftMargin = 72

        wdoc.Content.Text = ""
        # Create 9 items: 3 at level 1, 3 at level 2, 3 at level 3
        levels = [1, 1, 1, 2, 2, 2, 3, 3, 3]
        for i in range(9):
            if i > 0:
                rng = wdoc.Range(wdoc.Content.End - 1, wdoc.Content.End - 1)
                rng.InsertParagraphAfter()
            para = wdoc.Paragraphs(wdoc.Paragraphs.Count)
            para.Range.Text = f"Level {levels[i]} item text content."
            para.Range.Font.Name = "Calibri"
            para.Range.Font.Size = 11

        # Apply numbered list
        rng = wdoc.Range(wdoc.Paragraphs(1).Range.Start, wdoc.Paragraphs(9).Range.End)
        rng.ListFormat.ApplyNumberDefault()

        # Change levels by using ListIndent (Tab key equivalent)
        for i in range(4, 7):
            wdoc.Paragraphs(i).Range.Select()
            wdoc.Application.Selection.Range.ListFormat.ListIndent()
        for i in range(7, 10):
            wdoc.Paragraphs(i).Range.Select()
            wdoc.Application.Selection.Range.ListFormat.ListIndent()
            wdoc.Application.Selection.Range.ListFormat.ListIndent()

        wdoc.Repaginate()

        data = {"scenario": "nested_list", "paragraphs": []}
        for i in range(1, 10):
            para = wdoc.Paragraphs(i)
            prng = para.Range
            fmt = para.Format
            pd = {
                "index": i,
                "x_pt": round(prng.Information(5), 4),
                "left_indent": round(fmt.LeftIndent, 4),
                "first_line_indent": round(fmt.FirstLineIndent, 4),
                "list_string": prng.ListFormat.ListString,
                "list_level": prng.ListFormat.ListLevelNumber,
            }

            # Char positions for first item of each level
            if i in (1, 4, 7):
                segments = []
                cur = {"start_x": None, "chars": ""}
                for ci in range(prng.Start, min(prng.End, prng.Start + 60)):
                    cr = wdoc.Range(ci, ci + 1)
                    ch = cr.Text
                    x = cr.Information(5)
                    if ord(ch) == 9:
                        if cur["start_x"] is not None:
                            segments.append(cur)
                        cur = {"start_x": None, "chars": ""}
                    elif ord(ch) in (13, 7):
                        pass
                    else:
                        if cur["start_x"] is None:
                            cur["start_x"] = round(x, 4)
                        cur["chars"] += ch
                if cur["start_x"] is not None:
                    segments.append(cur)
                pd["segments"] = segments

            data["paragraphs"].append(pd)

        return data
    finally:
        wdoc.Close(False)


def test_list_tab_interaction():
    """List with custom tab stops - how do they interact?"""
    wdoc = word.Documents.Add()
    try:
        sec = wdoc.Sections(1)
        sec.PageSetup.LeftMargin = 72

        wdoc.Content.Text = ""
        for i in range(3):
            if i > 0:
                rng = wdoc.Range(wdoc.Content.End - 1, wdoc.Content.End - 1)
                rng.InsertParagraphAfter()
            para = wdoc.Paragraphs(wdoc.Paragraphs.Count)
            para.Range.Text = f"Item {i+1}\tTabbed text here."
            para.Range.Font.Name = "Calibri"
            para.Range.Font.Size = 11
            # Add custom tab at 144pt
            para.Format.TabStops.Add(144, 0, 0)

        rng = wdoc.Range(wdoc.Paragraphs(1).Range.Start, wdoc.Paragraphs(3).Range.End)
        rng.ListFormat.ApplyNumberDefault()

        wdoc.Repaginate()

        data = {"scenario": "list_tab_interaction", "paragraphs": []}
        for i in range(1, 4):
            para = wdoc.Paragraphs(i)
            prng = para.Range

            segments = []
            cur = {"start_x": None, "chars": ""}
            for ci in range(prng.Start, min(prng.End, prng.Start + 80)):
                cr = wdoc.Range(ci, ci + 1)
                ch = cr.Text
                x = cr.Information(5)
                if ord(ch) == 9:
                    if cur["start_x"] is not None:
                        segments.append(cur)
                    cur = {"start_x": None, "chars": ""}
                elif ord(ch) in (13, 7):
                    pass
                else:
                    if cur["start_x"] is None:
                        cur["start_x"] = round(x, 4)
                    cur["chars"] += ch
            if cur["start_x"] is not None:
                segments.append(cur)

            data["paragraphs"].append({
                "index": i,
                "left_indent": round(para.Format.LeftIndent, 4),
                "first_line_indent": round(para.Format.FirstLineIndent, 4),
                "list_string": prng.ListFormat.ListString,
                "segments": segments,
            })

        return data
    finally:
        wdoc.Close(False)


try:
    d1 = test_basic_numbered_list()
    results.append(d1)
    ml = 72.0
    print("=== basic_numbered ===")
    for p in d1["paragraphs"]:
        segs = [(round(s["start_x"] - ml, 2), s["chars"][:15]) for s in p.get("segments", [])]
        print(f"  P{p['index']}: li={p['left_indent']}, fli={p['first_line_indent']}, "
              f"list=\"{p['list_string']}\", segs={segs}")

    d2 = test_bullet_list()
    results.append(d2)
    print(f"\n=== bullet_list ===")
    for p in d2["paragraphs"]:
        ls = repr(p['list_string'])
        print(f"  P{p['index']}: li={p['left_indent']}, fli={p['first_line_indent']}, "
              f"list={ls}")

    d3 = test_nested_list()
    results.append(d3)
    print(f"\n=== nested_list ===")
    for p in d3["paragraphs"]:
        segs = [(round(s["start_x"] - ml, 2), s["chars"][:15]) for s in p.get("segments", [])]
        segs_str = f", segs={segs}" if segs else ""
        ls = repr(p['list_string'])
        print(f"  P{p['index']}: level={p['list_level']}, li={p['left_indent']}, fli={p['first_line_indent']}, "
              f"list={ls}{segs_str}")

    d4 = test_list_tab_interaction()
    results.append(d4)
    print(f"\n=== list_tab_interaction ===")
    for p in d4["paragraphs"]:
        segs = [(round(s["start_x"] - ml, 2), s["chars"][:15]) for s in p.get("segments", [])]
        print(f"  P{p['index']}: li={p['left_indent']}, fli={p['first_line_indent']}, "
              f"list=\"{p['list_string']}\", segs={segs}")

finally:
    word.Quit()

out_path = os.path.join(os.path.dirname(__file__), 'output', 'ra_numbering.json')
os.makedirs(os.path.dirname(out_path), exist_ok=True)
with open(out_path, 'w', encoding='utf-8') as f:
    json.dump(results, f, indent=2, ensure_ascii=False)
print(f"\nResults saved to {out_path}")
