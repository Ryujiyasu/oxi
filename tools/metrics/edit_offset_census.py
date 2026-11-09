# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.
"""Do edit positions count characters the way Word does?

For each document, list the body paragraphs as oxidocs-model's edit module
numbers them (`edit --text`) and as Word does (Document.Paragraphs, skipping
table end-of-row marks), and compare each paragraph's length with Word's
character count for it (Range.End - Range.Start, less the paragraph mark).

    python tools/metrics/edit_offset_census.py <docx or dir>... [--limit N] [--out report.json]

A document passes when the paragraph counts agree and every length agrees.
"""
import json
import os
import subprocess
import sys
from collections import Counter

import pythoncom
import win32com.client

EDIT = os.path.join(os.path.dirname(__file__), "..", "..", "target", "release", "examples", "edit.exe")


def docs(paths):
    for p in paths:
        if os.path.isdir(p):
            for root, _, files in os.walk(p):
                for f in sorted(files):
                    if f.endswith(".docx") and not f.startswith("~$"):
                        yield os.path.join(root, f)
        else:
            yield p


def units(s):
    """Word positions are UTF-16 code units."""
    return len(s.encode("utf-16-le")) // 2


def oxi_texts(path):
    r = subprocess.run([EDIT, "--text", path], capture_output=True, timeout=120)
    if r.returncode != 0:
        return None, r.stderr.decode("utf-8", "replace").strip()[-200:]
    return json.loads(r.stdout.decode("utf-8")), None


def _paragraphs(d):
    """(start, end, text) of each body paragraph as the model numbers them:
    Word's paragraphs without table end-of-row marks."""
    rows = []
    for p in d.Paragraphs:
        rg = p.Range
        text = rg.Text
        # Information(12) (in a table) is False for some table paragraphs;
        # only table paragraphs end in a cell mark, and the row number
        # (Information(13)) is -1 outside tables.
        in_table = text.endswith("\x07") or bool(rg.Information(12))
        row = rg.Information(13)
        no_cell = False
        if text.endswith("\x07"):
            try:
                rg.Cells(1)
            except Exception:
                no_cell = True
        rows.append((rg.Start, rg.End, text[:60], in_table, row, no_cell))
    out = []
    for i, (start, end, text, in_table, row, no_cell) in enumerate(rows):
        # A row's end mark is a paragraph with no cell that ends its row. A
        # continued vertically merged cell has no cell either, but more of
        # its row follows. (Rows() refuses tables with vertically merged
        # cells, so the test does not use it.)
        if no_cell:
            nxt = rows[i + 1] if i + 1 < len(rows) else None
            if nxt is None or not nxt[3] or nxt[4] != row:
                continue
        out.extend(_split_hidden_marks(d, start, end, text))
    return out


def _split_hidden_marks(d, start, end, text):
    """Word lists a paragraph whose mark is hidden (w:vanish) together with
    the next one; the file, and the model, keep them apart. Split at every
    paragraph mark inside the range, read with field codes and hidden text
    so that one character is one position."""
    rg = d.Range(start, end)
    rg.TextRetrievalMode.IncludeFieldCodes = True
    rg.TextRetrievalMode.IncludeHiddenText = True
    full = rg.Text
    if len(full) != end - start or "\r" not in full[:-1]:
        return [(start, end, text)]
    pieces, s = [], start
    for k, c in enumerate(full[:-1]):
        # Equations also hold "\r"; only a hidden paragraph mark splits.
        if c == "\r" and d.Range(start + k, start + k + 1).Font.Hidden:
            pieces.append((s, start + k + 1, full[s - start:k + 1][:60]))
            s = start + k + 1
    pieces.append((s, end, full[s - start:][:60]))
    return pieces


def word_paragraph_ranges(w, path):
    d = w.Documents.Open(os.path.abspath(path), ReadOnly=True, AddToRecentFiles=False, Visible=False)
    try:
        return [(s, e) for s, e, _ in _paragraphs(d)]
    finally:
        d.Close(0)


def word_lengths(w, path):
    d = w.Documents.Open(os.path.abspath(path), ReadOnly=True, AddToRecentFiles=False, Visible=False)
    try:
        return [(e - s - 1, t) for s, e, t in _paragraphs(d)]
    finally:
        d.Close(0)


def main():
    args = sys.argv[1:]
    limit = None
    out_path = None
    if "--limit" in args:
        i = args.index("--limit")
        limit = int(args[i + 1])
        del args[i:i + 2]
    if "--out" in args:
        i = args.index("--out")
        out_path = args[i + 1]
        del args[i:i + 2]
    pythoncom.CoInitialize()
    w = win32com.client.DispatchEx("Word.Application")
    w.Visible = False
    w.DisplayAlerts = 0
    report = []
    reasons = Counter()
    try:
        for n, path in enumerate(docs(args)):
            if limit is not None and n >= limit:
                break
            texts, error = oxi_texts(path)
            if texts is None:
                report.append({"doc": path, "status": "oxi_error", "error": error})
                reasons["oxi_error"] += 1
                continue
            try:
                word = word_lengths(w, path)
            except Exception as e:  # Word refuses the file
                report.append({"doc": path, "status": "word_error", "error": str(e)[:200]})
                reasons["word_error"] += 1
                continue
            entry = {"doc": path, "paragraphs": len(texts), "word_paragraphs": len(word)}
            if len(texts) != len(word):
                entry["status"] = "count"
            else:
                bad = [
                    {"para": i, "oxi": units(t), "word": wl, "oxi_text": t[:80], "word_text": wt}
                    for i, (t, (wl, wt)) in enumerate(zip(texts, word))
                    if units(t) != wl
                ]
                entry["status"] = "length" if bad else "pass"
                if bad:
                    entry["first"] = bad[:3]
                    entry["mismatched"] = len(bad)
            reasons[entry["status"]] += 1
            report.append(entry)
            print(f"[{n}] {entry['status']:6} {os.path.basename(path)}", flush=True)
    finally:
        w.Quit()
    print(dict(reasons))
    if out_path:
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(report, f, ensure_ascii=False, indent=1)


if __name__ == "__main__":
    main()
