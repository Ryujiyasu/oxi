# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.
"""Does an oxidocs-model position edit do what the same action does in Word?

A base document is made in Word. Each scenario is done twice from it: in
Word, through Selection (what the keyboard does), and by oxidocs-model's
`edit` example with the equivalent edit. Word then reads both results the
same way (per paragraph: text, alignment, space after, and per character:
bold, italic, inside a hyperlink) and the two readings are compared. Word
opening the Oxi file at all is part of the check.

    python tools/metrics/edit_word_diff.py [--keep DIR]

Positions in the edits are Word character offsets within a paragraph, so a
scenario names the same place on both sides.
"""
import json
import os
import subprocess
import sys
import tempfile

import pythoncom
import win32com.client

EDIT = os.path.join(os.path.dirname(__file__), "..", "..", "target", "release", "examples", "edit.exe")


def make_base(w, path):
    d = w.Documents.Add()
    d.Range().Text = "AAAA\rBBBB\rCC link DD\rEEEE"
    P = lambda i: d.Paragraphs(i)
    P(1).Alignment = 1
    P(1).SpaceAfter = 30
    P(2).Alignment = 2
    P(2).SpaceAfter = 5
    d.Range(P(1).Range.Start, P(1).Range.Start + 2).Bold = True
    s = P(3).Range.Start
    d.Hyperlinks.Add(d.Range(s + 3, s + 7), "http://example.com/")
    d.Range(P(4).Range.Start + 1, P(4).Range.Start + 3).Italic = True
    d.SaveAs2(path, 16)
    d.Close(0)


def read(w, path):
    d = w.Documents.Open(path, ReadOnly=True, AddToRecentFiles=False)
    try:
        out = []
        for p in d.Paragraphs:
            chars = []
            for c in p.Range.Characters:
                if c.Text == "\r":
                    continue
                flags = ("B" if c.Bold else "") + ("I" if c.Italic else "") + ("L" if c.Hyperlinks.Count else "")
                chars.append(f"{c.Text}{flags}")
            out.append({"text": p.Range.Text.rstrip("\r"), "align": p.Alignment, "after": p.SpaceAfter, "chars": " ".join(chars)})
        return out
    finally:
        d.Close(0)


K = {}


def pos(para, offset):
    return {"para": para, "offset": offset}


# Each scenario: name, Word action (document, absolute-position helper), Oxi edits.
# at(i, k) is paragraph i's (0-based) start plus k.
SCENARIOS = [
    ("type inside bold run", lambda d, at, s: (s.SetRange(at(0, 1), at(0, 1)), s.TypeText("x")),
     [{"op": "insert_text", "at": pos(0, 1), "text": "x"}]),
    ("type at bold|plain edge", lambda d, at, s: (s.SetRange(at(0, 2), at(0, 2)), s.TypeText("x")),
     [{"op": "insert_text", "at": pos(0, 2), "text": "x"}]),
    ("type at paragraph start", lambda d, at, s: (s.SetRange(at(1, 0), at(1, 0)), s.TypeText("x")),
     [{"op": "insert_text", "at": pos(1, 0), "text": "x"}]),
    ("type at italic start", lambda d, at, s: (s.SetRange(at(3, 1), at(3, 1)), s.TypeText("x")),
     [{"op": "insert_text", "at": pos(3, 1), "text": "x"}]),
    ("delete within paragraph", lambda d, at, s: (s.SetRange(at(0, 1), at(0, 3)), s.Delete()),
     [{"op": "delete", "from": pos(0, 1), "to": pos(0, 3)}]),
    ("delete mid..mid across", lambda d, at, s: (s.SetRange(at(0, 2), at(1, 2)), s.Delete()),
     [{"op": "delete", "from": pos(0, 2), "to": pos(1, 2)}]),
    ("delete start..mid across", lambda d, at, s: (s.SetRange(at(0, 0), at(1, 2)), s.Delete()),
     [{"op": "delete", "from": pos(0, 0), "to": pos(1, 2)}]),
    ("delete just the mark", lambda d, at, s: (s.SetRange(at(0, 4), at(1, 0)), s.Delete()),
     [{"op": "delete", "from": pos(0, 4), "to": pos(1, 0)}]),
    ("delete over three paragraphs", lambda d, at, s: (s.SetRange(at(0, 2), at(2, 1)), s.Delete()),
     [{"op": "delete", "from": pos(0, 2), "to": pos(2, 1)}]),
    ("Delete key at paragraph end", lambda d, at, s: (s.SetRange(at(0, 4), at(0, 4)), s.Delete()),
     [{"op": "join_paragraphs", "para": 0}]),
    ("Backspace at paragraph start", lambda d, at, s: (s.SetRange(at(1, 0), at(1, 0)), s.TypeBackspace()),
     [{"op": "join_paragraphs", "para": 0}]),
    ("type over across", lambda d, at, s: (s.SetRange(at(0, 2), at(1, 2)), s.TypeText("X")),
     [{"op": "replace", "from": pos(0, 2), "to": pos(1, 2), "text": "X"}]),
    ("type over within, from bold", lambda d, at, s: (s.SetRange(at(0, 1), at(0, 3)), s.TypeText("xy")),
     [{"op": "replace", "from": pos(0, 1), "to": pos(0, 3), "text": "xy"}]),
    ("type over within, from plain", lambda d, at, s: (s.SetRange(at(3, 0), at(3, 2)), s.TypeText("xy")),
     [{"op": "replace", "from": pos(3, 0), "to": pos(3, 2), "text": "xy"}]),
    # K is the offset just after the "l" of the link text, taken from Oxi's
    # listing of the base document: the same number is used as a Word
    # position, which is the claim under test.
    ("type inside link text", lambda d, at, s: (s.SetRange(at(2, K["link"]), at(2, K["link"])), s.TypeText("x")),
     lambda: [{"op": "insert_text", "at": pos(2, K["link"]), "text": "x"}]),
    ("type after link", lambda d, at, s: (s.SetRange(at(2, K["after_link"]), at(2, K["after_link"])), s.TypeText("x")),
     lambda: [{"op": "insert_text", "at": pos(2, K["after_link"]), "text": "x"}]),
    ("delete across link edge", lambda d, at, s: (s.SetRange(at(2, 1), at(2, K["link"])), s.Delete()),
     lambda: [{"op": "delete", "from": pos(2, 1), "to": pos(2, K["link"])}]),
    ("delete from link text past its end", lambda d, at, s: (s.SetRange(at(2, K["link"]), at(2, K["after_link"] + 2)), s.Delete()),
     lambda: [{"op": "delete", "from": pos(2, K["link"]), "to": pos(2, K["after_link"] + 2)}]),
    ("delete inside link text", lambda d, at, s: (s.SetRange(at(2, K["link"]), at(2, K["link"] + 2)), s.Delete()),
     lambda: [{"op": "delete", "from": pos(2, K["link"]), "to": pos(2, K["link"] + 2)}]),
    ("Enter mid bold", lambda d, at, s: (s.SetRange(at(0, 1), at(0, 1)), s.TypeParagraph()),
     [{"op": "split_paragraph", "at": pos(0, 1)}]),
    ("Enter at end", lambda d, at, s: (s.SetRange(at(1, 4), at(1, 4)), s.TypeParagraph()),
     [{"op": "split_paragraph", "at": pos(1, 4)}]),
    ("Enter at start", lambda d, at, s: (s.SetRange(at(1, 0), at(1, 0)), s.TypeParagraph()),
     [{"op": "split_paragraph", "at": pos(1, 0)}]),
    ("bold a range across runs", lambda d, at, s: (s.SetRange(at(0, 1), at(0, 3)), setattr(s.Font, "Bold", True)),
     [{"op": "format_text", "from": pos(0, 1), "to": pos(0, 3), "patch": {"set": [{"name": "w:b"}]}}]),
    ("unbold part of bold", lambda d, at, s: (s.SetRange(at(0, 0), at(0, 1)), setattr(s.Font, "Bold", False)),
     [{"op": "format_text", "from": pos(0, 0), "to": pos(0, 1), "patch": {"remove": ["w:b"]}}]),
    ("center two paragraphs", lambda d, at, s: (s.SetRange(at(1, 1), at(2, 1)), setattr(s.ParagraphFormat, "Alignment", 1)),
     [{"op": "format_paragraphs", "from": pos(1, 1), "to": pos(2, 1), "patch": {"set": [{"name": "w:jc", "attrs": [["w:val", "center"]]}]}}]),
]


def main():
    keep = None
    if "--keep" in sys.argv:
        keep = sys.argv[sys.argv.index("--keep") + 1]
        os.makedirs(keep, exist_ok=True)
    work = keep or tempfile.mkdtemp(prefix="edit_word_diff_")
    base = os.path.join(work, "base.docx")
    pythoncom.CoInitialize()
    w = win32com.client.DispatchEx("Word.Application")
    w.Visible = False
    w.DisplayAlerts = 0
    passed = 0
    try:
        make_base(w, base)
        texts = json.loads(subprocess.run([EDIT, "--text", base], capture_output=True, check=True).stdout)
        link_para = texts[2]
        K["link"] = link_para.index("\u0014link") + 2
        K["after_link"] = link_para.index("\u0015") + 1
        for i, (name, word_action, edits) in enumerate(SCENARIOS):
            word_out = os.path.join(work, f"{i:02}_word.docx")
            oxi_out = os.path.join(work, f"{i:02}_oxi.docx")
            d = w.Documents.Open(base, AddToRecentFiles=False)
            starts = [p.Range.Start for p in d.Paragraphs]
            w.Selection.SetRange(0, 0)
            word_action(d, lambda p, k: starts[p] + k, w.Selection)
            d.SaveAs2(word_out, 16)
            d.Close(0)
            if callable(edits):
                edits = edits()
            edits_path = os.path.join(work, f"{i:02}_edits.json")
            with open(edits_path, "w", encoding="utf-8") as f:
                json.dump(edits, f)
            r = subprocess.run([EDIT, base, edits_path, oxi_out], capture_output=True)
            if r.returncode != 0:
                print(f"FAIL {name}: oxi refused: {r.stderr.decode('utf-8', 'replace').strip()}")
                continue
            try:
                got = read(w, oxi_out)
            except Exception as e:
                print(f"FAIL {name}: Word cannot open the Oxi file: {e}")
                continue
            want = read(w, word_out)
            if got == want:
                passed += 1
                print(f"ok   {name}")
            else:
                print(f"FAIL {name}")
                for a, b in zip(want + [None] * (len(got) - len(want)), got + [None] * (len(want) - len(got))):
                    if a != b:
                        print(f"     word: {a}\n     oxi:  {b}")
    finally:
        w.Quit()
    print(f"{passed}/{len(SCENARIOS)} match Word")


if __name__ == "__main__":
    main()
