# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.
"""Show where Oxi's position count for a paragraph parts from Word's.

    python tools/metrics/edit_offset_diag.py <docx> <para> [<para>...]

`para` is the index as edit_offset_census.py reports it. For each, prints
Word's characters with field codes and hidden text shown (one per
position), Oxi's (`edit --text`), and the first position where they differ.
"""
import json
import os
import subprocess
import sys

import pythoncom
import win32com.client

sys.path.insert(0, os.path.dirname(__file__))
import edit_offset_census as census  # noqa: E402


def main():
    path = os.path.abspath(sys.argv[1])
    wanted = [int(a) for a in sys.argv[2:]]
    texts, error = census.oxi_texts(path)
    if texts is None:
        sys.exit(error)
    pythoncom.CoInitialize()
    w = win32com.client.DispatchEx("Word.Application")
    w.Visible = False
    w.DisplayAlerts = 0
    try:
        kept = census.word_paragraph_ranges(w, path)
        d = w.Documents.Open(path, ReadOnly=True, AddToRecentFiles=False, Visible=False)
        try:
            for n in wanted:
                start, end = kept[n]
                rg = d.Range(start, end - 1)
                rg.TextRetrievalMode.IncludeFieldCodes = True
                rg.TextRetrievalMode.IncludeHiddenText = True
                # One character per position: walk Characters (Range.Text
                # still drops some marks).
                word = "".join(d.Range(i, i + 1).Text or "?" for i in range(start, end - 1)) if end - start < 3000 else rg.Text
                oxi = texts[n]
                # Word gives field marks and objects back as None ("?") and
                # sometimes upper-cases a field code.
                def same(a, b):
                    return a == b or a.lower() == b.lower() or (a == "?" and (ord(b) < 32 or b in "\u0001\u0002"))
                k = next((i for i, (a, b) in enumerate(zip(word, oxi)) if not same(a, b)), min(len(word), len(oxi)))
                print(f"para {n}: word {end - start - 1} positions, oxi {census.units(oxi)}; first difference at {k}")
                print("  word:", json.dumps(word[max(0, k - 20):k + 40], ensure_ascii=False))
                print("  oxi: ", json.dumps(oxi[max(0, k - 20):k + 40], ensure_ascii=False))
        finally:
            d.Close(0)
    finally:
        w.Quit()


if __name__ == "__main__":
    main()
