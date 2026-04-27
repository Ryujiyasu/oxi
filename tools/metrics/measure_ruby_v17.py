"""V17 — pure no_ruby_LH per font × base size profiling.

For each fixture (one per font), measure dy of same-base-size paragraph
pairs. Compares to:
  - CLAUDE.md `base × 9/7` (≈ 1.286 × base)
  - usWinAsc + usWinDesc / upem × base (full TTF line height)
  - hheaAsc + |hheaDesc| + hheaLG / upem × base (typo full)

Writes pipeline_data/ruby_v17_no_ruby_lh.json.
"""
import json
import os
import struct
import sys
import time

import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOCX_DIR = os.path.abspath("pipeline_data/docx")
OUT_PATH = os.path.abspath("pipeline_data/ruby_v17_no_ruby_lh.json")

V17_FIXTURES = [
    ("MSMincho_control", "MS Mincho",     "C:/Windows/Fonts/msmincho.ttc",   0),
    ("YuMincho",         "Yu Mincho",     "C:/Windows/Fonts/YuMin.ttf",      0),
    ("YuGothic",         "Yu Gothic R",   "C:/Windows/Fonts/YuGothR.ttc",    0),
    ("YuGothicUI",       "Yu Gothic UI",  "C:/Windows/Fonts/YuGothR.ttc",    1),
    ("Meiryo",           "Meiryo Reg",    "C:/Windows/Fonts/meiryo.ttc",     0),
    ("MeiryoUI",         "Meiryo UI",     "C:/Windows/Fonts/meiryo.ttc",     2),
]

BASES_PT = [9.0, 10.5, 11.0, 12.0, 14.0]


def parse_ttf(path: str, idx: int) -> dict:
    data = open(path, "rb").read()
    if data[0:4] == b"ttcf":
        n = struct.unpack(">I", data[8:12])[0]
        offs = [struct.unpack(">I", data[12 + i*4:16 + i*4])[0] for i in range(n)]
        face_off = offs[idx]
    else:
        face_off = 0
    nt = struct.unpack(">H", data[face_off+4:face_off+6])[0]
    tab = {}
    for i in range(nt):
        rec = face_off + 12 + i*16
        ttag = data[rec:rec+4].decode("latin1")
        offset = struct.unpack(">I", data[rec+8:rec+12])[0]
        tab[ttag] = offset
    head_off = tab.get("head", 0)
    upem = struct.unpack(">H", data[head_off+18:head_off+20])[0] if head_off else 0
    out = {"upem": upem}
    os2_off = tab.get("OS/2", 0)
    if os2_off:
        out["sTypoAsc"]  = struct.unpack(">h", data[os2_off+68:os2_off+70])[0]
        out["sTypoDesc"] = struct.unpack(">h", data[os2_off+70:os2_off+72])[0]
        out["sTypoLG"]   = struct.unpack(">h", data[os2_off+72:os2_off+74])[0]
        out["usWinAsc"]  = struct.unpack(">H", data[os2_off+74:os2_off+76])[0]
        out["usWinDesc"] = struct.unpack(">H", data[os2_off+76:os2_off+78])[0]
    hhea_off = tab.get("hhea", 0)
    if hhea_off:
        out["hheaAsc"]  = struct.unpack(">h", data[hhea_off+4:hhea_off+6])[0]
        out["hheaDesc"] = struct.unpack(">h", data[hhea_off+6:hhea_off+8])[0]
        out["hheaLG"]   = struct.unpack(">h", data[hhea_off+8:hhea_off+10])[0]
    return out


def measure_doc(word_app, docx_path: str) -> list[dict]:
    abs_path = os.path.abspath(docx_path)
    doc = word_app.Documents.Open(abs_path, ReadOnly=True)
    time.sleep(0.4)
    paras = []
    n = doc.Paragraphs.Count
    for pi in range(1, n + 1):
        p = doc.Paragraphs(pi)
        rng = p.Range
        try:
            y = rng.Information(6)
        except Exception:
            y = None
        text = (rng.Text or "").replace("\r", "").replace("\x07", "")
        paras.append({"i": pi, "y_pt": y, "text": text[:60]})
    doc.Close(SaveChanges=False)
    return paras


def main() -> None:
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    out: dict = {
        "_meta": {
            "tool": "tools/metrics/measure_ruby_v17.py",
            "purpose": "pure no_ruby_LH per font × base size profiling",
            "bases_pt": BASES_PT,
        },
        "fixtures": {},
    }
    try:
        for suffix, label, ttf_path, ttf_idx in V17_FIXTURES:
            ttf = parse_ttf(ttf_path, ttf_idx)
            fname = f"RUBY_V17_{suffix}_no_ruby_LH"
            print(f"\n=== {fname} ({label}) ===")
            print(f"  TTF: upem={ttf['upem']} usWinAsc={ttf['usWinAsc']} usWinDesc={ttf['usWinDesc']} sTypoLG={ttf['sTypoLG']} hheaTotal={ttf['hheaAsc']+abs(ttf['hheaDesc'])+ttf['hheaLG']}")

            docx_path = os.path.join(DOCX_DIR, fname + ".docx")
            paras = measure_doc(word, docx_path)
            # paragraphs: 5 size groups × 2 paragraphs + 1 closer = 11 paragraphs
            # group i (0-indexed) = paragraphs at indices (2i, 2i+1) (0-indexed in py = 1-based 2i+1, 2i+2)
            cells = []
            for i, base_pt in enumerate(BASES_PT):
                p_a_idx = 2*i + 1  # 1-based
                p_b_idx = 2*i + 2
                if p_b_idx > len(paras):
                    continue
                p_a = paras[p_a_idx - 1]
                p_b = paras[p_b_idx - 1]
                if p_a["y_pt"] is None or p_b["y_pt"] is None:
                    continue
                dy = p_b["y_pt"] - p_a["y_pt"]
                # Predictions:
                cjk_97 = base_pt * 9.0 / 7.0
                win_total_pt = (ttf["usWinAsc"] + ttf["usWinDesc"]) / ttf["upem"] * base_pt
                hhea_total_pt = (ttf["hheaAsc"] + abs(ttf["hheaDesc"]) + ttf["hheaLG"]) / ttf["upem"] * base_pt
                typo_total_lg_pt = (ttf["sTypoAsc"] + abs(ttf["sTypoDesc"]) + ttf["sTypoLG"]) / ttf["upem"] * base_pt
                cells.append({
                    "base_pt": base_pt,
                    "dy_pt": round(dy, 3),
                    "cjk_9/7_pred": round(cjk_97, 3),
                    "win_total_pred": round(win_total_pt, 3),
                    "hhea_total_pred": round(hhea_total_pt, 3),
                    "typo_total_lg_pred": round(typo_total_lg_pt, 3),
                })
                print(f"  base={base_pt:>5}pt: dy={dy:.3f}  vs 9/7={cjk_97:.3f} winTot={win_total_pt:.3f} hheaTot={hhea_total_pt:.3f} typoTot+LG={typo_total_lg_pt:.3f}")

            out["fixtures"][fname] = {
                "label": label,
                "ttf": ttf,
                "cells": cells,
            }
    finally:
        word.Quit()

    os.makedirs(os.path.dirname(OUT_PATH), exist_ok=True)
    with open(OUT_PATH, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"\nWrote {OUT_PATH}")


if __name__ == "__main__":
    main()
