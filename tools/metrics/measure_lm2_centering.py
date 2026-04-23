"""Measure Word's first-line Y via COM for 6 LM2 centering repros.

For each docx:
- Para 1 first line y (Information(6))
- Para 2 y  (gap = para2.y - para1.y = line_height used by Word)
- page_top = 1418/20 = 70.9pt
- offset_from_margin = para1.y - page_top
- If offset ≈ (pitch - fs)/2: Oxi's centering formula is correct (A)
- If offset ≈ 1.1pt: Word places glyph top at line top, no centering (B)
- If offset is something else: investigate further (C)
"""
import win32com.client
import json
from pathlib import Path

REPRO_DIR = Path(__file__).parent / "lm2_centering_repro"
OUT = Path(__file__).parent.parent.parent / "pipeline_data" / "lm2_centering_measurements.json"

PAGE_TOP_TW = 1418
PAGE_TOP_PT = PAGE_TOP_TW / 20.0  # 70.9

REPRO_SPECS = [
    ("C1_fs10_5_pitch18", dict(fs=10.5, pitch=18.0, font="MS Mincho")),
    ("C2_fs12_pitch18",   dict(fs=12.0, pitch=18.0, font="MS Mincho")),
    ("C3_fs14_pitch18",   dict(fs=14.0, pitch=18.0, font="MS Mincho")),
    ("C4_fs10_5_pitch15", dict(fs=10.5, pitch=15.0, font="MS Mincho")),
    ("C5_fs10_5_nogrid",  dict(fs=10.5, pitch=None, font="MS Mincho")),
    ("C6_fs10_5_compat14",dict(fs=10.5, pitch=18.0, font="MS Mincho")),
    ("C7_mincho_fs10_5",  dict(fs=10.5, pitch=18.0, font="MS Mincho")),
    ("C8_gothic_fs10_5",  dict(fs=10.5, pitch=18.0, font="MS Gothic")),
    ("C9_meiryo_fs10_5",  dict(fs=10.5, pitch=18.0, font="Meiryo")),
    ("C10_gothic_fs12",   dict(fs=12.0, pitch=18.0, font="MS Gothic")),
    ("C11_meiryo_fs12",   dict(fs=12.0, pitch=18.0, font="Meiryo")),
]


def measure_one(word, docx_path):
    doc = word.Documents.Open(str(docx_path), ReadOnly=True)
    try:
        p1 = doc.Paragraphs(1)
        p2 = doc.Paragraphs(2)
        y1 = p1.Range.Information(6)
        y2 = p2.Range.Information(6)
        return dict(y1=round(y1, 3), y2=round(y2, 3), gap=round(y2 - y1, 3))
    finally:
        doc.Close(SaveChanges=False)


def main():
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    results = {}
    try:
        for name, spec in REPRO_SPECS:
            docx = REPRO_DIR / f"{name}.docx"
            m = measure_one(word, docx)
            offset = round(m["y1"] - PAGE_TOP_PT, 3)
            fs = spec["fs"]
            pitch = spec["pitch"]
            predicted_center = round((pitch - fs) / 2, 3) if pitch else None
            results[name] = {**m, **spec, "offset_from_margin": offset,
                             "predicted_center_offset": predicted_center}
            print(f"{name}: fs={fs} pitch={pitch}")
            print(f"  y1={m['y1']} y2={m['y2']} gap={m['gap']}")
            print(f"  offset_from_margin={offset}pt  predicted_center={(pitch-fs)/2 if pitch else 'N/A'}pt")
    finally:
        word.Quit()

    OUT.parent.mkdir(parents=True, exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f"\nSaved: {OUT}")


if __name__ == "__main__":
    main()
