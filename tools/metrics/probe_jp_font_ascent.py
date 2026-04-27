"""Probe OS/2 sTypoAscender / head.unitsPerEm for installed Japanese fonts.

Confirms (or refutes) the §18.7 ruby-ascent generalization:
    base_ascent_pt = base_pt × OS/2.sTypoAscender / head.unitsPerEm

Usage:
    python tools/metrics/probe_jp_font_ascent.py

Outputs a table per font and writes JSON to pipeline_data/jp_font_ascent_metrics.json.
"""
import json
import os
import struct
import sys
from typing import Optional

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_PATH = os.path.abspath("pipeline_data/jp_font_ascent_metrics.json")


def _u16(b: bytes, off: int) -> int:
    return struct.unpack(">H", b[off:off+2])[0]


def _i16(b: bytes, off: int) -> int:
    return struct.unpack(">h", b[off:off+2])[0]


def _u32(b: bytes, off: int) -> int:
    return struct.unpack(">I", b[off:off+4])[0]


def parse_face(data: bytes, face_off: int) -> Optional[dict]:
    n_tables = _u16(data, face_off + 4)
    tables: dict[str, tuple[int, int]] = {}
    for i in range(n_tables):
        rec = face_off + 12 + i * 16
        ttag = data[rec:rec+4].decode("latin1", "replace")
        offset = _u32(data, rec + 8)
        length = _u32(data, rec + 12)
        tables[ttag] = (offset, length)

    metrics: dict = {}
    head_off, _ = tables.get("head", (0, 0))
    if head_off:
        metrics["unitsPerEm"] = _u16(data, head_off + 18)
    os2_off, os2_len = tables.get("OS/2", (0, 0))
    if os2_off:
        metrics["sTypoAscender"]  = _i16(data, os2_off + 68)
        metrics["sTypoDescender"] = _i16(data, os2_off + 70)
        metrics["sTypoLineGap"]   = _i16(data, os2_off + 72)
        metrics["usWinAscent"]    = _u16(data, os2_off + 74)
        metrics["usWinDescent"]   = _u16(data, os2_off + 76)
        if os2_len > 89:
            metrics["sxHeight"]   = _i16(data, os2_off + 86)
            metrics["sCapHeight"] = _i16(data, os2_off + 88)
    hhea_off, _ = tables.get("hhea", (0, 0))
    if hhea_off:
        metrics["hheaAscender"]  = _i16(data, hhea_off + 4)
        metrics["hheaDescender"] = _i16(data, hhea_off + 6)
        metrics["hheaLineGap"]   = _i16(data, hhea_off + 8)

    name_off, _ = tables.get("name", (0, 0))
    if name_off:
        n = _u16(data, name_off + 2)
        string_off = _u16(data, name_off + 4)
        for i in range(n):
            rec = name_off + 6 + i * 12
            platform = _u16(data, rec)
            language = _u16(data, rec + 4)
            nameID = _u16(data, rec + 6)
            length = _u16(data, rec + 8)
            offset = _u16(data, rec + 10)
            if nameID == 1 and platform == 3 and language == 0x0409:
                s_off = name_off + string_off + offset
                metrics["family"] = data[s_off:s_off+length].decode("utf-16-be", "replace")
                break
    return metrics


def parse_file(path: str) -> list[dict]:
    if not os.path.exists(path):
        return []
    data = open(path, "rb").read()
    if data[0:4] == b"ttcf":
        n_fonts = _u32(data, 8)
        offsets = [_u32(data, 12 + i*4) for i in range(n_fonts)]
    else:
        offsets = [0]
    out = []
    for idx, off in enumerate(offsets):
        m = parse_face(data, off)
        if m:
            m["_path"] = path
            m["_face_idx"] = idx
            out.append(m)
    return out


FILES = [
    "C:/Windows/Fonts/msmincho.ttc",
    "C:/Windows/Fonts/msgothic.ttc",
    "C:/Windows/Fonts/YuMin.ttf",
    "C:/Windows/Fonts/YuGothR.ttc",
    "C:/Windows/Fonts/YuGothM.ttc",
    "C:/Windows/Fonts/YuGothB.ttc",
    "C:/Windows/Fonts/BIZ-UDMinchoM.ttc",
    "C:/Windows/Fonts/BIZ-UDGothicR.ttc",
    "C:/Windows/Fonts/BIZ-UDGothicB.ttc",
    "C:/Windows/Fonts/meiryo.ttc",
    "C:/Windows/Fonts/meiryob.ttc",
    "C:/Windows/Fonts/UDDigiKyokashoN-B.ttc",
]


def main() -> None:
    all_faces: list[dict] = []
    for f in FILES:
        all_faces.extend(parse_file(f))

    # Round 11 update: usWinAscent is the CORRECT field for Word ruby ascent
    # (see ra_manual_measurements.json entry ruby_ascent_constant_uses_usWinAscent_*).
    # sTypoAscender shown for legacy reference; ratio_winAsc/upem is the ship constant.
    print(
        f'{"family":<28}{"upem":>5}{"sTypoAsc":>10}{"usWinAsc":>10}'
        f'{"winRatio":>10}{"asc(10.5pt)":>13}{"asc(14pt)":>11}'
    )
    for m in all_faces:
        family = m.get("family", os.path.basename(m["_path"]))
        upem = m.get("unitsPerEm", 0)
        sta = m.get("sTypoAscender", 0)
        win = m.get("usWinAscent", 0)
        ratio = win / upem if upem else 0
        asc105 = ratio * 10.5
        asc14 = ratio * 14.0
        print(f'{family:<28}{upem:>5}{sta:>10}{win:>10}{ratio:>10.4f}{asc105:>13.3f}{asc14:>11.3f}')

    os.makedirs(os.path.dirname(OUT_PATH), exist_ok=True)
    with open(OUT_PATH, "w", encoding="utf-8") as f:
        json.dump({"faces": all_faces}, f, ensure_ascii=False, indent=2)
    print(f"\nWrote {OUT_PATH}")


if __name__ == "__main__":
    main()
