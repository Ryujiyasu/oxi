# -*- coding: utf-8 -*-
"""Derive Word's per-line oikomi/oidashi DECISION on nedocontract.

The s475 flat-cap break (and the even-distribution variant, both falsified on the
gate 2026-06-23) wrap word_i=337's spill root via an upstream over-wrap, while a
high cap over-fits word_i {400,434,465} (Word WRAPS them = oidashi). The decision
is non-monotonic in even-share → a per-line BADNESS / lookahead property.

This dumps, from the Word PDF render-truth, each contested paragraph's lines with
per-char advances (compression) + the wrap boundary + the NEXT line, to find the
feature that makes Word oidashi {400,434,465} but oikomi the +1 root.
"""
import sys, fitz
sys.stdout.reconfigure(encoding="utf-8")

PDF = r"C:\tmp\nedocontract_word.pdf"
# (word_i, page(1-based), distinctive-substring, label)
TARGETS = [
    (337, 20, "乙が大学等における", "+1 spill root (Word fits)"),
    (400, 26, "乙の責に帰すべき", "-1 over-fit (Word WRAPS)"),
    (434, 28, "甲又は機構は、不正等の事実", "-1 over-fit (Word WRAPS)"),
    (465, 30, "機構は共有知的財産権の自己", "-1 over-fit (Word WRAPS)"),
]


def page_lines(page):
    """Return list of lines; each = {'y':baseline, 'chars':[(ch,x0,x1)...], 'text':str, 'x0','x1'}."""
    d = page.get_text("rawdict")
    out = []
    for blk in d["blocks"]:
        if blk.get("type") != 0:
            continue
        for ln in blk.get("lines", []):
            chars = []
            for sp in ln.get("spans", []):
                for c in sp.get("chars", []):
                    bb = c["bbox"]
                    chars.append((c["c"], bb[0], bb[2]))
            if not chars:
                continue
            chars.sort(key=lambda t: t[1])
            txt = "".join(c[0] for c in chars)
            out.append({"y": round(ln["bbox"][1], 1), "x0": round(chars[0][1], 1),
                        "x1": round(chars[-1][2], 1), "chars": chars, "text": txt})
    out.sort(key=lambda l: (l["y"], l["x0"]))
    return out


def main():
    doc = fitz.open(PDF)
    for wi, pg, prefix, label in TARGETS:
        page = doc.load_page(pg - 1)
        lines = page_lines(page)
        # find the start line by distinctive substring (search anywhere on page)
        start = None
        for i, ln in enumerate(lines):
            if prefix[:8] in ln["text"]:
                start = i
                break
        print(f"\n===== word_i={wi} [{label}] page {pg} =====")
        if start is None:
            print("  prefix NOT FOUND; first lines on page:")
            for ln in lines[:6]:
                print("   ", repr(ln["text"][:40]))
            continue
        # print this paragraph's lines (until x0 jumps back to margin = next para,
        # heuristic: stop after 5 lines or when a clear new marker appears)
        margin_x0 = min(l["x0"] for l in lines)
        right_max = max(l["x1"] for l in lines)
        print(f"  page margin_x0~{margin_x0} right_max~{right_max}")
        for j in range(start, min(start + 5, len(lines))):
            ln = lines[j]
            chars = ln["chars"]
            # advances of last 6 chars
            tail = []
            for k in range(max(0, len(chars) - 6), len(chars)):
                ch, x0, x1 = chars[k]
                adv = round(x1 - x0, 2)
                tail.append(f"{ch}{adv}")
            nxt = lines[j + 1]["text"][:6] if j + 1 < len(lines) else "<end>"
            print(f"  L{j-start}: x0={ln['x0']} x1={ln['x1']} n={len(chars)} "
                  f"end={ln['text'][-3:]!r} -> next={nxt!r}")
            print(f"        tail-adv: {' '.join(tail)}")
    doc.close()


if __name__ == "__main__":
    main()
