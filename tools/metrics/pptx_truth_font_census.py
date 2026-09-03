"""What the corpus's truth PDFs say about the faces PowerPoint actually used.

Each truth PDF embeds a subset of every face it drew, and names it. Two things
in that record are checkable without opening PowerPoint:

  * a face named `-Regular` (or with no style suffix) whose descriptor carries a
    non-zero `/ItalicAngle` was drawn with SLANTED outlines -- the machine
    handed PowerPoint the italic file for an upright request;
  * the `/Widths` array is the advance PowerPoint used, in 1/1000 em, so it can
    be compared against the installed file's own `hmtx`.

Deck 47 is the case this was written for: its truth PDF embeds
`Caladea-Regular` at `/ItalicAngle -9` with the italic file's advances, while
live PowerPoint measured through COM agrees with the upright file. A PDF that
disagrees with today's PowerPoint is a stale reference, not a defect in Oxi.

    python tools/metrics/pptx_truth_font_census.py [--doc 47]
"""
import argparse
import glob
import os
import re
import struct
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
PDFS = os.path.join(ROOT, "pipeline_data", "pptx_benchmark", "ssim_pptx", "ppt_pdf")
FONTDIRS = ["C:/Windows/Fonts", os.path.expandvars("%LOCALAPPDATA%/Microsoft/Windows/Fonts")]


def file_advances(path, chars):
    b = open(path, "rb").read()
    if b[:4] not in (b"\x00\x01\x00\x00", b"true"):
        return None
    n = struct.unpack(">H", b[4:6])[0]
    tabs = {}
    for i in range(n):
        o = 12 + 16 * i
        tabs[b[o:o + 4].decode("latin1")] = struct.unpack(">I", b[o + 8:o + 12])[0]
    if not {"head", "hhea", "hmtx", "cmap"} <= set(tabs):
        return None
    upem = struct.unpack(">H", b[tabs["head"] + 18:tabs["head"] + 20])[0]
    nhm = struct.unpack(">H", b[tabs["hhea"] + 34:tabs["hhea"] + 36])[0]
    co = tabs["cmap"]
    nt = struct.unpack(">H", b[co + 2:co + 4])[0]
    sub = None
    for i in range(nt):
        pid, eid, off = struct.unpack(">HHI", b[co + 4 + 8 * i:co + 12 + 8 * i])
        if (pid, eid) in ((3, 1), (3, 10), (0, 3), (0, 4)):
            sub = co + off
    if sub is None or struct.unpack(">H", b[sub:sub + 2])[0] != 4:
        return None
    segx2 = struct.unpack(">H", b[sub + 6:sub + 8])[0]
    seg = segx2 // 2
    ends = [struct.unpack(">H", b[sub + 14 + 2 * i:sub + 16 + 2 * i])[0] for i in range(seg)]
    sto = sub + 16 + segx2
    starts = [struct.unpack(">H", b[sto + 2 * i:sto + 2 + 2 * i])[0] for i in range(seg)]
    dto = sto + segx2
    deltas = [struct.unpack(">h", b[dto + 2 * i:dto + 2 + 2 * i])[0] for i in range(seg)]
    rto = dto + segx2
    out = {}
    for ch in chars:
        cc = ord(ch)
        g = 0
        for i in range(seg):
            if starts[i] <= cc <= ends[i]:
                ro = struct.unpack(">H", b[rto + 2 * i:rto + 2 + 2 * i])[0]
                if ro == 0:
                    g = (cc + deltas[i]) & 0xFFFF
                else:
                    gi = rto + 2 * i + ro + 2 * (cc - starts[i])
                    gg = struct.unpack(">H", b[gi:gi + 2])[0]
                    g = (gg + deltas[i]) & 0xFFFF if gg else 0
                break
        gi = g if g < nhm else nhm - 1
        out[ch] = round(struct.unpack(">H", b[tabs["hmtx"] + 4 * gi:tabs["hmtx"] + 2 + 4 * gi])[0]
                        * 1000.0 / upem)
    return out


def installed_file(base):
    for d in FONTDIRS:
        for cand in (base + ".ttf", base.replace("-", "") + ".ttf"):
            p = os.path.join(d, cand)
            if os.path.exists(p):
                return p
    return None


def main():
    import pymupdf

    ap = argparse.ArgumentParser()
    ap.add_argument("--doc", default=None)
    args = ap.parse_args()

    pdfs = sorted(glob.glob(os.path.join(PDFS, (args.doc or "*") + ".pdf")))
    slanted, mismatched, total = [], [], 0
    for path in pdfs:
        doc = os.path.basename(path)[:-4]
        pdf = pymupdf.open(path)
        seen = set()
        for pno in range(len(pdf)):
            for f in pdf[pno].get_fonts(full=True):
                xref, name = f[0], f[3]
                if xref in seen:
                    continue
                seen.add(xref)
                total += 1
                base = name.split("+")[-1]
                d = pdf.xref_object(xref)
                fd = re.search(r"/FontDescriptor (\d+) 0 R", d)
                desc = pdf.xref_object(int(fd.group(1))) if fd else ""
                ia = re.search(r"/ItalicAngle ([-\d.]+)", desc)
                angle = float(ia.group(1)) if ia else 0.0
                upright_name = not re.search(r"italic|oblique|,I|-I\b", base, re.I)
                if angle and upright_name:
                    slanted.append((doc, base, angle))
                fc = re.search(r"/FirstChar (\d+)", d)
                m = re.search(r"/Widths (\d+) 0 R", d)
                if not fc or not m:
                    continue
                first = int(fc.group(1))
                nums = [int(float(x)) for x in re.findall(r"[-\d.]+", pdf.xref_object(int(m.group(1))))]
                src = installed_file(base)
                if not src:
                    continue
                want = file_advances(src, "Yo")
                if not want:
                    continue
                got = {}
                for ch in "Yo":
                    i = ord(ch) - first
                    if 0 <= i < len(nums) and nums[i]:
                        got[ch] = nums[i]
                if len(got) == 2 and any(abs(got[c] - want[c]) > 2 for c in got):
                    mismatched.append((doc, base, got, want))
    print("%d embedded faces across %d truth PDFs" % (total, len(pdfs)))
    print("named upright but drawn slanted: %d" % len(slanted))
    for doc, base, angle in slanted:
        print("   %-4s %-30s ItalicAngle=%s" % (doc, base, angle))
    print("advances disagree with the installed file of the same name: %d" % len(mismatched))
    for doc, base, got, want in mismatched:
        print("   %-4s %-30s pdf=%s file=%s" % (doc, base, got, want))


if __name__ == "__main__":
    main()
