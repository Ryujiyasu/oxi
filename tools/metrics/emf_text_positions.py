# -*- coding: utf-8 -*-
"""S492 BIG-JOB FOUNDATION: parse a Word EMF's EMR_EXTTEXTOUTW (type 84) records to extract
Word's EXACT rendered glyph reference points (x,y) + text. This is RENDER-TRUTH (the EMF is the
source of the screenshots), solving the COM-logical != render wall that blocked per-line/per-glyph
matching. Maps logical units -> points via the EMR_HEADER frame. Usage: emf_text_positions.py
<file.emf> [out.json]. cp932-safe (UTF-8 file, results to JSON)."""
import struct, sys, json

EMR_HEADER = 1
EMR_EXTTEXTOUTW = 84
EMR_EXTTEXTOUTA = 83


def parse(path):
    data = open(path, 'rb').read()
    # Header record
    itype, nsize = struct.unpack_from('<II', data, 0)
    assert itype == EMR_HEADER, "not an EMF (first record type=%d)" % itype
    rclFrame = struct.unpack_from('<iiii', data, 24)          # .01mm units
    szlDevice = struct.unpack_from('<ii', data, 72)           # device px (after the fixed header fields)
    szlMM = struct.unpack_from('<ii', data, 80)               # device mm
    # logical units: EMF records use device units of the reference DC.
    # ptlReference is in logical units == device pixels of szlDevice over rclFrame(.01mm).
    # pt = logical_px * (frame_mm / device_px) / 0.3528  ... we calibrate empirically below.
    recs = []
    off = nsize
    n = len(data)
    while off + 8 <= n:
        itype, nsize = struct.unpack_from('<II', data, off)
        if nsize < 8 or off + nsize > n:
            break
        if itype in (EMR_EXTTEXTOUTW, EMR_EXTTEXTOUTA):
            # EMRTEXT starts at record off+36
            refx, refy = struct.unpack_from('<ii', data, off + 36)
            nchars = struct.unpack_from('<I', data, off + 44)[0]
            offString = struct.unpack_from('<I', data, off + 48)[0]
            try:
                if itype == EMR_EXTTEXTOUTW:
                    s = data[off + offString: off + offString + nchars * 2].decode('utf-16-le', 'replace')
                else:
                    s = data[off + offString: off + offString + nchars].decode('cp932', 'replace')
            except Exception:
                s = ''
            recs.append({'x': refx, 'y': refy, 'n': nchars, 'text': s})
        off += nsize
    return rclFrame, szlDevice, szlMM, recs


def main():
    path = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) > 2 else None
    rclFrame, szlDevice, szlMM, recs = parse(path)
    # calibrate: device px -> pt. szlDevice px over szlMM mm. pt = px * (mm/px_dev?) ...
    # Simpler: the ptlReference logical units == 0.01mm if MM_HIMETRIC, OR device px.
    # Empirically: A4 page text y spans ~60..820pt. Find raw y range, scale to fit.
    ys = [r['y'] for r in recs if r['text'].strip()]
    xs = [r['x'] for r in recs if r['text'].strip()]
    info = {'rclFrame_001mm': rclFrame, 'szlDevice_px': szlDevice, 'szlMM': szlMM,
            'n_text_records': len(recs), 'raw_y_range': [min(ys), max(ys)] if ys else None,
            'raw_x_range': [min(xs), max(xs)] if xs else None}
    # candidate scales: device px -> pt = pt_per_px. If szlDevice & szlMM known: px->mm = mm/px, mm->pt=/0.352778
    sx_dev = (szlMM[0] / szlDevice[0]) / 0.352778 if szlDevice[0] else 0
    sy_dev = (szlMM[1] / szlDevice[1]) / 0.352778 if szlDevice[1] else 0
    info['scale_device_pt_per_unit'] = [round(sx_dev, 5), round(sy_dev, 5)]
    print(json.dumps(info, ensure_ascii=False, indent=1))
    print("\nfirst 12 text records (raw x,y, text):")
    for r in recs[:12]:
        print("  x=%8d y=%8d n=%2d %r" % (r['x'], r['y'], r['n'], r['text'][:14]))
    if out:
        json.dump({'info': info, 'records': recs}, open(out, 'w', encoding='utf-8'), ensure_ascii=False)
        print("wrote", out)


if __name__ == '__main__':
    main()
