# -*- coding: utf-8 -*-
"""Dump all text spans of a PDF (font / size / origin / bbox)."""
import sys
sys.stdout.reconfigure(encoding="utf-8")
import fitz

for path in sys.argv[1:]:
    print(f"=== {path}")
    doc = fitz.open(path)
    for pno in range(len(doc)):
        page = doc[pno]
        d = page.get_text("dict")
        for block in d["blocks"]:
            if block.get("type") != 0:
                continue
            for line in block["lines"]:
                for span in line["spans"]:
                    if span["text"].strip():
                        print(f"  p{pno+1} font={span['font']} size={span['size']:.2f} "
                              f"origin=({span['origin'][0]:.2f},{span['origin'][1]:.2f}) "
                              f"bbox=({span['bbox'][0]:.2f},{span['bbox'][1]:.2f},"
                              f"{span['bbox'][2]:.2f},{span['bbox'][3]:.2f}) "
                              f"text={span['text']!r}")
    doc.close()
