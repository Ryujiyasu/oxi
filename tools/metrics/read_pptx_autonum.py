"""Read autonum.pdf spans (baseline origin / text / font / size) via fitz."""
import json
import sys

import fitz


def main(pdf):
    doc = fitz.open(pdf)
    out = []
    for i, page in enumerate(doc):
        d = page.get_text("rawdict")
        spans = []
        for b in d["blocks"]:
            if b["type"] != 0:
                continue
            for l in b["lines"]:
                for s in l["spans"]:
                    text = "".join(ch["c"] for ch in s.get("chars", []))
                    spans.append({
                        "x": round(s["origin"][0], 2),
                        "y": round(s["origin"][1], 2),
                        "text": text,
                        "font": s["font"],
                        "size": round(s["size"], 3),
                    })
        out.append({"page": i + 1, "spans": spans})
    print(json.dumps(out, ensure_ascii=False, indent=1))


if __name__ == "__main__":
    main(sys.argv[1])
