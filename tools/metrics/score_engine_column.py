# -*- coding: utf-8 -*-
"""Score one engine against the frozen Word truth and fold it into the result.

Every engine in the comparison leaves its output in the same place under a
different name — some as a PDF per document, some as a folder of page images —
and every one of them has to be scored the same way, against the same Word
PDFs, or the table means nothing. So there is one scorer rather than one per
engine.

    python tools/metrics/score_engine_column.py en D oo
    python tools/metrics/score_engine_column.py ja D silurus

The engine's own column is replaced; every other column in the file is left
exactly as it was. Nothing is written if the engine produced nothing.
"""
from __future__ import annotations

import json
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

import fitz
import numpy as np
from PIL import Image
from skimage.metrics import structural_similarity

REPO = Path(__file__).resolve().parents[2]
DPI = 150

# Where each engine leaves its work, and in what shape.
SHAPES = {
    "lo": ("lo_pdf", "pdf"),
    "oo": ("oo_pdf", "pdf"),
    "polaris": ("polaris_pdf", "pdf"),
    "silurus": ("silurus_png", "png"),
    "eigenpal": ("eigenpal_png", "png"),
    "betteroffice": ("betteroffice_png", "png"),
    "genoffice": ("genoffice_png", "png"),
    "officecli": ("officecli_png", "png"),
}


def rgb_from_pdf(pdf: fitz.Document, index: int) -> np.ndarray:
    page = pdf.load_page(index)
    pix = page.get_pixmap(dpi=DPI, colorspace=fitz.csRGB, alpha=False)
    return np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, 3)


def rgb_from_png(path: Path) -> np.ndarray:
    return np.array(Image.open(path).convert("RGB"))


def resize(candidate: np.ndarray, reference: np.ndarray) -> np.ndarray:
    if candidate.shape[:2] == reference.shape[:2]:
        return candidate
    height, width = reference.shape[:2]
    return np.array(Image.fromarray(candidate).resize((width, height), Image.LANCZOS))


def score(reference: np.ndarray, candidate: np.ndarray) -> float:
    candidate = resize(candidate, reference)
    return float(structural_similarity(reference, candidate, channel_axis=2))


def pages_of(where: Path, shape: str, doc_id: str) -> list:
    """The engine's pages for one document, in order, however it stored them."""
    if shape == "pdf":
        at = where / f"{doc_id}.pdf"
        return [at] if at.is_file() else []
    folder = where / doc_id
    if not folder.is_dir():
        return []
    found = []
    n = 1
    while (folder / f"p_p{n}.png").is_file():
        found.append(folder / f"p_p{n}.png")
        n += 1
    return found


def main() -> int:
    if len(sys.argv) < 4 or sys.argv[1] not in ("en", "ja") or sys.argv[3] not in SHAPES:
        print(f"usage: {Path(sys.argv[0]).name} [en|ja] <letter> "
              f"[{'|'.join(SHAPES)}]")
        return 2
    lang, letter, engine = sys.argv[1], sys.argv[2].upper(), sys.argv[3]
    out = REPO / "pipeline_data" / f"{lang}_benchmark" / f"ssim_blind{letter}50"
    word_pdf = out / "word_pdf"
    result_path = out / "_result.json"
    folder_name, shape = SHAPES[engine]
    where = out / folder_name

    if not result_path.is_file():
        print(f"no result yet at {result_path} — measure Oxi first")
        return 1
    if not where.is_dir():
        print(f"{engine} produced nothing at {where.name}")
        return 1

    data = json.loads(result_path.read_text(encoding="utf-8"))
    rows = data["docs"]

    def one(row: dict):
        doc_id = row["doc"]
        truth = word_pdf / f"{doc_id}.pdf"
        if not truth.is_file():
            return doc_id, None
        pages = pages_of(where, shape, doc_id)
        if not pages:
            return doc_id, None
        reference = fitz.open(truth)
        n_word = reference.page_count
        if shape == "pdf":
            theirs = fitz.open(pages[0])
            n_theirs = theirs.page_count
            get = lambda i: rgb_from_pdf(theirs, i)  # noqa: E731
        else:
            n_theirs = len(pages)
            get = lambda i: rgb_from_png(pages[i])  # noqa: E731
        scores = []
        for i in range(min(n_word, n_theirs)):
            scores.append(score(rgb_from_pdf(reference, i), get(i)))
        reference.close()
        if shape == "pdf":
            theirs.close()
        denom = max(n_word, n_theirs)
        return doc_id, {
            "pages": n_theirs,
            "page_delta": n_theirs - n_word,
            "common_pages": len(scores),
            "common_mean": round(sum(scores) / len(scores), 6) if scores else None,
            "penalized_mean": round(sum(scores) / denom, 6) if denom else None,
            "page_min": round(min(scores), 6) if scores else None,
        }

    scored = missing = 0
    with ThreadPoolExecutor(max_workers=4) as pool:
        futures = {pool.submit(one, r): r for r in rows}
        for n, fut in enumerate(as_completed(futures), 1):
            doc_id, column = fut.result()
            for row in rows:
                if row["doc"] == doc_id:
                    if column is None:
                        row[engine] = None
                        missing += 1
                    else:
                        row[engine] = column
                        scored += 1
                    break
            if n % 10 == 0:
                print(f"  {n}/{len(rows)}", flush=True)

    result_path.write_text(json.dumps(data, indent=1, ensure_ascii=False), encoding="utf-8")
    got = [r[engine]["common_mean"] for r in rows
           if r.get(engine) and r[engine].get("common_mean") is not None]
    pen = [r[engine]["penalized_mean"] for r in rows
           if r.get(engine) and r[engine].get("penalized_mean") is not None]
    match = sum(1 for r in rows if r.get(engine) and r[engine].get("page_delta") == 0)
    print(f"\n{engine}: {scored} scored, {missing} produced nothing")
    if got:
        print(f"  common {sum(got)/len(got):.4f}  penalized {sum(pen)/len(pen):.4f}"
              f"  pages match {match}/{len(rows)}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    raise SystemExit(main())
