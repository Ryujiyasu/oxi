"""
Pixel-sampling pass of the Tick 2-3 measurement.

The object-model pass (`measure_comments_tracked_changes_com.py`) covers the
structural side — Word's Revisions / Comments collections agree with our
authored fixture XML. This script covers the *visual* side: author RGB,
underline thickness, strikethrough Y, and balloon geometry that the renderer
rows (R-01 .. R-09) depend on.

Pipeline per fixture:

  1. Open the fixture in Word (Documents.Open, ReadOnly=True).
  2. Switch the view to Print Layout with all markup visible
     (`Document.ActiveWindow.View.RevisionsView = wdRevisionsViewFinal/Original`
     would change the markup; we want the default markup-on view).
  3. For each `Revision` and `Comment`, capture the page-relative
     `Information(wdHorizontalPositionRelativeToPage=5,
     wdVerticalPositionRelativeToPage=6)` rectangles.
  4. Save the document as PDF and rasterize page 1 (or all pages) to PNG at
     150 DPI via PyMuPDF.
  5. Convert the captured rectangles from pt → pixels (× 150/72) and sample
     RGB triplets at strategic offsets:
        * Inside a `<w:ins>` glyph cell → author ink RGB (R-01/R-02).
        * Centerline of a `<w:del>` glyph cell → strikethrough RGB (R-03).
        * Just below a `<w:del>` glyph cell → strikethrough Y offset.
        * Right-margin column at the comment's reference Y → balloon
          presence + fill RGB (R-04/R-05).

Output:
    tools/metrics/output/comments_tracked_changes_pixels.json

Run:
    python tools/metrics/measure_comments_tracked_changes_pixels.py
"""

from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any

try:
    import win32com.client as win32  # type: ignore
except ImportError:
    print("pywin32 not installed; cannot run COM measurement.", file=sys.stderr)
    sys.exit(1)

try:
    import fitz  # PyMuPDF
except ImportError:
    print("PyMuPDF not installed; cannot rasterize PDF.", file=sys.stderr)
    sys.exit(1)

try:
    from PIL import Image
except ImportError:
    print("Pillow not installed; cannot pixel-sample.", file=sys.stderr)
    sys.exit(1)


FIXTURES_DIR = Path(__file__).resolve().parents[1] / "fixtures" / "comments_samples"
OUT_DIR = Path(__file__).resolve().parent / "output"
OUT_PATH = OUT_DIR / "comments_tracked_changes_pixels.json"
PNG_DIR = OUT_DIR / "comments_tracked_changes_png"
PDF_DIR = OUT_DIR / "comments_tracked_changes_pdf"

DPI = 150
PT_PER_INCH = 72.0
PX_PER_PT = DPI / PT_PER_INCH  # 2.0833…


# Word enums we depend on
WD_HORIZONTAL_POSITION_RELATIVE_TO_PAGE = 5
WD_VERTICAL_POSITION_RELATIVE_TO_PAGE = 6
WD_ACTIVE_END_PAGE_NUMBER = 3
WD_FORMAT_PDF = 17
# wdExportFormat
WD_EXPORT_FORMAT_PDF = 17
# wdExportItem
WD_EXPORT_DOCUMENT_CONTENT = 0
WD_EXPORT_DOCUMENT_WITH_MARKUP = 7

WD_REVISION_TYPE = {
    0: "wdNoRevision", 1: "wdRevisionInsert", 2: "wdRevisionDelete",
    3: "wdRevisionProperty", 14: "wdRevisionMovedFrom", 15: "wdRevisionMovedTo",
}


def safe(fn, default=None):
    try:
        return fn()
    except Exception:
        return default


def pt_to_px(pt: float) -> int:
    return int(round(pt * PX_PER_PT))


def render_to_png(doc, pdf_path: Path, png_path_template: Path, with_markup: bool) -> list[Path]:
    """ExportAsFixedFormat → PDF → PNG per page. Returns list of png paths.

    `with_markup` = True asks Word to include comment balloons + revision
    markup in the PDF. Without this, SaveAs2(FileFormat=PDF) drops the markup
    layer and we never see balloons in the rasterized output.
    """
    pdf_path.parent.mkdir(parents=True, exist_ok=True)
    if pdf_path.exists():
        pdf_path.unlink()
    item = WD_EXPORT_DOCUMENT_WITH_MARKUP if with_markup else WD_EXPORT_DOCUMENT_CONTENT
    doc.ExportAsFixedFormat(
        OutputFileName=str(pdf_path.resolve()),
        ExportFormat=WD_EXPORT_FORMAT_PDF,
        OpenAfterExport=False,
        OptimizeFor=0,             # wdExportOptimizeForPrint
        Range=0,                   # wdExportAllDocument
        Item=item,
        IncludeDocProps=False,
        KeepIRM=True,
        CreateBookmarks=0,
        DocStructureTags=False,
        BitmapMissingFonts=True,
        UseISO19005_1=False,
    )

    pdf = fitz.open(str(pdf_path))
    zoom = DPI / PT_PER_INCH
    mat = fitz.Matrix(zoom, zoom)
    out: list[Path] = []
    for i, page in enumerate(pdf):
        png_path = png_path_template.with_name(
            png_path_template.stem + f"_p{i+1}" + png_path_template.suffix
        )
        png_path.parent.mkdir(parents=True, exist_ok=True)
        pix = page.get_pixmap(matrix=mat)
        pix.save(str(png_path))
        out.append(png_path)
    pdf.close()
    return out


def sample_pixel(im: Image.Image, x_px: int, y_px: int, span: int = 0) -> tuple[int, int, int] | None:
    """Sample one pixel, or the median of a (2*span+1)² block."""
    w, h = im.size
    if not (0 <= x_px < w and 0 <= y_px < h):
        return None
    if span == 0:
        rgb = im.getpixel((x_px, y_px))
        if isinstance(rgb, int):
            return (rgb, rgb, rgb)
        return rgb[:3]
    # Median over a small neighbourhood — robust to subpixel anti-aliasing.
    pixels = []
    for dy in range(-span, span + 1):
        for dx in range(-span, span + 1):
            xx, yy = x_px + dx, y_px + dy
            if 0 <= xx < w and 0 <= yy < h:
                p = im.getpixel((xx, yy))
                if isinstance(p, int):
                    pixels.append((p, p, p))
                else:
                    pixels.append(p[:3])
    if not pixels:
        return None
    pixels.sort()
    return pixels[len(pixels) // 2]


def _saturation(rgb: tuple[int, int, int]) -> float:
    """HSV-style saturation in [0, 255]. Black/white/grey → 0; pure color → 255."""
    mn = min(rgb); mx = max(rgb)
    if mx == 0:
        return 0.0
    return (mx - mn) * 255.0 / mx


def find_ink_pixel_in_block(
    im: Image.Image,
    x0: int, y0: int, x1: int, y1: int,
    prefer_saturated: bool = True,
) -> tuple[int, int, tuple[int, int, int]] | None:
    """Find the most-likely-ink pixel inside (x0..x1, y0..y1).

    If `prefer_saturated` is True (the default), prefer colored ink over black
    — author-revision text is colored (red/blue/green/purple) while regular
    body text is black. This avoids leaking into adjacent body text when the
    search window is loose.
    """
    w, h = im.size
    x0 = max(0, x0); y0 = max(0, y0)
    x1 = min(w, x1); y1 = min(h, y1)

    best_sat: tuple[int, int, tuple[int, int, int], float] | None = None  # (x, y, rgb, lum)
    best_dark: tuple[int, int, tuple[int, int, int], float] | None = None
    for y in range(y0, y1):
        for x in range(x0, x1):
            p = im.getpixel((x, y))
            if isinstance(p, int):
                rgb = (p, p, p)
            else:
                rgb = p[:3]
            lum = 0.299 * rgb[0] + 0.587 * rgb[1] + 0.114 * rgb[2]
            if lum >= 240:
                continue  # white-ish background, skip
            sat = _saturation(rgb)
            if prefer_saturated and sat > 80 and lum < 220:
                # Colored ink: keep the darkest most-saturated pixel.
                key = lum - sat  # lower = better (darker, more saturated)
                if best_sat is None or key < best_sat[3]:
                    best_sat = (x, y, rgb, key)
            else:
                if best_dark is None or lum < best_dark[3]:
                    best_dark = (x, y, rgb, lum)
    chosen = best_sat or best_dark
    if chosen is None:
        return None
    return (chosen[0], chosen[1], chosen[2])


def measure_revision(rev, page_pngs: list[Image.Image]) -> dict[str, Any]:
    """Sample the rendered ink for a single Revision and report RGB + position."""
    info: dict[str, Any] = {
        "index": int(rev.Index),
        "type_raw": int(rev.Type),
        "type_name": WD_REVISION_TYPE.get(int(rev.Type), f"raw={int(rev.Type)}"),
        "author": rev.Author,
        "range_text": rev.Range.Text[:80] if rev.Range.Text else "",
    }
    rng = rev.Range
    try:
        page = int(rng.Information(WD_ACTIVE_END_PAGE_NUMBER))
    except Exception:
        page = 1
    info["page"] = page

    try:
        x_pt = float(rng.Information(WD_HORIZONTAL_POSITION_RELATIVE_TO_PAGE))
        y_pt = float(rng.Information(WD_VERTICAL_POSITION_RELATIVE_TO_PAGE))
    except Exception:
        info["error"] = "could_not_get_position"
        return info
    info["x_pt"] = round(x_pt, 2)
    info["y_pt"] = round(y_pt, 2)

    # Convert to px coordinates of the page PNG.
    x_px = pt_to_px(x_pt)
    y_px = pt_to_px(y_pt)
    info["x_px"] = x_px
    info["y_px"] = y_px

    page_idx = page - 1
    if not (0 <= page_idx < len(page_pngs)):
        info["error"] = f"page_{page}_not_rendered"
        return info
    im = page_pngs[page_idx]

    # Search a narrow window starting at COM-reported (x, y). 12pt wide × 14pt
    # tall covers ~1.5 glyphs of 11pt body text (the intent is to land on the
    # FIRST glyph of the revision, not its neighbours). Prefer saturated ink
    # so we never leak the adjacent body text's black pixels.
    win_x = pt_to_px(12)
    win_y = pt_to_px(14)
    ink = find_ink_pixel_in_block(im, x_px, y_px, x_px + win_x, y_px + win_y)
    if ink is None:
        info["ink_rgb"] = None
        info["ink_search_window"] = [x_px, y_px, x_px + win_x, y_px + win_y]
        info["note"] = "no non-white pixel found in COM-reported window"
        return info
    ink_x, ink_y, ink_rgb = ink
    info["ink_rgb"] = list(ink_rgb)
    info["ink_x_px"] = ink_x
    info["ink_y_px"] = ink_y
    info["ink_dy_pt"] = round((ink_y - y_px) / PX_PER_PT, 2)
    info["ink_dx_pt"] = round((ink_x - x_px) / PX_PER_PT, 2)
    return info


def measure_comment(cmt, page_pngs: list[Image.Image], page_w_pt: float) -> dict[str, Any]:
    info: dict[str, Any] = {
        "index": int(cmt.Index),
        "author": cmt.Author,
        "scope_text": cmt.Scope.Text[:80] if cmt.Scope.Text else "",
        "comment_text": cmt.Range.Text[:80] if cmt.Range.Text else "",
    }
    # The commentReference marker (cmt.Reference) is the inline anchor — its
    # Y is the same as the line containing the marker, but the BALLOON aligns
    # with the FIRST character of the comment range (Scope.Start). Use scope
    # start; fall back to Reference if Scope is empty.
    scope = cmt.Scope
    rng = scope if scope.End > scope.Start else (cmt.Reference if hasattr(cmt, "Reference") else scope)
    try:
        page = int(rng.Information(WD_ACTIVE_END_PAGE_NUMBER))
        x_pt = float(rng.Information(WD_HORIZONTAL_POSITION_RELATIVE_TO_PAGE))
        y_pt = float(rng.Information(WD_VERTICAL_POSITION_RELATIVE_TO_PAGE))
    except Exception:
        info["error"] = "could_not_get_position"
        return info
    info["page"] = page
    info["scope_x_pt"] = round(x_pt, 2)
    info["scope_y_pt"] = round(y_pt, 2)

    page_idx = page - 1
    if not (0 <= page_idx < len(page_pngs)):
        info["error"] = f"page_{page}_not_rendered"
        return info
    im = page_pngs[page_idx]

    # In Print Layout with comments enabled, Word *compresses* the body and
    # places balloons in the right portion of the page. The balloon column is
    # typically the right ~250pt of a 595pt-wide A4 page. Scan from the
    # midpoint of the page (rightward) and from scope_y - 30pt to scope_y +
    # 200pt (covering the balloon's typical Y range — Word may push it down
    # if higher balloons stack above it).
    y_px = pt_to_px(y_pt)
    page_w_px = pt_to_px(page_w_pt)
    body_right_px = pt_to_px(page_w_pt * 0.5)  # rightmost half of the page
    body_right_px = max(0, body_right_px)
    y_min = max(0, y_px - pt_to_px(30))
    y_max = min(im.size[1], y_px + pt_to_px(200))

    found_x = None
    found_rgb = None
    found_y = None
    # column-scan: for each x, check the y range; we want the leftmost
    # non-white column inside the right-half balloon strip.
    for x in range(body_right_px, min(im.size[0], page_w_px)):
        for yy in range(y_min, y_max):
            p = im.getpixel((x, yy))
            if isinstance(p, int):
                rgb = (p, p, p)
            else:
                rgb = p[:3]
            lum = 0.299 * rgb[0] + 0.587 * rgb[1] + 0.114 * rgb[2]
            if lum < 240:
                found_x = x
                found_rgb = rgb
                found_y = yy
                break
        if found_x is not None:
            break

    if found_x is None:
        info["balloon"] = None
        info["note"] = "no balloon ink found in right half of page"
        return info

    info["balloon_left_px"] = found_x
    info["balloon_left_pt"] = round(found_x / PX_PER_PT, 2)
    info["balloon_first_y_px"] = found_y
    info["balloon_first_y_pt"] = round(found_y / PX_PER_PT, 2)
    info["balloon_first_rgb"] = list(found_rgb)
    info["balloon_left_dx_from_pageright_pt"] = round(page_w_pt - found_x / PX_PER_PT, 2)

    # Probe the balloon further: scan rightward from balloon_left to find the
    # right edge (last non-white at this y); scan upward/downward at
    # balloon_left+5px to find the balloon's top/bottom Y.
    rightmost = found_x
    for x in range(found_x, min(im.size[0], page_w_px)):
        p = im.getpixel((x, found_y))
        if isinstance(p, int):
            rgb = (p, p, p)
        else:
            rgb = p[:3]
        lum = 0.299 * rgb[0] + 0.587 * rgb[1] + 0.114 * rgb[2]
        if lum < 240:
            rightmost = x
    info["balloon_right_px"] = rightmost
    info["balloon_width_pt"] = round((rightmost - found_x) / PX_PER_PT, 2)

    # Vertical extent: scan y from balloon_first_y up and down, looking for
    # ANY non-white pixel in the [balloon_left, balloon_right] strip.
    def row_has_ink(y: int) -> bool:
        if not (0 <= y < im.size[1]):
            return False
        for x in range(found_x, min(im.size[0], rightmost + 1), 4):
            p = im.getpixel((x, y))
            rgb = p if isinstance(p, tuple) else (p, p, p)
            lum = 0.299 * rgb[0] + 0.587 * rgb[1] + 0.114 * rgb[2]
            if lum < 240:
                return True
        return False

    top_y = found_y
    while top_y > 0 and row_has_ink(top_y - 1):
        top_y -= 1
    bot_y = found_y
    while bot_y < im.size[1] - 1 and row_has_ink(bot_y + 1):
        bot_y += 1
    info["balloon_top_pt"] = round(top_y / PX_PER_PT, 2)
    info["balloon_bottom_pt"] = round(bot_y / PX_PER_PT, 2)
    info["balloon_height_pt"] = round((bot_y - top_y) / PX_PER_PT, 2)

    # Anchor offset: how far below the comment's scope_y does the balloon top
    # sit? This is what R-05 needs for balloon placement.
    info["balloon_top_dy_from_scope_pt"] = round((top_y - y_px) / PX_PER_PT, 2)
    return info


def main() -> int:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    PNG_DIR.mkdir(parents=True, exist_ok=True)
    PDF_DIR.mkdir(parents=True, exist_ok=True)

    fixtures = sorted(FIXTURES_DIR.glob("fixture_*.docx"))
    if not fixtures:
        print(f"No fixtures found under {FIXTURES_DIR}", file=sys.stderr)
        return 1

    app = win32.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0

    results: list[dict[str, Any]] = []
    try:
        for fx in fixtures:
            entry: dict[str, Any] = {"file": fx.name}
            print(f"  measuring {fx.name} ...")
            try:
                doc = app.Documents.Open(
                    str(fx.resolve()),
                    ReadOnly=False,        # SaveAs2 needs write access on the *document* (writes to PDF, not the docx)
                    AddToRecentFiles=False,
                    ConfirmConversions=False,
                )
            except Exception as e:
                entry["error"] = f"open_failed: {e}"
                results.append(entry)
                continue

            try:
                # Page geometry: PageSetup returns pt directly.
                ps = doc.PageSetup
                page_w_pt = float(ps.PageWidth)
                page_h_pt = float(ps.PageHeight)
                entry["page_width_pt"] = round(page_w_pt, 2)
                entry["page_height_pt"] = round(page_h_pt, 2)

                # Force markup display to "Final: Show Markup" with comment
                # balloons enabled. By default Word's `View.ShowComments` is
                # False, which means even with `Item=wdExportDocumentWithMarkup`
                # the right-margin balloons don't render in the PDF.
                view = doc.Windows(1).View
                try:
                    view.RevisionsView = 0           # wdRevisionsViewFinal
                    view.ShowRevisionsAndComments = True
                    view.ShowComments = True         # the load-bearing flag for balloons
                    view.ShowFormatChanges = True
                    view.ShowInsertionsAndDeletions = True
                    # Leave MarkupMode at default (2 = wdMixedRevisions): comments
                    # in right-margin balloons, ins/del inline. R-01/R-03 expect
                    # inline rendering so this matches the spec the renderer rows
                    # need to mimic.
                except Exception as e:
                    entry["view_setup_warn"] = repr(e)
                entry["balloon_width_pt"] = round(float(safe(lambda: view.RevisionsBalloonWidth, 0.0)), 2)
                entry["balloon_side"] = int(safe(lambda: view.RevisionsBalloonSide, -1))

                # Render to PDF + PNGs (with markup so balloons appear).
                pdf_path = PDF_DIR / (fx.stem + ".pdf")
                png_template = PNG_DIR / (fx.stem + ".png")
                png_paths = render_to_png(doc, pdf_path, png_template, with_markup=True)
                entry["page_png_paths"] = [str(p.resolve()) for p in png_paths]

                page_pngs = [Image.open(p).convert("RGB") for p in png_paths]

                revs = []
                for i in range(1, doc.Revisions.Count + 1):
                    revs.append(measure_revision(doc.Revisions(i), page_pngs))
                cmts = []
                for i in range(1, doc.Comments.Count + 1):
                    cmts.append(measure_comment(doc.Comments(i), page_pngs, page_w_pt))
                entry["revisions"] = revs
                entry["comments"] = cmts
            except Exception as e:
                entry["error"] = f"measure_failed: {type(e).__name__}: {e}"
            finally:
                doc.Close(SaveChanges=0)
            results.append(entry)

        payload = {
            "generated": "2026-04-25",
            "word_version": safe(lambda: app.Version, "?"),
            "fixtures_dir": str(FIXTURES_DIR),
            "dpi": DPI,
            "px_per_pt": round(PX_PER_PT, 4),
            "note": (
                "Tick 2-3 pixel-sampling pass. For each Revision and Comment, "
                "we capture COM-reported page coordinates, render the page to "
                "PNG via Word→PDF→PyMuPDF at 150 DPI, then sample RGB at the "
                "ink centre / inside the right-margin balloon band. Author "
                "RGB, strikethrough Y offset, and balloon left-edge offsets "
                "are surfaced for R-01..R-05."
            ),
            "results": results,
        }
        OUT_PATH.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")
        print(f"\nWrote {OUT_PATH}")
    finally:
        app.Quit()

    print()
    print("=== Summary ===")
    for r in results:
        if "error" in r:
            print(f"  {r['file']:50s}  ERROR: {r['error']}")
            continue
        n_rev = len(r.get("revisions", []))
        n_cmt = len(r.get("comments", []))
        author_rgbs = sorted({tuple(rev.get("ink_rgb")) for rev in r.get("revisions", []) if rev.get("ink_rgb")})
        balloon = [c for c in r.get("comments", []) if c.get("balloon_left_px") is not None]
        print(f"  {r['file']:50s}  revs={n_rev} cmts={n_cmt} | author RGBs={author_rgbs} | balloons_found={len(balloon)}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
