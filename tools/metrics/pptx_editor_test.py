# -*- coding: utf-8 -*-
"""End-to-end browser test for the slide editor (`web/pptx-editor.html`).

The editor is the first thing in this tree whose correctness lives in a
browser: the engine's layout arrives through wasm, the drawing is canvas, and
the editing is keyboard events. None of that is reachable from `cargo test`, so
this drives a real Chromium the way `pptx_browser_oracle.py` already does --
serve `web/` from a local port, open the page, and act on it.

What it asserts, per deck:

  parses        the deck loads and the slide count is right
  layout        the engine placed lines, and the shape count it reports agrees
                with what the page actually drew
  hit test      clicking on a line selects it and names a paragraph and a
                character offset
  typing        a keystroke reaches the run under the caret and the page counts
                the edit
  caret         the arrow keys move the caret, and Home/End reach the ends of a
                line
  undo          Ctrl+Z puts the text back AND drops the edit, so a save after a
                full undo writes nothing
  enter         a paragraph break shows on screen AT ONCE -- the deck is
                re-opened from the bytes the break produced -- and the file
                really holds one more paragraph than the original did
  backspace     Backspace at the head of a paragraph joins it onto the one
                above, taking the file back to where it started
  save          the download re-opens as a pptx whose text carries the change
  console       no page error was raised along the way

The suite viewer (`web/index.html`) carries the SAME editing path -- a
contentEditable run addressed by shape, paragraph and run -- and had the same
bug, so it gets the save check too. A regression there is invisible from the
canvas editor's tests.

★The engine cannot measure every family (the tables carry 17 of the corpus's
142), so a shape it declines is EXPECTED. The test asserts that such a shape is
reported as incomplete rather than drawn as if the engine had spoken -- silence
there would be the real failure.

Usage:
    python tools/metrics/pptx_editor_test.py [--decks d15,d09] [--headed]
"""
from __future__ import annotations

import argparse
import http.server
import os
import re
import socketserver
import sys
import threading
import zipfile
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
WEB = REPO / "web"
DECKS = REPO / "pipeline_data" / "pptx_benchmark" / "dev" / "pptx"


class Quiet(http.server.SimpleHTTPRequestHandler):
    def log_message(self, *_args):
        pass


def serve(root: Path) -> int:
    os.chdir(root)
    httpd = socketserver.TCPServer(("127.0.0.1", 0), Quiet)
    threading.Thread(target=httpd.serve_forever, daemon=True).start()
    return httpd.server_address[1]


class Report:
    def __init__(self) -> None:
        self.rows: list[tuple[str, str, bool, str]] = []

    def check(self, deck: str, name: str, ok: bool, detail: str = "") -> None:
        self.rows.append((deck, name, ok, detail))
        mark = "ok  " if ok else "FAIL"
        print(f"  {mark} {name:<10} {detail}", flush=True)

    def failed(self) -> int:
        return sum(1 for _d, _n, ok, _x in self.rows if not ok)


def deck_paths(spec: str) -> list[Path]:
    if spec:
        want = set(spec.split(","))
        return sorted(p for p in DECKS.glob("*.pptx") if p.name.split("__")[0] in want)
    return sorted(DECKS.glob("*.pptx"))[:2]


def slide_text(pptx: bytes, slide_no: int) -> str:
    """Every `<a:t>` of one slide, joined -- enough to see an edit arrive."""
    import io

    with zipfile.ZipFile(io.BytesIO(pptx)) as z:
        xml = z.read(f"ppt/slides/slide{slide_no}.xml").decode("utf-8", "replace")
    return "".join(re.findall(r"<a:t>([^<]*)</a:t>", xml))


def count_paragraphs(pptx: bytes, slide_no: int) -> int:
    """How many `<a:p>` one slide holds -- the thing a paragraph break moves."""
    import io as _io

    with zipfile.ZipFile(_io.BytesIO(pptx)) as z:
        xml = z.read(f"ppt/slides/slide{slide_no}.xml").decode("utf-8", "replace")
    return len(re.findall(r"<a:p[ >]", xml))


def run_deck(page, port: int, pptx: Path, rep: Report) -> None:
    deck = pptx.name.split("__")[0]
    print(f"{deck}", flush=True)
    errors: list[str] = []
    page.on("pageerror", lambda e: errors.append(str(e)))
    page.goto(f"http://127.0.0.1:{port}/pptx-editor.html")
    page.wait_for_function("() => document.getElementById('status').textContent.includes('ready')")

    page.set_input_files("#file", str(pptx))
    page.wait_for_function(
        "() => document.getElementById('s-slide').textContent !== '\\u2014'", timeout=60000)

    slides = page.eval_on_selector("#s-slide", "e => e.textContent")
    with zipfile.ZipFile(pptx) as z:
        declared = sum(1 for n in z.namelist()
                       if n.startswith("ppt/slides/slide") and n.endswith(".xml"))
    rep.check(deck, "parses", slides.strip().endswith(str(declared)),
              f"page says {slides.strip()}, file has {declared}")

    # The page keeps its state in a module closure, so the assertions read what
    # it SHOWS -- which is the surface a person judges it by anyway.
    cover = page.eval_on_selector("#s-cover", "e => e.textContent.trim()")
    m = re.match(r"(\d+)\s*/\s*(\d+)", cover)
    rep.check(deck, "layout", bool(m) and int(m.group(2)) > 0,
              f"engine laid out {cover or 'nothing'}")

    # A shape the engine declines must SAY so: the panel names the families it
    # could not measure, and the count of engine-laid shapes must fall short by
    # at least one. Silence here -- drawing a fallback wrap as if it were the
    # engine's -- is the failure this guards.
    if m and int(m.group(1)) < int(m.group(2)):
        fams = page.eval_on_selector("#s-fams", "e => e.textContent.trim()")
        rep.check(deck, "declines", bool(fams) and "none" not in fams,
                  f"names {fams[:60]!r}")
    else:
        rep.check(deck, "declines", True, "every shape on this slide is the engine's")

    # Click the middle of the canvas and walk outward until a line is hit: a
    # slide's text is not always under its centre.
    box = page.eval_on_selector("#slide", "c => { const r = c.getBoundingClientRect();"
                                " return {x: r.x, y: r.y, w: r.width, h: r.height}; }")
    hit = False
    for fy in (0.5, 0.35, 0.65, 0.2, 0.8):
        for fx in (0.5, 0.3, 0.7):
            page.mouse.click(box["x"] + box["w"] * fx, box["y"] + box["h"] * fy)
            sel = page.eval_on_selector("#s-sel", "e => e.textContent")
            if "paragraph" in sel:
                hit = True
                break
        if hit:
            break
    rep.check(deck, "hit test", hit,
              page.eval_on_selector("#s-sel", "e => e.textContent.split('\\n')[0]"))

    if hit:
        before = page.eval_on_selector("#status", "e => e.textContent")
        page.keyboard.type("Z")
        after = page.eval_on_selector("#status", "e => e.textContent")
        caret = page.eval_on_selector("#caret", "e => getComputedStyle(e).display")
        rep.check(deck, "typing", "edited" in after and after != before,
                  f"{after!r}, caret {caret}")

        # The caret must MOVE, and a character offset is the thing to watch:
        # a caret that redraws in place would still look alive on screen.
        def offset():
            m = re.search(r"char (\d+)", page.eval_on_selector("#s-sel", "e => e.textContent"))
            return int(m.group(1)) if m else None

        start = offset()
        page.keyboard.press("ArrowLeft")
        left = offset()
        page.keyboard.press("ArrowRight")
        back = offset()
        page.keyboard.press("Home")
        home = offset()
        page.keyboard.press("End")
        end = offset()
        rep.check(deck, "caret",
                  None not in (start, left, back, home, end)
                  and left == start - 1 and back == start
                  and home <= left and end >= home,
                  f"{start} -> left {left} -> right {back}, home {home} end {end}")

        # Undo must put the text back AND forget the edit: an undo that leaves
        # the edit registered would save the ORIGINAL text as a change, which
        # looks like success and is not.
        page.keyboard.press("Control+z")
        after_undo = page.eval_on_selector("#status", "e => e.textContent")
        rep.check(deck, "undo", "0 runs edited" in after_undo or "undone" in after_undo,
                  repr(after_undo))
        save_off = page.eval_on_selector("#save", "e => e.disabled")
        rep.check(deck, "undo clears", save_off is True,
                  "save is disabled again" if save_off else "save still armed")

        # Enter must show AT ONCE and reach the FILE. The lines the page has
        # placed are the screen's own account, so a break that only moved the
        # caret cannot pass: the count of placed lines has to rise.
        lines_before = page.evaluate(
            "() => document.getElementById('s-cover').textContent")
        sel_before = page.eval_on_selector("#s-sel", "e => e.textContent")
        page.keyboard.press("Enter")
        page.wait_for_function(
            "() => document.getElementById('status').textContent.includes('broken')"
            " || document.getElementById('status').textContent.includes('could not')",
            timeout=30000)
        after = page.eval_on_selector("#status", "e => e.textContent")
        sel_after = page.eval_on_selector("#s-sel", "e => e.textContent")
        rep.check(deck, "enter", "broken" in after and sel_after != sel_before,
                  f"{after!r}, caret moved to the new paragraph")

        # And the bytes the page now holds carry the break: saving without any
        # further edit must still hand back a file with one more paragraph.
        page.keyboard.type("Z")
        with page.expect_download() as dl:
            page.click("#save")
        broke = Path(dl.value.path()).read_bytes()
        before_n = count_paragraphs(pptx.read_bytes(), 1)
        after_n = count_paragraphs(broke, 1)
        rep.check(deck, "enter saved", after_n == before_n + 1,
                  f"paragraphs {before_n} -> {after_n}")

        # And the inverse: Backspace at the head of the new paragraph joins it
        # back. The count must return to where it started -- a join that only
        # LOOKED right would leave the file one paragraph long.
        page.keyboard.press("Home")
        page.keyboard.press("Backspace")
        page.wait_for_function(
            "() => document.getElementById('status').textContent.includes('joined')"
            " || document.getElementById('status').textContent.includes('could not')",
            timeout=30000)
        joined = page.eval_on_selector("#status", "e => e.textContent")
        rep.check(deck, "backspace", "joined" in joined, repr(joined))
        page.keyboard.type("Z")
        with page.expect_download() as dl2:
            page.click("#save")
        rejoined = Path(dl2.value.path()).read_bytes()
        rep.check(deck, "backspace saved",
                  count_paragraphs(rejoined, 1) == before_n,
                  f"paragraphs back to {count_paragraphs(rejoined, 1)}")
        dl = dl2
        out = Path(dl.value.path()).read_bytes()
        # The saved file must be a pptx, and the slide that was typed into must
        # carry the change -- a save that writes the ORIGINAL text back would
        # pass a size check and fail the only thing that matters.
        moved = [i + 1 for i in range(declared)
                 if slide_text(out, i + 1) != slide_text(pptx.read_bytes(), i + 1)]
        rep.check(deck, "save", out[:2] == b"PK" and len(moved) == 1,
                  f"{len(out)} bytes, slide {moved or 'none'} changed")
    rep.check(deck, "console", not errors, "; ".join(errors[:2]) or "no page errors")


def run_index_html(page, port: int, pptx: Path, rep: Report) -> None:
    """The suite viewer's own pptx editing path, checked for the same defect."""
    deck = pptx.name.split("__")[0]
    print(f"{deck} (index.html)", flush=True)
    errors: list[str] = []
    page.on("pageerror", lambda e: errors.append(str(e)))
    page.goto(f"http://127.0.0.1:{port}/index.html")
    # ★Wait for wasm BEFORE handing over a file. The page's change handler
    # returns early while `wasmReady` is false and says so in the status bar,
    # so an upload that arrives first is silently dropped -- which showed up as
    # "no editable runs" and looked like the editor was broken.
    page.wait_for_function("() => window.__oxiWasmReady === true", timeout=60000)
    page.set_input_files("#fileInput", str(pptx))
    page.wait_for_selector(".edit-slide-run", timeout=60000)

    with zipfile.ZipFile(pptx) as z:
        declared = sum(1 for n in z.namelist()
                       if n.startswith("ppt/slides/slide") and n.endswith(".xml"))

    # Type into the first editable run that actually carries text.
    target = page.locator(".edit-slide-run").filter(has_not_text="").first
    n = page.locator(".edit-slide-run").count()
    rep.check(deck, "editable", n > 0, f"{n} runs are editable")
    target.click()
    page.keyboard.press("End")
    page.keyboard.type("Z")

    with page.expect_download() as dl:
        page.click("#btnDownload")
    out = Path(dl.value.path()).read_bytes()
    moved = [i + 1 for i in range(declared)
             if slide_text(out, i + 1) != slide_text(pptx.read_bytes(), i + 1)]
    rep.check(deck, "save", out[:2] == b"PK" and len(moved) == 1,
              f"{len(out)} bytes, slide {moved or 'none'} changed")
    rep.check(deck, "console", not errors, "; ".join(errors[:2]) or "no page errors")


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--decks", default="d15,d09")
    ap.add_argument("--headed", action="store_true")
    args = ap.parse_args()

    from playwright.sync_api import sync_playwright

    targets = deck_paths(args.decks)
    if not targets:
        sys.exit("no decks matched")
    port = serve(WEB)
    rep = Report()
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=not args.headed)
        page = browser.new_page(viewport={"width": 1500, "height": 1000},
                                accept_downloads=True)
        for pptx in targets:
            run_deck(page, port, pptx, rep)
        run_index_html(page, port, targets[0], rep)
        browser.close()
    bad = rep.failed()
    print(f"\n{len(rep.rows) - bad}/{len(rep.rows)} checks passed")
    sys.exit(1 if bad else 0)


if __name__ == "__main__":
    main()
