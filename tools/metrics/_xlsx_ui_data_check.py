# -*- coding: utf-8 -*-
"""Drive docs/sheets.html in Chromium: insert a pivot table from a selection,
refresh it, run Solver, and save -- then read the saved files back.

The page and a wasm build are served from a scratch copy of docs/ (the site
engine is built by CI, not committed):

    python tools/metrics/_xlsx_ui_data_check.py <docs-dir> <wasm-pkg-dir>
"""
import json, os, shutil, socket, subprocess, sys, tempfile, threading, time
from pathlib import Path

import openpyxl
from playwright.sync_api import sync_playwright

REPO = Path(__file__).resolve().parents[2]
RENDERER = Path(os.environ.get("OXI_XLSX_RENDERER", REPO / "tools/oxi-xlsx-renderer/target/release/oxi-xlsx-renderer.exe"))


def serve(root: Path) -> int:
    import http.server, functools
    sock = socket.socket(); sock.bind(("127.0.0.1", 0)); port = sock.getsockname()[1]; sock.close()
    handler = functools.partial(http.server.SimpleHTTPRequestHandler, directory=str(root))
    http.server.SimpleHTTPRequestHandler.extensions_map[".wasm"] = "application/wasm"
    server = http.server.ThreadingHTTPServer(("127.0.0.1", port), handler)
    threading.Thread(target=server.serve_forever, daemon=True).start()
    return port


def data_book(path: Path):
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    import _xlsx_pivot_gen as gen
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Data"
    ws.append(["Region", "Product", "Month", "Sales", "Qty"])
    for r in gen.data_rows():
        ws.append(r)
    wb.save(path)


def main():
    docs, pkg = Path(sys.argv[1]), Path(sys.argv[2])
    site = Path(tempfile.mkdtemp(prefix="oxi_site_"))
    shutil.copytree(docs, site, dirs_exist_ok=True)
    # The Pages job also copies the shared scripts from web/.
    for js in (docs.parent / "web").glob("*.js"):
        if not (site / js.name).exists():
            shutil.copy(js, site / js.name)
    for f in ("oxidocs_wasm.js", "oxidocs_wasm_bg.wasm"):
        shutil.copy(pkg / f, site / f)
    port = serve(site)
    work = Path(tempfile.mkdtemp(prefix="oxi_ui_"))
    data = work / "data.xlsx"; data_book(data)
    model = REPO / "pipeline_data/xlsx_solver/s01_lp_max.xlsx"
    ok = True
    with sync_playwright() as pw:
        browser = pw.chromium.launch()
        page = browser.new_page(accept_downloads=True)
        page.on("pageerror", lambda e: print("  page error:", e, getattr(e, "stack", "")[:600]))
        page.goto(f"http://127.0.0.1:{port}/sheets.html")
        page.wait_for_function("!document.getElementById('btnSample').disabled", timeout=120000)
        # The truth is a Japanese Excel: captions in Japanese.
        page.click("#langJa")
        # --- pivot
        page.set_input_files("#fileIn", str(data))
        page.wait_for_function("!document.getElementById('btnPivot').disabled", timeout=60000)
        page.fill("#cellRef", "A1:E61"); page.press("#cellRef", "Enter")
        page.click("#btnPivot")
        page.select_option("#pvRows", "0"); page.select_option("#pvRows2", "1")
        page.select_option("#pvVal", "3"); page.select_option("#pvFn", "sum")
        page.click("#pivotDlg button[value=ok]")
        page.wait_for_timeout(1500)
        print("  pivot message:", page.inner_text("#msg") if page.query_selector("#msg") else "?")
        refresh_visible = page.is_visible("#btnPivotRefresh")
        with page.expect_download() as dl:
            page.click("#btnSave")
        saved = work / "pivot_saved.xlsx"; dl.value.save_as(saved)
        run = subprocess.run([str(RENDERER), str(saved), "--pivot"], capture_output=True, text=True, encoding="utf-8")
        got = json.loads(run.stdout or "[]")
        truth = json.loads((REPO / "pipeline_data/xlsx_pivot/p02_region_product_sum.truth.json").read_text(encoding="cp932"))
        norm = lambda v: int(v) if isinstance(v, float) and v.is_integer() else v
        same = bool(got) and [[norm(v) for v in r] for r in got[0]["grid"]] == [[norm(v) for v in r] for r in truth["grid"]]
        # the cells the page wrote, read back from the saved file
        ws = openpyxl.load_workbook(saved, data_only=True)["Data"]
        cell_grid = [[ws.cell(1 + i, 7 + j).value for j in range(2)] for i in range(len(truth["grid"]))]
        cells_same = [[norm(v) for v in r] for r in cell_grid] == [[norm(v) for v in r] for r in truth["grid"]]
        print(f"  pivot: refresh button {'shown' if refresh_visible else 'missing'}, definition {'matches' if same else 'DIFFERS'} Excel, saved cells {'match' if cells_same else 'DIFFER'}")
        ok &= refresh_visible and same and cells_same
        # --- solver
        page.set_input_files("#fileIn", str(model))
        page.wait_for_function("!document.getElementById('btnSolver').hidden", timeout=60000)
        page.click("#btnSolver")
        page.wait_for_timeout(1500)
        message = page.inner_text("#msg") if page.query_selector("#msg") else ""
        with page.expect_download() as dl:
            page.click("#btnSave")
        solved = work / "solved.xlsx"; dl.value.save_as(solved)
        ws = openpyxl.load_workbook(solved, data_only=True).worksheets[0]
        values = [ws["A1"].value, ws["A2"].value]
        good = values == [4, 0] or values == [4.0, 0.0]
        print(f"  solver: message {message!r}, variables {values} {'OK' if good else 'DIFF'}")
        ok &= good
        # --- add-in through oxi-office.js
        page.set_input_files("#fileIn", str(data))
        page.wait_for_function("!document.getElementById('btnAddin').disabled", timeout=60000)
        page.click("#langJa")
        page.on("dialog", lambda d: d.accept(f"http://127.0.0.1:{port}/addin-sample.html"))
        page.click("#btnAddin")
        frame = page.frame_locator("#addinFrame")
        frame.locator("#fill").click()
        page.wait_for_timeout(1500)
        frame.locator("#read").click()
        page.wait_for_timeout(1500)
        said = frame.locator("#out").inner_text()
        good = said.strip() == "Data!B5 = 245"
        print(f"  add-in: {said!r} {'OK' if good else 'DIFF'}")
        ok &= good
        # --- .xlam add-in: its functions recalculate the sheet, its macro runs
        addin_dir = REPO / "tests/fixtures/xlsx_addin"
        page.set_input_files("#fileIn", str(addin_dir / "uses_addin.xlsx"))
        page.wait_for_function("!document.getElementById('btnVba').disabled", timeout=60000)
        page.click("#btnVba")
        page.wait_for_selector("#vbaDialog[open]")
        page.set_input_files("#vbaAddinFile", str(addin_dir / "TaxTools.xlam"))
        page.wait_for_function("/再計算しました|Recalculated/.test(document.getElementById('vbaOutput').textContent)", timeout=120000)
        page.fill("#vbaProcedure", "StampTotals")
        page.click("#vbaRun")
        page.wait_for_function("!/Running|実行中/.test(document.getElementById('vbaOutput').textContent) && document.getElementById('vbaOutput').textContent.length > 0", timeout=120000)
        page.wait_for_timeout(1000)
        page.click("#vbaClose")
        with page.expect_download() as dl:
            page.click("#btnSave")
        stamped = work / "stamped.xlsx"; dl.value.save_as(stamped)
        ws = openpyxl.load_workbook(stamped, data_only=True).worksheets[0]
        got = {k: ws[k].value for k in ("D1", "D2", "C1", "C2")}
        want = {"D1": 330, "D2": 216, "C1": "Stamped by add-in", "C2": 55}
        good = all((got[k] == v) or (isinstance(v, (int, float)) and isinstance(got[k], (int, float)) and abs(got[k] - v) < 1e-9) for k, v in want.items())
        print(f"  xlam: {got} {'OK' if good else 'DIFF'}")
        ok &= good
        browser.close()
    print("UI data features:", "OK" if ok else "FAILED")


if __name__ == "__main__":
    main()
