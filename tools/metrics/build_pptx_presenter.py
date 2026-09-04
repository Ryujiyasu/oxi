# -*- coding: utf-8 -*-
"""Turn Oxi's own render of a deck into a presenter you can run from a browser.

The pptx renderer already puts one PNG on disk per slide. This wraps those PNGs
in a single self-contained page: a start screen, full-screen playback, keyboard
and click navigation, an overview grid, and a countdown sized to the talk. The
images are inlined, so the file works with no network and no server -- open it
from a USB stick if the room's wifi is bad.

    python tools/metrics/build_pptx_presenter.py <png-dir> <out.html>
        [--title "..."] [--subtitle "..."] [--minutes 7]

`<png-dir>` holds `slide_s1.png` .. `slide_sN.png`, which is what
`oxi-pptx-renderer <deck.pptx> <dir>/slide <dpi>` writes.
"""
from __future__ import annotations

import argparse
import base64
import glob
import html
import os
import re


def slide_files(png_dir: str) -> list[str]:
    files = glob.glob(os.path.join(png_dir, "slide_s*.png"))
    if not files:
        raise SystemExit("no slide_sN.png in " + png_dir)
    return sorted(files, key=lambda p: int(re.search(r"_s(\d+)\.png$", p).group(1)))


def data_uri(path: str) -> str:
    with open(path, "rb") as fh:
        return "data:image/png;base64," + base64.b64encode(fh.read()).decode("ascii")


PAGE = """<title>{title_tag}</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@400;500;700&family=Roboto+Mono:wght@400;500&display=swap">
<style>
:root {{
  --ground: #0A0D1F;
  --surface: #141C45;
  --surface-2: #1E2761;
  --ink: #EDEFF7;
  --muted: #8A90AD;
  --accent: #C2511C;
  --accent-soft: #E4713C;
  --warn: #D9A22B;
  --rule: rgba(237, 239, 247, 0.12);
  --sans: "Noto Sans JP", "Yu Gothic UI", "Yu Gothic", "Hiragino Kaku Gothic ProN",
          "Meiryo", system-ui, sans-serif;
  --mono: "Roboto Mono", ui-monospace, "SFMono-Regular", Consolas, monospace;
}}
* {{ box-sizing: border-box; }}
html, body {{ height: 100%; }}
body {{
  margin: 0; background: var(--ground); color: var(--ink);
  font-family: var(--sans); -webkit-font-smoothing: antialiased;
  overflow: hidden;
}}
button {{ font: inherit; color: inherit; }}
:focus-visible {{ outline: 2px solid var(--accent-soft); outline-offset: 3px; }}

/* ---------- start screen ---------- */
#start {{
  position: fixed; inset: 0; display: grid;
  grid-template-rows: 1fr auto; gap: 0;
  background:
    radial-gradient(1200px 600px at 82% -10%, rgba(30, 39, 97, 0.85), transparent 60%),
    radial-gradient(700px 500px at 8% 108%, rgba(194, 81, 28, 0.18), transparent 62%),
    var(--ground);
}}
.lead {{
  display: flex; flex-direction: column; justify-content: center;
  padding: clamp(28px, 5vw, 72px); max-width: 1100px;
}}
.mark {{
  font-weight: 700; letter-spacing: 0.02em; color: var(--accent-soft);
  font-size: clamp(20px, 2.4vw, 28px); margin: 0 0 clamp(14px, 2vw, 26px);
}}
h1 {{
  margin: 0; font-weight: 700; text-wrap: balance;
  font-size: clamp(28px, 4.4vw, 56px); line-height: 1.28; letter-spacing: -0.01em;
}}
.sub {{
  margin: clamp(12px, 1.6vw, 20px) 0 0; color: var(--muted);
  font-size: clamp(14px, 1.35vw, 17px); line-height: 1.75; max-width: 62ch;
}}
.actions {{ display: flex; flex-wrap: wrap; align-items: center; gap: 18px;
            margin-top: clamp(22px, 3vw, 38px); }}
.go {{
  border: 0; border-radius: 4px; cursor: pointer;
  background: var(--accent); color: #fff; font-weight: 700;
  font-size: clamp(15px, 1.5vw, 18px);
  padding: 0.95em 2.1em; letter-spacing: 0.04em;
  transition: background 140ms ease, transform 140ms ease;
}}
.go:hover {{ background: var(--accent-soft); }}
.go:active {{ transform: translateY(1px); }}
.keys {{ color: var(--muted); font-size: 13px; line-height: 1.9; }}
kbd {{
  font-family: var(--mono); font-size: 12px; color: var(--ink);
  border: 1px solid var(--rule); border-bottom-width: 2px; border-radius: 3px;
  padding: 1px 6px; margin: 0 2px; background: rgba(237, 239, 247, 0.05);
}}

/* ---------- filmstrip ---------- */
.strip {{
  display: flex; gap: 10px; padding: 14px clamp(28px, 5vw, 72px) 22px;
  overflow-x: auto; overflow-y: hidden; border-top: 1px solid var(--rule);
  background: rgba(10, 13, 31, 0.6); scrollbar-width: thin;
}}
.strip button {{
  flex: 0 0 auto; padding: 0; border: 1px solid var(--rule); border-radius: 3px;
  background: none; cursor: pointer; line-height: 0; position: relative;
  transition: border-color 120ms ease, transform 120ms ease;
}}
.strip button:hover {{ border-color: var(--accent-soft); transform: translateY(-2px); }}
.strip img {{ width: 152px; height: auto; display: block; border-radius: 2px; }}
.strip .n {{
  position: absolute; left: 5px; bottom: 4px; font-family: var(--mono);
  font-size: 10px; color: var(--ink); background: rgba(10, 13, 31, 0.75);
  padding: 1px 5px; border-radius: 2px; line-height: 1.6;
}}

/* ---------- player ---------- */
#player {{ position: fixed; inset: 0; display: none; background: #000; }}
#player.on {{ display: block; }}
#stage {{ position: absolute; inset: 0; display: grid; place-items: center; }}
#stage img {{ max-width: 100%; max-height: 100%; display: block; }}
.chrome {{
  position: fixed; left: 0; right: 0; bottom: 0; padding: 14px 20px 16px;
  display: flex; align-items: center; gap: 16px;
  background: linear-gradient(to top, rgba(0, 0, 0, 0.72), transparent);
  transition: opacity 260ms ease; opacity: 1;
}}
#player.idle .chrome {{ opacity: 0; }}
.clock {{
  font-family: var(--mono); font-size: 15px; letter-spacing: 0.04em;
  font-variant-numeric: tabular-nums; color: var(--ink);
  border: 1px solid var(--rule); border-radius: 3px; padding: 4px 10px;
  background: rgba(0, 0, 0, 0.45);
}}
.clock.warn {{ color: var(--warn); border-color: rgba(217, 162, 43, 0.5); }}
.clock.over {{ color: #fff; background: var(--accent); border-color: transparent; }}
.clock.off {{ opacity: 0.28; }}
.count {{
  margin-left: auto; font-family: var(--mono); font-size: 14px;
  font-variant-numeric: tabular-nums; color: var(--muted);
}}
.count b {{ color: var(--ink); font-weight: 500; }}
#rail {{
  position: fixed; left: 0; bottom: 0; height: 3px; background: var(--accent);
  transition: width 180ms ease; z-index: 3;
}}

/* ---------- overview ---------- */
#grid {{
  position: fixed; inset: 0; display: none; overflow: auto; z-index: 4;
  background: rgba(10, 13, 31, 0.97); padding: clamp(20px, 3vw, 44px);
}}
#grid.on {{ display: block; }}
.gridhead {{ display: flex; align-items: baseline; gap: 14px; margin-bottom: 18px; }}
.gridhead h2 {{ margin: 0; font-size: 16px; font-weight: 700; letter-spacing: 0.06em; }}
.gridhead span {{ color: var(--muted); font-size: 13px; }}
.tiles {{ display: grid; gap: 14px; grid-template-columns: repeat(auto-fill, minmax(230px, 1fr)); }}
.tiles button {{
  padding: 0; border: 1px solid var(--rule); border-radius: 3px; background: none;
  cursor: pointer; line-height: 0; position: relative;
}}
.tiles button[aria-current="true"] {{ border-color: var(--accent); }}
.tiles img {{ width: 100%; height: auto; display: block; border-radius: 2px; }}
.tiles .n {{
  position: absolute; left: 6px; bottom: 6px; font-family: var(--mono); font-size: 11px;
  background: rgba(10, 13, 31, 0.78); padding: 2px 6px; border-radius: 2px; color: var(--ink);
}}
@media (prefers-reduced-motion: reduce) {{
  * {{ transition: none !important; }}
}}
</style>

<main id="start">
  <div class="lead">
    <p class="mark">{mark}</p>
    <h1>{title}</h1>
    <p class="sub">{subtitle}</p>
    <div class="actions">
      <button class="go" id="go" type="button">プレゼンを開始</button>
      <p class="keys">
        <kbd>&rarr;</kbd> <kbd>Space</kbd> 次へ &nbsp; <kbd>&larr;</kbd> 前へ &nbsp;
        <kbd>O</kbd> 一覧 &nbsp; <kbd>T</kbd> タイマー &nbsp; <kbd>R</kbd> 時間をリセット &nbsp;
        <kbd>Esc</kbd> 終了
      </p>
    </div>
  </div>
  <div class="strip" id="strip"></div>
</main>

<div id="player" aria-live="polite">
  <div id="stage"><img id="shot" alt=""></div>
  <div class="chrome">
    <span class="clock" id="clock">{clock0}</span>
    <span class="count"><b id="cur">1</b> / {n}</span>
  </div>
  <div id="rail"></div>
</div>

<div id="grid">
  <div class="gridhead">
    <h2>スライド一覧</h2><span>クリックでそのページへ &nbsp;/&nbsp; Esc で戻る</span>
  </div>
  <div class="tiles" id="tiles"></div>
</div>

<script>
const SLIDES = {slides_js};
const TOTAL = SLIDES.length;
const LIMIT = {seconds};
let i = 0, started = 0, timerOn = true, idleTimer = 0;

const $ = (id) => document.getElementById(id);
const shot = $("shot"), player = $("player"), gridEl = $("grid");

function show(n) {{
  i = Math.max(0, Math.min(TOTAL - 1, n));
  shot.src = SLIDES[i];
  shot.alt = "スライド " + (i + 1);
  $("cur").textContent = i + 1;
  $("rail").style.width = ((i + 1) / TOTAL * 100) + "%";
  for (const b of gridEl.querySelectorAll("button")) {{
    b.setAttribute("aria-current", String(Number(b.dataset.i) === i));
  }}
}}

function begin(n) {{
  show(n || 0);
  player.classList.add("on");
  $("start").style.display = "none";
  if (!started) started = Date.now();
  const el = document.documentElement;
  if (el.requestFullscreen) el.requestFullscreen().catch(() => {{}});
  wake();
}}

function finish() {{
  player.classList.remove("on");
  gridEl.classList.remove("on");
  $("start").style.display = "grid";
  if (document.fullscreenElement) document.exitFullscreen().catch(() => {{}});
}}

function pad(v) {{ return String(v).padStart(2, "0"); }}

function tick() {{
  const c = $("clock");
  c.classList.toggle("off", !timerOn);
  if (!started) {{ c.textContent = "{clock0}"; return; }}
  const left = LIMIT - Math.floor((Date.now() - started) / 1000);
  const over = left < 0, a = Math.abs(left);
  c.textContent = (over ? "+" : "") + Math.floor(a / 60) + ":" + pad(a % 60);
  c.classList.toggle("warn", !over && left <= 60);
  c.classList.toggle("over", over);
}}
setInterval(tick, 250);

function wake() {{
  player.classList.remove("idle");
  clearTimeout(idleTimer);
  idleTimer = setTimeout(() => player.classList.add("idle"), 2600);
}}

document.addEventListener("keydown", (e) => {{
  const playing = player.classList.contains("on");
  const grid = gridEl.classList.contains("on");
  if (e.key === "Escape") {{
    if (grid) {{ gridEl.classList.remove("on"); e.preventDefault(); }}
    else if (playing) {{ finish(); e.preventDefault(); }}
    return;
  }}
  if (!playing) {{
    if (e.key === "Enter" || e.key === " ") {{ begin(0); e.preventDefault(); }}
    return;
  }}
  switch (e.key) {{
    case "ArrowRight": case " ": case "PageDown": case "n": case "N":
      show(i + 1); wake(); e.preventDefault(); break;
    case "ArrowLeft": case "PageUp": case "p": case "P":
      show(i - 1); wake(); e.preventDefault(); break;
    case "Home": show(0); wake(); e.preventDefault(); break;
    case "End": show(TOTAL - 1); wake(); e.preventDefault(); break;
    case "o": case "O": case "g": case "G":
      gridEl.classList.toggle("on"); e.preventDefault(); break;
    case "t": case "T": timerOn = !timerOn; tick(); e.preventDefault(); break;
    case "r": case "R": started = Date.now(); tick(); e.preventDefault(); break;
    case "f": case "F":
      if (document.fullscreenElement) document.exitFullscreen();
      else document.documentElement.requestFullscreen().catch(() => {{}});
      e.preventDefault(); break;
  }}
}});

$("stage").addEventListener("click", (e) => {{
  show(e.clientX < window.innerWidth * 0.32 ? i - 1 : i + 1);
  wake();
}});
player.addEventListener("mousemove", wake);

let touchX = null;
player.addEventListener("touchstart", (e) => {{ touchX = e.touches[0].clientX; }}, {{ passive: true }});
player.addEventListener("touchend", (e) => {{
  if (touchX === null) return;
  const dx = e.changedTouches[0].clientX - touchX;
  if (Math.abs(dx) > 40) show(dx < 0 ? i + 1 : i - 1);
  touchX = null;
}});

$("go").addEventListener("click", () => begin(0));

const strip = $("strip"), tiles = $("tiles");
SLIDES.forEach((src, k) => {{
  const b = document.createElement("button");
  b.type = "button"; b.dataset.i = k;
  b.innerHTML = '<img loading="lazy" alt="スライド ' + (k + 1) + '"><span class="n">' + (k + 1) + "</span>";
  b.querySelector("img").src = src;
  b.addEventListener("click", () => begin(k));
  strip.appendChild(b);

  const t = b.cloneNode(true);
  t.addEventListener("click", () => {{ gridEl.classList.remove("on"); show(k); }});
  tiles.appendChild(t);
}});

show(0);
tick();
</script>
"""


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("png_dir")
    ap.add_argument("out")
    ap.add_argument("--title", default="")
    ap.add_argument("--subtitle", default="")
    ap.add_argument("--mark", default="Oxi")
    ap.add_argument("--tab-title", default="")
    ap.add_argument("--minutes", type=float, default=7.0)
    args = ap.parse_args()

    files = slide_files(args.png_dir)
    uris = [data_uri(f) for f in files]
    seconds = int(args.minutes * 60)
    page = PAGE.format(
        title_tag=html.escape(args.tab_title or args.title or "Presenter"),
        mark=html.escape(args.mark),
        title=html.escape(args.title).replace("\n", "<br>"),
        subtitle=html.escape(args.subtitle),
        n=len(files),
        seconds=seconds,
        clock0="%d:%02d" % (seconds // 60, seconds % 60),
        slides_js="[\n" + ",\n".join('"%s"' % u for u in uris) + "\n]",
    )
    with open(args.out, "w", encoding="utf-8") as fh:
        fh.write(page)
    print("wrote %s  (%d slides, %.1f MB)"
          % (args.out, len(files), os.path.getsize(args.out) / 1048576))


if __name__ == "__main__":
    main()
