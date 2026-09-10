// Records the three editors doing their job, and writes one GIF each.
//
// A README can say "opens a .docx in the browser with no server" for as long as
// it likes; a moving picture of a real file opening is the only version of that
// sentence a reader can check at a glance. Chrome drives the same pages the
// desktop app ships, so what the GIF shows is what the app does — there is no
// mock-up anywhere in here.
//
//   node tools/make-demo-gifs.mjs            # all three
//   node tools/make-demo-gifs.mjs cells      # just one
//
// ffmpeg does the encoding: a palette is generated from the frames themselves,
// because GIF's 256 colours cut through a document's greys badly otherwise.
import { createRequire } from 'node:module';
import { createServer } from 'node:http';
import { spawnSync } from 'node:child_process';
import { existsSync, mkdirSync, rmSync, readFileSync } from 'node:fs';
import { dirname, extname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const repo = join(here, '..');
const out = join(repo, 'docs', 'img');
const work = join(repo, 'target', 'demo-frames');

const require = createRequire(join(repo, 'tools', 'oxi-chromium-renderer', 'package.json'));
const puppeteer = require('puppeteer');

const BROWSERS = [
  process.env.OXI_BROWSER,
  'C:/Program Files/Google/Chrome/Application/chrome.exe',
  'C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe',
  '/usr/bin/google-chrome',
  '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
].filter(Boolean);

const WIDE = 1160;
const TALL = 720;
const FPS = 10;

const TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.mjs': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.wasm': 'application/wasm',
  '.json': 'application/json',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
};

/**
 * Serves the assembled desktop bundle the way a web server would, because wasm
 * will not load off file://.
 *
 * The bundle rather than `web/` on purpose: it is the exact set of files the
 * installer ships, flattened and with the analytics stripped, so a reader
 * watching the GIF is watching the application rather than a development page
 * that happens to resemble it.
 */
function serve() {
  const root = join(repo, 'dist-desktop', 'web');
  const server = createServer((req, res) => {
    const path = join(root, decodeURIComponent(req.url.split('?')[0]));
    if (!path.startsWith(root) || !existsSync(path)) {
      res.writeHead(404).end('no');
      return;
    }
    res.writeHead(200, { 'content-type': TYPES[extname(path)] || 'application/octet-stream' });
    res.end(readFileSync(path));
  });
  return new Promise((ok) => {
    server.listen(0, '127.0.0.1', () => ok({ server, port: server.address().port }));
  });
}

/** One recording: frames go out numbered, and the caption strip is part of the frame. */
class Take {
  constructor(page, folder) {
    this.page = page;
    this.folder = folder;
    this.at = 0;
    rmSync(folder, { recursive: true, force: true });
    mkdirSync(folder, { recursive: true });
  }

  async shot() {
    await this.page.screenshot({
      path: join(this.folder, `${String(this.at).padStart(4, '0')}.png`),
    });
    this.at += 1;
  }

  /** Frames for as long as the reader needs to take the picture in. */
  async hold(ms) {
    const frames = Math.max(1, Math.round((ms / 1000) * FPS));
    for (let n = 0; n < frames; n += 1) {
      await this.shot();
      await new Promise((ok) => setTimeout(ok, 1000 / FPS));
    }
  }

  async say(text) {
    await this.page.evaluate((words) => {
      let bar = document.getElementById('__cap');
      if (!bar) {
        bar = document.createElement('div');
        bar.id = '__cap';
        bar.style.cssText = 'position:fixed;left:0;right:0;bottom:0;z-index:2147483647;'
          + 'background:rgba(16,17,21,.9);color:#fff;padding:11px 18px;'
          + 'font:15px/1.45 "Segoe UI",-apple-system,system-ui,sans-serif;'
          + 'letter-spacing:.01em;pointer-events:none';
        document.body.appendChild(bar);
      }
      bar.textContent = words;
    }, text);
  }

  /** Typed a letter at a time, so the reader sees it being typed. */
  async type(text, gap = 70) {
    for (const letter of text) {
      await this.page.keyboard.type(letter);
      await new Promise((ok) => setTimeout(ok, gap));
      await this.shot();
    }
  }
}

/** True when the selector turned up; a beat that cannot run is skipped, not fatal. */
async function there(page, selector, ms = 20000) {
  try {
    await page.waitForSelector(selector, { timeout: ms });
    return true;
  } catch {
    return false;
  }
}

/**
 * Each page binds its file input only after `await init()` has finished loading
 * a 28MB wasm module, so the input existing proves nothing: a file dropped in
 * before that lands on an element with no handler and simply vanishes. What
 * marks the engine as up is the handler arriving, which is watched for two ways
 * because the pages bind it two ways.
 */
const WATCH = `
  window.__oxiReady = false;
  const real = EventTarget.prototype.addEventListener;
  EventTarget.prototype.addEventListener = function (type, fn, opts) {
    if (type === 'change' && this && this.type === 'file') window.__oxiReady = true;
    return real.call(this, type, fn, opts);
  };
`;

/** The pages remember a language; a README in English should record in English. */
async function english(page, control) {
  await page.evaluate((id) => {
    const at = document.querySelector(id);
    if (!at) return;
    if (at.tagName === 'SELECT') {
      at.value = 'en';
      at.dispatchEvent(new Event('change', { bubbles: true }));
    } else {
      at.click();
    }
  }, control);
}

async function docs(page, take) {
  await english(page, '#langSel');
  await take.say('Oxidocs — a .docx opens in the browser. No server, nothing uploaded.');
  await take.hold(1500);
  await (await page.$('#fileIn')).uploadFile(join(repo, 'docs', 'sample.docx'));
  await page.waitForFunction('document.querySelectorAll("canvas").length > 0', { timeout: 60000 })
    .catch(() => {});
  // Opening leaves the caret at the end of the document, and the view with it.
  await page.evaluate(() => {
    const area = document.getElementById('editArea');
    if (area) area.scrollTop = 0;
  });
  await take.hold(2600);
  await take.say('Laid out by the engine that is scored page by page against Word.');
  await take.hold(1700);
  await page.mouse.move(WIDE / 2, TALL / 2);
  for (let n = 0; n < 22; n += 1) {
    await page.mouse.wheel({ deltaY: 70 });
    await take.shot();
  }
  await take.hold(1200);

  // The same thing the slide clip was missing: a still cannot show that the
  // page is editable, so the clip has to do it.
  await page.evaluate(() => {
    const area = document.getElementById('editArea');
    if (area) area.scrollTop = 0;
  });
  await take.hold(1200);
  await take.say('The page is editable in place: pick a word and type over it.');
  // A body line rather than the title. Bold on a heading that is already bold
  // shows a viewer nothing; replacing a word shows them everything.
  const lines = await page.$$('#editor p');
  let box = null;
  for (const line of lines) {
    const size = await line.boundingBox();
    if (size && size.width > 380 && size.height > 12 && size.height < 40) {
      box = size;
      break;
    }
  }
  if (box) {
    await page.mouse.click(box.x + 60, box.y + box.height / 2, { clickCount: 2 });
    await take.hold(1000);
    await take.type('書き換えた', 90);
    await take.hold(1400);
  }
  await take.say('Saving patches only the XML that changed — the rest of the file is untouched.');
  await take.hold(2400);
}

async function cells(page, take) {
  await english(page, '#langEn');
  await take.say('Oxicells — a .xlsx, its formulas, and its macros.');
  await take.hold(1500);
  await (await page.$('#pick')).uploadFile(join(repo, 'web', 'samples', 'quarterly-sales.xlsx'));
  await take.hold(2800);
  const box = await (await page.$('#paper')).boundingBox();
  // F4 carries a formula, so the bar shows the formula and not the number.
  await take.say('Click a cell and the bar shows what is really in it.');
  await page.mouse.click(box.x + 601, box.y + 111);
  await take.hold(2000);
  // An empty cell well below the table, and the formula goes in through the bar
  // where a viewer can watch it being written.
  await take.say('Write a formula and the engine recalculates the sheet.');
  await page.mouse.click(box.x + 197, box.y + 286);
  await take.hold(500);
  await page.click('#formula');
  await take.type('=SUM(B4:B8)');
  await page.keyboard.press('Enter');
  await take.hold(2400);
  if (await there(page, '#macros', 3000)) {
    await take.say("The workbook's own VBA runs here, checked against Excel itself.");
    await page.click('#macros');
    await take.hold(2600);
    await page.click('#macroClose').catch(() => {});
  }
  await take.hold(900);
}

async function slides(page, take) {
  await english(page, '#langEn');
  await take.say('Oxislides — a .pptx, slide by slide.');
  await take.hold(1500);
  await (await page.$('#fileIn')).uploadFile(join(repo, 'docs', 'showcase.pptx'));
  await take.hold(2800);
  await take.say('Every slide is drawn by the engine, not shown as a picture of one.');
  const thumbs = await page.$$('#rail > *');
  for (const thumb of thumbs.slice(1, 3)) {
    await thumb.click().catch(() => {});
    await take.hold(1300);
  }
  if (thumbs.length) await thumbs[0].click().catch(() => {});
  await take.hold(1000);

  // The part a viewer cannot guess from a still: the deck is editable, and the
  // text is edited in place on the slide rather than in a side panel.
  if (await there(page, '#btnEdit', 3000)) {
    await take.say('Turn on editing and the text on the slide takes a caret.');
    await page.click('#btnEdit');
    await take.hold(1600);
    const runs = await page.$$('.slide-shape span[contenteditable]');
    let box = null;
    for (const run of runs) {
      const size = await run.boundingBox();
      if (size && size.width > 120 && size.height > 14) { box = size; break; }
    }
    if (box) {
      const x = box.x + Math.min(box.width / 2, 200);
      const y = box.y + box.height / 2;
      // Three clicks take the whole line, the way they do anywhere else.
      await page.mouse.click(x, y, { clickCount: 3 });
      await take.hold(900);
      await take.say('Type straight onto the slide; the thumbnail follows.');
      await take.type('Edited here, on the slide', 55);
      await take.hold(1200);
      await page.mouse.click(x, y, { clickCount: 3 });
      await take.hold(600);
      if (await there(page, '#alnC', 2000)) {
        // Centring moves the words visibly; bold on a title that is already
        // bold changes nothing a viewer can see.
        await take.say('Weight, alignment, size and colour apply to the selection.');
        await page.click('#alnC');
        await take.hold(1500);
        await page.click('#alnR');
        await take.hold(1500);
      }
    }
    await take.say('Saving writes a .pptx, patching only what changed.');
    await take.hold(2200);
  }
}

const CLIPS = {
  docs: { page: 'docs.html', ready: 'window.__oxiReady', play: docs },
  // This one assigns `.onchange` rather than adding a listener.
  cells: {
    page: 'sheets.html',
    ready: "!!(document.getElementById('pick') || {}).onchange",
    play: cells,
  },
  // This one, too, assigns `.onchange`.
  slides: {
    page: 'slides.html',
    ready: "!!(document.getElementById('fileIn') || {}).onchange",
    play: slides,
  },
};

/** Frames to GIF, with a palette cut from the frames so the greys survive. */
function encode(folder, target) {
  const filter = `fps=${FPS},scale=920:-1:flags=lanczos,split[a][b];`
    + '[a]palettegen=max_colors=160:stats_mode=diff[p];'
    + '[b][p]paletteuse=dither=bayer:bayer_scale=3';
  const run = spawnSync('ffmpeg', [
    '-y', '-loglevel', 'error', '-framerate', String(FPS),
    '-i', join(folder, '%04d.png'), '-vf', filter, '-loop', '0', target,
  ], { encoding: 'utf8' });
  if (run.status !== 0) {
    throw new Error(`ffmpeg would not encode ${target}: ${run.stderr || run.error}`);
  }
}

const wanted = process.argv.slice(2).filter((word) => CLIPS[word]);
const picked = wanted.length ? wanted : Object.keys(CLIPS);

const executablePath = BROWSERS.find((path) => existsSync(path));
if (!executablePath) {
  throw new Error(`no browser to record with — set OXI_BROWSER to one of: ${BROWSERS.join(', ')}`);
}
mkdirSync(out, { recursive: true });

const { server, port } = await serve();
const browser = await puppeteer.launch({
  headless: 'new',
  executablePath,
  args: ['--no-sandbox', '--hide-scrollbars', '--force-device-scale-factor=1'],
  defaultViewport: { width: WIDE, height: TALL },
});

try {
  for (const name of picked) {
    const clip = CLIPS[name];
    const page = await browser.newPage();
    page.on('pageerror', (error) => console.log(`  ${name}: ${error.message}`));
    page.on('response', (res) => {
      if (res.status() >= 400) console.log(`  ${name}: ${res.status()} ${res.url()}`);
    });
    page.on('console', (line) => {
      if (line.type() === 'error') console.log(`  ${name}: ${line.text()}`);
    });
    await page.evaluateOnNewDocument(WATCH);
    await page.goto(`http://127.0.0.1:${port}/${clip.page}`, { waitUntil: 'domcontentloaded' });
    try {
      await page.waitForFunction(clip.ready, { timeout: 120000 });
    } catch {
      console.log(`${name}: the engine never came up — skipped`);
      await page.close();
      continue;
    }
    const folder = join(work, name);
    const take = new Take(page, folder);
    await clip.play(page, take);
    await page.close();
    const target = join(out, `oxi-${name}.gif`);
    encode(folder, target);
    const size = (readFileSync(target).length / 1048576).toFixed(1);
    console.log(`${name}: ${take.at} frames -> ${target} (${size} MB)`);
  }
} finally {
  await browser.close();
  server.close();
}
