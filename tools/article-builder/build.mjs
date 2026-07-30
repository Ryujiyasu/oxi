import { readFile, writeFile } from "node:fs/promises";
import { resolve } from "node:path";
import { marked } from "marked";

const root = resolve(import.meta.dirname, "../..");
const articles = [
  {
    source: "docs/articles/word-pagination-gdi-rounding.md",
    output: "docs/articles/word-pagination-gdi-rounding.html",
    lang: "en",
    description: "How Oxi reproduced Microsoft Word pagination by measuring and implementing GDI integer rounding in Rust.",
    alternate: "./word-pagination-gdi-rounding.ja.html",
    alternateLabel: "日本語",
    home: "../",
  },
  {
    source: "docs/articles/word-pagination-gdi-rounding.ja.md",
    output: "docs/articles/word-pagination-gdi-rounding.ja.html",
    lang: "ja",
    description: "Microsoft Wordの改ページ位置を決めるGDI整数丸めを、実測からRustで再現した方法。",
    alternate: "./word-pagination-gdi-rounding.html",
    alternateLabel: "English",
    home: "../ja/",
  },
];

const style = `
:root{--ink:#182227;--muted:#59666c;--paper:#fbfaf7;--card:#fff;--line:#d9d5cd;--brand:#264653;--accent:#c84c3a;--code:#f2efe9}
*{box-sizing:border-box}html{scroll-behavior:smooth}body{margin:0;background:var(--paper);color:var(--ink);font:17px/1.78 system-ui,-apple-system,"Segoe UI","Noto Sans JP",sans-serif}
a{color:#17627c;text-underline-offset:3px}.site{background:var(--brand);color:#fff}.site-inner{max-width:900px;margin:auto;padding:12px 24px;display:flex;align-items:center;gap:14px}.brand{color:#fff;text-decoration:none;font-weight:800;letter-spacing:.04em}.site nav{margin-left:auto;display:flex;gap:18px}.site nav a{color:#fff;text-decoration:none;font-size:14px}
article{max-width:820px;margin:auto;padding:64px 24px 80px}h1{font-size:clamp(2.15rem,6vw,4.25rem);line-height:1.06;letter-spacing:-.045em;margin:0 0 28px;color:#102d39}h2{font-size:1.7rem;line-height:1.25;margin:3.2rem 0 1rem;color:#173d4b}h3{font-size:1.2rem;margin:2rem 0 .55rem}p{margin:0 0 1.15rem}ul,ol{padding-left:1.45rem;margin:0 0 1.35rem}li{margin:.35rem 0}pre{overflow:auto;background:var(--code);border:1px solid var(--line);border-radius:9px;padding:16px 18px;line-height:1.5;margin:1.3rem 0 1.6rem}code{font-family:"Cascadia Code","SFMono-Regular",Consolas,monospace;font-size:.9em}p code,li code{background:var(--code);padding:.12em .34em;border-radius:4px}table{width:100%;border-collapse:collapse;margin:1.3rem 0 1.7rem;background:var(--card);font-variant-numeric:tabular-nums}th,td{border:1px solid var(--line);padding:9px 12px;text-align:left}th{background:#eaf0f1}.lede{font-size:1.16rem;color:var(--muted);border-left:4px solid var(--accent);padding-left:18px;margin-bottom:2rem}.meta{font-size:.85rem;color:var(--muted);margin-bottom:18px;text-transform:uppercase;letter-spacing:.08em}.end{margin-top:4rem;padding-top:1.5rem;border-top:1px solid var(--line);color:var(--muted)}
@media(max-width:600px){body{font-size:16px}.site-inner{padding:10px 16px}article{padding:42px 18px 60px}h1{font-size:2.35rem}.site nav a:first-child{display:none}th,td{padding:7px 8px;font-size:.88rem}}
`;

for (const item of articles) {
  const markdown = await readFile(resolve(root, item.source), "utf8");
  const title = markdown.match(/^# (.+)$/m)?.[1] ?? "Oxi";
  const body = await marked.parse(markdown);
  const firstParagraph = body.replace(/^<h1[^>]*>.*?<\/h1>\s*/s, "").match(/^<p>(.*?)<\/p>/s)?.[1] ?? "";
  const articleBody = body
    .replace(/^<h1[^>]*>.*?<\/h1>\s*/s, "")
    .replace(/^<p>.*?<\/p>/s, `<p class="lede">${firstParagraph}</p>`);
  const canonicalName = item.output.split("/").at(-1);
  const html = `<!doctype html>
<html lang="${item.lang}">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>${title} — Oxi</title>
<meta name="description" content="${item.description}">
<link rel="canonical" href="https://oxi-dd65f4.gitlab.io/articles/${canonicalName}">
<link rel="alternate" hreflang="${item.lang === "en" ? "ja" : "en"}" href="https://oxi-dd65f4.gitlab.io/articles/${item.alternate.split("/").at(-1)}">
<meta property="og:type" content="article">
<meta property="og:title" content="${title}">
<meta property="og:description" content="${item.description}">
<meta property="og:image" content="https://oxi-dd65f4.gitlab.io/oxi-logo.png">
<meta name="twitter:card" content="summary_large_image">
<link rel="icon" href="../favicon.ico">
<script src="../analytics.js" defer></script>
<style>${style}</style>
</head>
<body>
<header class="site"><div class="site-inner"><a class="brand" href="${item.home}">Oxi</a><nav><a href="${item.home}">Home</a><a href="${item.alternate}" hreflang="${item.lang === "en" ? "ja" : "en"}">${item.alternateLabel}</a></nav></div></header>
<article>
<div class="meta">Engineering · Word compatibility · Rust</div>
<h1>${title}</h1>
${articleBody}
<p class="end"><a href="${item.home}">← Oxi</a></p>
</article>
</body>
</html>`;
  await writeFile(resolve(root, item.output), html);
}
