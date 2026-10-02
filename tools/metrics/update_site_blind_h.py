# -*- coding: utf-8 -*-
"""Rewrite the docx blind-benchmark sections of docs/index.html and
docs/ja/index.html from the blind-H result files (2026-10-02).

What changes: the two docx bar charts (rows regenerated from the engine
columns), the two prose paragraphs under the h3 headings, the h4 dates, the
figcaption measurement date, the stat tiles, the FAQ answer (JSON-LD and the
visible copy), and — through update_site_blind_numbers — every scatter plot's
points and "Oxi ahead on" caption. The BetterOffice scatter is removed because
that engine has no blind-H column. The PPTX chart is untouched.

  python tools/metrics/update_site_blind_h.py           # write both pages
  python tools/metrics/update_site_blind_h.py --check   # report only
"""
from __future__ import annotations

import json
import re
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "tools" / "metrics"))
import update_site_blind_numbers as scat  # noqa: E402

RESULTS = {
    "en": REPO / "pipeline_data/en_benchmark/ssim_blindH50/_result.json",
    "ja": REPO / "pipeline_data/ja_benchmark/ssim_blindH50/_result.json",
}
PAGES = {"en": REPO / "docs/index.html", "ja": REPO / "docs/ja/index.html"}
DATE = "2026-10-02"
NAMES = {
    "oxi": ("Oxi", "0.8.2"), "oo": ("ONLYOFFICE", "9.3.1.8"), "libre": ("LibreOffice", "26.8.0.3"),
    "polaris": ("Polaris Office", "11.115"), "genoffice": ("GenOffice", "6a241c4"),
    "silurus": ("SILURUS", "0.86.1"), "eigenpal": ("eigenpal", "1.9.0"),
}


def columns(rows: list[dict]) -> list[dict]:
    out = []
    for key, (name, ver) in NAMES.items():
        vals = [r[key] for r in rows if isinstance(r.get(key), dict) and r[key].get("common_mean") is not None]
        if not vals:
            continue
        n = len(vals)
        common = sum(v["common_mean"] for v in vals) / n
        pen = sum(v.get("penalized_mean", v["common_mean"]) for v in vals) / n
        pages = sum(1 for v in vals if v.get("page_delta") == 0)
        out.append({"key": key, "name": name, "ver": ver, "common": common, "pen": pen, "pages": pages, "n": n, "total": len(rows)})
    out.sort(key=lambda c: -c["common"])
    return out


def bar_rows(cols: list[dict], lang: str) -> str:
    rows = []
    for c in cols:
        if lang == "en":
            sub = f"page count matches Word on {c['pages']}/{c['total']}" + (f" · v{c['ver']}" if c["key"] != "oxi" else "")
            if c["n"] < c["total"]:
                sub += f" · {c['n']} of {c['total']} rendered"
        else:
            sub = f"ページ数一致 {c['pages']}/{c['total']}" + (f" · v{c['ver']}" if c["key"] != "oxi" else "")
            if c["n"] < c["total"]:
                sub += f" · {c['total']} 文書中 {c['n']} を描画"
        label = f"<strong>{c['name']}</strong>" if c["key"] == "oxi" else c["name"]
        val = f"<strong>{c['common']:.3f}</strong>" if c["key"] == "oxi" else f"{c['common']:.3f}"
        fill = ' oxi' if c["key"] == "oxi" else ''
        rows.append(
            f'          <div class="brow">\n'
            f'            <div class="blabel">{label}<span>{sub}</span></div>\n'
            f'            <div class="btrack"><div class="bfill{fill}" style="width:{c["common"]*100:.1f}%"></div></div>\n'
            f'            <div class="bval">{val}</div>\n'
            f'          </div>')
    return "\n".join(rows)


def prose(lang: str, en: list[dict], ja: list[dict], wins: dict) -> dict[str, str]:
    e = {c["key"]: c for c in en}; j = {c["key"]: c for c in ja}
    if lang == "en":
        p_en = (f'<p>English is measured the hard way: a <strong>frozen blind benchmark</strong> of 50 documents from the public docx-corpus that are never used as fix targets. '
                f'The current set is <strong>blind-H50</strong> (the next unclaimed SHA-256 ranks per type, frozen {DATE} before any measurement, rule in <code>pipeline_data/FROZEN_BLIND_H.md</code>); this is its first and only measurement. '
                f'<strong>Oxi is second on pixels at {e["oxi"]["common"]:.3f}</strong>, behind ONLYOFFICE ({e["oo"]["common"]:.3f}) and ahead of LibreOffice ({e["libre"]["common"]:.3f}, Oxi closer to Word on {wins["en_libre"]} of 50), '
                f'and reproduces Word’s page count on {e["oxi"]["pages"]} of 50 against ONLYOFFICE’s {e["oo"]["pages"]} and LibreOffice’s {e["libre"]["pages"]}. '
                f'Charging every page Oxi adds or drops as zero, the penalized means are ONLYOFFICE {e["oo"]["pen"]:.3f}, Oxi {e["oxi"]["pen"]:.3f}, LibreOffice {e["libre"]["pen"]:.3f}, Polaris Office {e["polaris"]["pen"]:.3f}, GenOffice {e["genoffice"]["pen"]:.3f}, SILURUS {e["silurus"]["pen"]:.3f}, eigenpal {e["eigenpal"]["pen"]:.3f}. '
                f'Ground truth is Microsoft Word (Microsoft 365 16.0.20430.20092) at 150 DPI; engines ONLYOFFICE 9.3.1.8, LibreOffice 26.8.0.3, Polaris Office 11.115, GenOffice 6a241c4, SILURUS @silurus/ooxml 0.86.1, eigenpal @eigenpal/docx-editor-react 1.9.0. '
                f'BetterOffice and OfficeCLI were not completed on this rotation (the measurement host ran out of memory during their browser renders); a blind set is scored once, so they are absent rather than measured later against a spent set.</p>')
        p_ja = (f'<p>Japanese is measured to the same standard. The current set is <strong>jaBlind-H50</strong> (frozen {DATE} <strong>before</strong> any measurement; the public manifest has no Japanese <em>technical</em> documents left, so those five slots went to the next entries across the other types by the rule declared in advance — all five are forms). '
                f'<strong>Oxi leads on pixels ({j["oxi"]["common"]:.3f})</strong> by {j["oxi"]["common"]-j["polaris"]["common"]:.3f} over Polaris Office ({j["polaris"]["common"]:.3f}) and {j["oxi"]["common"]-j["libre"]["common"]:.3f} over LibreOffice ({j["libre"]["common"]:.3f}, Oxi closer to Word on {wins["ja_libre"]} of 50), and this time also places the most page breaks where Word does ({j["oxi"]["pages"]}/50 against ONLYOFFICE’s {j["oo"]["pages"]}). '
                f'Engine rankings do not transfer between languages &mdash; ONLYOFFICE is first in English and sixth in Japanese &mdash; which is why a blind set is kept per language. '
                f'Three earlier Japanese sets scored 0.842 / 0.802 / 0.802 on the engines of their day; this one, three weeks and one engine generation later, is the first above 0.85 &mdash; one sample, so the usual &plusmn;0.02 applies.</p>')
        faq = (f'Oxi\'s layout is measured against Microsoft Word page by page with pixel-level SSIM, on frozen blind benchmarks of never-seen documents that are never used as fix targets. On the English blind-H50 set (frozen and measured {DATE}) Oxi scores {e["oxi"]["common"]:.3f} per document and reproduces Word\'s page count on {e["oxi"]["pages"]} of 50; ONLYOFFICE is ahead on pixels ({e["oo"]["common"]:.3f}, {e["oo"]["pages"]}/50) and LibreOffice behind ({e["libre"]["common"]:.3f}, {e["libre"]["pages"]}/50). On the Japanese blind-H50 set Oxi leads at {j["oxi"]["common"]:.3f}, ahead of Polaris Office ({j["polaris"]["common"]:.3f}) and LibreOffice ({j["libre"]["common"]:.3f}), and places the most page breaks correctly ({j["oxi"]["pages"]}/50). Engine rankings do not transfer between languages, which is why a blind set is kept per language.')
        return {"p_en": p_en, "p_ja": p_ja, "faq": faq,
                "h4_en": f"English blind set — 50 never-touched documents, per-document mean SSIM vs Word ({DATE})",
                "h4_ja": f"Japanese blind set — 50 never-touched documents, per-document mean SSIM vs Word ({DATE})"}
    p_en = (f'<p>英語は最初から一番厳しい測り方をしています — 公開 docx-corpus から、修正のターゲットには一切使わない<strong>凍結ブラインドベンチ</strong> 50 文書。現行セットは <strong>blind-H50</strong>（種別ごとに未使用の次の SHA-256 順位、{DATE} に<strong>測定前</strong>凍結、規則は <code>pipeline_data/FROZEN_BLIND_H.md</code>）で、これがその唯一の測定です。'
            f'<strong>英語で Oxi はピクセル 2 位（{e["oxi"]["common"]:.3f}）</strong> — ONLYOFFICE（{e["oo"]["common"]:.3f}）に次ぎ、LibreOffice（{e["libre"]["common"]:.3f}、50 文書中 {wins["en_libre"]} 文書で Oxi が Word に近い）を上回ります。Word とのページ数一致は Oxi {e["oxi"]["pages"]}/50、ONLYOFFICE {e["oo"]["pages"]}/50、LibreOffice {e["libre"]["pages"]}/50。'
            f'増減したページを 0 点で数える penalized 平均は ONLYOFFICE {e["oo"]["pen"]:.3f}、Oxi {e["oxi"]["pen"]:.3f}、LibreOffice {e["libre"]["pen"]:.3f}、Polaris Office {e["polaris"]["pen"]:.3f}、GenOffice {e["genoffice"]["pen"]:.3f}、SILURUS {e["silurus"]["pen"]:.3f}、eigenpal {e["eigenpal"]["pen"]:.3f}。'
            f'真値は Microsoft Word（Microsoft 365 16.0.20430.20092）、比較対象は ONLYOFFICE 9.3.1.8、LibreOffice 26.8.0.3、Polaris Office 11.115、GenOffice 6a241c4、SILURUS @silurus/ooxml 0.86.1、eigenpal @eigenpal/docx-editor-react 1.9.0 — すべて 150 DPI。'
            f'BetterOffice と OfficeCLI は今回の回では未完（ブラウザ描画中に測定機のメモリが尽きたため）。ブラインドは一度しか採点しないので、後から測って足すことはせず空欄のままにしています。</p>')
    p_ja = (f'<p>日本語も同じ厳しさで測ります。現行セットは <strong>jaBlind-H50</strong>（{DATE} に<strong>測定前</strong>凍結。公開マニフェストの日本語 <em>technical</em> は枯渇しているため、その 5 枠は事前に宣言した規則で他種別の次の順位から補充 — 5 本とも forms）。'
            f'<strong>Oxi はピクセルで首位（{j["oxi"]["common"]:.3f}）</strong>で、Polaris Office（{j["polaris"]["common"]:.3f}）に {j["oxi"]["common"]-j["polaris"]["common"]:.3f}、LibreOffice（{j["libre"]["common"]:.3f}、50 文書中 {wins["ja_libre"]} 文書で Oxi が Word に近い）に {j["oxi"]["common"]-j["libre"]["common"]:.3f} の差。今回はページ数一致も最多です（{j["oxi"]["pages"]}/50、ONLYOFFICE は {j["oo"]["pages"]}/50）。'
            f'エンジンの順位は言語で入れ替わります — ONLYOFFICE は英語 1 位・日本語 6 位 — これが言語ごとにブラインドを持つ理由です。'
            f'過去 3 つの日本語セットは当時のエンジンで 0.842 / 0.802 / 0.802。3 週間と 1 世代後の本セットが初めて 0.85 を超えました — ただし 1 サンプルなので、例の ±0.02 がそのまま付きます。</p>')
    faq = (f'修正のターゲットには使わない初見文書だけの凍結ブラインドベンチマークで、Word とピクセル単位の SSIM を測定しています。英語 blind-H50（{DATE} 凍結・測定）で Oxi は文書平均 {e["oxi"]["common"]:.3f}、Word とのページ数一致は 50 文書中 {e["oxi"]["pages"]}。ピクセルでは ONLYOFFICE が上（{e["oo"]["common"]:.3f}、{e["oo"]["pages"]}/50）、LibreOffice が下（{e["libre"]["common"]:.3f}、{e["libre"]["pages"]}/50）。日本語 blind-H50 では Oxi が {j["oxi"]["common"]:.3f} で Polaris Office（{j["polaris"]["common"]:.3f}）と LibreOffice（{j["libre"]["common"]:.3f}）を上回り、ページ数一致も最多（{j["oxi"]["pages"]}/50）です。エンジンの順位は言語によって入れ替わるため、言語ごとにブラインドセットを持っています。')
    return {"p_en": p_en, "p_ja": p_ja, "faq": faq,
            "h4_en": f"英語ブラインドセット — 初見の 50 文書、Word との文書平均 SSIM（{DATE}）",
            "h4_ja": f"日本語ブラインドセット — 初見の 50 文書、Word との文書平均 SSIM（{DATE}）"}


def rewrite_page(lang: str, html: str, cols: dict, wins: dict) -> tuple[str, list[str]]:
    notes = []
    texts = prose(lang, cols["en"], cols["ja"], wins)
    # bar charts: the first two <div class="barchart"> blocks (docx EN, docx JA)
    blocks = list(re.finditer(r'<div class="barchart">\n(.*?)\n        </div>', html, re.S))
    if len(blocks) < 2:
        notes.append("  ?? fewer than two barchart blocks"); return html, notes
    for m, l in ((blocks[1], "ja"), (blocks[0], "en")):  # replace from the back so offsets stay valid
        html = html[:m.start(1)] + bar_rows(cols[l], lang) + html[m.end(1):]
        notes.append(f"  {l} bar chart: {len(cols[l])} rows")
    # h4 headings (first = EN docx, second = JA docx)
    if lang == "en":
        html = re.sub(r'<h4>English blind set — \d+ never-touched documents, per-document mean SSIM vs Word \([^)]*\)</h4>', f'<h4>{texts["h4_en"]}</h4>', html, count=1)
        html = re.sub(r'<h4>Japanese blind set — \d+ never-touched documents, per-document mean SSIM vs Word \([^)]*\)</h4>', f'<h4>{texts["h4_ja"]}</h4>', html, count=1)
        html = re.sub(r'(<h3>English blind set — 50 never-seen documents</h3>\n\s*)<p>.*?</p>', lambda mm: mm.group(1) + texts["p_en"], html, count=1, flags=re.S)
        html = re.sub(r'(<h3>Japanese blind set — 50 never-seen documents</h3>\n\s*)<p>.*?</p>', lambda mm: mm.group(1) + texts["p_ja"], html, count=1, flags=re.S)
        html = re.sub(r'"text": "Oxi\'s layout is measured against Microsoft Word page by page[^"]*"', '"text": "' + texts["faq"].replace('"', '\\"') + '"', html, count=1)
        html = re.sub(r'<p>Oxi&rsquo;s layout is measured against Microsoft Word page by page.*?</p>', '<p>' + texts["faq"].replace("Oxi's", "Oxi&rsquo;s").replace("Word's", "Word&rsquo;s") + ' See the <a href="#accuracy">accuracy section</a> for the comparison charts and methodology.</p>', html, count=1, flags=re.S)
        html = re.sub(r'<div class="num">0\.\d\d</div><div class="lbl">SSIM vs Word — Japanese blind<br>benchmark \([^)]*\)</div>', f'<div class="num">{cols["ja"][0]["common"]:.2f}</div><div class="lbl">SSIM vs Word — Japanese blind<br>benchmark (50 never-seen docs)</div>', html)
        html = re.sub(r'<div class="num">0\.\d\d</div><div class="lbl">SSIM vs Word — English blind<br>benchmark \([^)]*\)</div>', f'<div class="num">{[c for c in cols["en"] if c["key"]=="oxi"][0]["common"]:.2f}</div><div class="lbl">SSIM vs Word — English blind<br>benchmark (50 never-seen docs)</div>', html)
        html = html.replace("every bar is the 2026-08-31 measurement at the version printed on it", f"every bar is the {DATE} measurement at the version printed on it")
    else:
        html = re.sub(r'<h4>英語ブラインドセット — 初見の \d+ 文書、Word との文書平均 SSIM（[^）]*）</h4>', f'<h4>{texts["h4_en"]}</h4>', html, count=1)
        html = re.sub(r'<h4>日本語ブラインドセット — 初見の \d+ 文書、Word との文書平均 SSIM（[^）]*）</h4>', f'<h4>{texts["h4_ja"]}</h4>', html, count=1)
        html = re.sub(r'(<h3>英語ブラインドセット — 初見の 50 文書</h3>\n\s*)<p>.*?</p>', lambda mm: mm.group(1) + texts["p_en"], html, count=1, flags=re.S)
        html = re.sub(r'(<h3>日本語ブラインドセット — 初見の 50 文書</h3>\n\s*)<p>.*?</p>', lambda mm: mm.group(1) + texts["p_ja"], html, count=1, flags=re.S)
        html = re.sub(r'"text": "修正のターゲットには使わない初見文書だけの凍結ブラインドベンチマーク[^"]*"', '"text": "' + texts["faq"] + '"', html, count=1)
        html = re.sub(r'<p>修正のターゲットには使わない初見文書だけの凍結ブラインドベンチマーク.*?</p>', '<p>' + texts["faq"] + ' 比較チャートと方法は<a href="#accuracy">精度の節</a>を参照。</p>', html, count=1, flags=re.S)
        html = re.sub(r'<div class="num">0\.\d\d</div><div class="lbl">Word との SSIM — 日本語ブラインド<br>', f'<div class="num">{cols["ja"][0]["common"]:.2f}</div><div class="lbl">Word との SSIM — 日本語ブラインド<br>', html)
        html = re.sub(r'<div class="num">0\.\d\d</div><div class="lbl">Word との SSIM — 英語ブラインド<br>', f'<div class="num">{[c for c in cols["en"] if c["key"]=="oxi"][0]["common"]:.2f}</div><div class="lbl">Word との SSIM — 英語ブラインド<br>', html)
        html = re.sub(r'すべての棒は 2026-\d\d-\d\d の測定', f'すべての棒は {DATE} の測定', html)
    # drop the BetterOffice scatter (no blind-H column)
    html, k = re.subn(r'\n?<svg viewBox="0 0 340 340" role="img" aria-label="[^"]*(?:vs BetterOffice|BetterOffice の SSIM)[^"]*"[^>]*>.*?</svg>\n?', "\n", html, flags=re.S)
    notes.append(f"  removed {k} BetterOffice scatter(s)")
    return html, notes


def main() -> None:
    check = "--check" in sys.argv
    rows = {k: json.loads(p.read_text(encoding="utf-8"))["docs"] for k, p in RESULTS.items()}
    cols = {k: columns(v) for k, v in rows.items()}
    wins = {}
    for l in ("en", "ja"):
        _, w, n = scat.circles([{k: (v if v is not None else {}) for k, v in r.items()} for r in rows[l]], "libre"); wins[f"{l}_libre"] = w
    for l, cs in cols.items():
        print(l, [(c["name"], round(c["common"], 4), round(c["pen"], 4), c["pages"], c["n"]) for c in cs])
    scat.RESULTS = RESULTS
    # engines that produced nothing for a document leave None in the row; the
    # scatter code expects a dict or a missing key
    clean = {l: [{k: (v if v is not None else {}) for k, v in r.items()} for r in rs] for l, rs in rows.items()}
    for lang, page in PAGES.items():
        html = page.read_text(encoding="utf-8")
        # scatters first (their language fallback counts the EN plots), then
        # the prose / bars / BetterOffice removal
        new, snotes = scat.rewrite(html, clean)
        new, notes = rewrite_page(lang, new, cols, wins)
        print(f"--- {page.relative_to(REPO)}"); [print(n) for n in notes + snotes]
        if not check and new != html:
            page.write_text(new, encoding="utf-8"); print("  written")


if __name__ == "__main__":
    main()
