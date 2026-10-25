# -*- coding: utf-8 -*-
"""Rewrite the docx blind-benchmark sections of docs/index.html and
docs/ja/index.html from the blind-I result files (frozen 2026-10-25).

Same surgery as update_site_blind_h.py: the two docx bar charts (rows regenerated
from the engine columns), the two prose paragraphs under the h3 headings, the h4
dates, the figcaption measurement date, the stat tiles, the FAQ answer (JSON-LD
and the visible copy), and -- through update_site_blind_numbers -- every scatter
plot's points and "Oxi ahead on" caption, plus each scatter's aria-label
direction ("most points sit above / below the diagonal"), which follows the
document-by-document count. The PPTX chart is untouched.

  python tools/metrics/update_site_blind_i.py           # write both pages
  python tools/metrics/update_site_blind_i.py --check   # report only
"""
from __future__ import annotations

import json
import re
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "tools" / "metrics"))
import update_site_blind_numbers as scat  # noqa: E402
import update_site_blind_h as H  # noqa: E402  (bar_rows / columns / rewrite surgery)

RESULTS = {
    "en": REPO / "pipeline_data/en_benchmark/ssim_blindI50/_result.json",
    "ja": REPO / "pipeline_data/ja_benchmark/ssim_blindI50/_result.json",
}
PAGES = {"en": REPO / "docs/index.html", "ja": REPO / "docs/ja/index.html"}
DATE = "2026-10-25"
WORD = "16.0.20430.20182"
NAMES = {
    "oxi": ("Oxi", "0.8.9"), "oo": ("ONLYOFFICE", "10.0.0"), "wordcraft": ("WordCraft", "0.10.0"),
    "libre": ("LibreOffice", "26.8.1.1"), "silurus": ("SILURUS", "0.89.4"),
    "genoffice": ("GenOffice", "28cee5ca"), "ibis": ("ibisOffice", "0.8.4"),
    "officecli": ("OfficeCLI", "1.0.158"), "betteroffice": ("BetterOffice", "0.4.3"),
    "eigenpal": ("eigenpal", "1.9.0"),
}


def wins_vs(rows: list[dict], eng: str) -> tuple[int, int]:
    w = l = 0
    for r in rows:
        o = (r.get("oxi") or {}).get("common_mean"); e = (r.get(eng) or {}).get("common_mean")
        if o is None or e is None:
            continue
        w += o > e; l += e > o
    return w, l


def prose(lang: str, en: list[dict], ja: list[dict], W: dict) -> dict[str, str]:
    e = {c["key"]: c for c in en}; j = {c["key"]: c for c in ja}
    pen_en = "、".join if lang == "ja" else ", ".join
    en_pen_list = [(c["name"], c["pen"]) for c in sorted(en, key=lambda c: -c["pen"])]
    if lang == "en":
        pens = ", ".join(f"{n} {p:.3f}" for n, p in en_pen_list)
        p_en = (f'<p>English is measured the hard way: a <strong>frozen blind benchmark</strong> of 50 documents from the public docx-corpus that are never used as fix targets. '
                f'The current set is <strong>blind-I50</strong> (the next unclaimed SHA-256 ranks per type, frozen {DATE} before any measurement, rule in <code>pipeline_data/FROZEN_BLIND_I.md</code>); this is its first and only measurement, and every engine was at its latest public release that day. '
                f'<strong>Oxi is second on pixels at {e["oxi"]["common"]:.3f}</strong>, {e["oo"]["common"]-e["oxi"]["common"]:.3f} behind ONLYOFFICE ({e["oo"]["common"]:.3f}) and ahead of WordCraft ({e["wordcraft"]["common"]:.3f}) and LibreOffice ({e["libre"]["common"]:.3f}), and reproduces Word’s page count on the most documents: {e["oxi"]["pages"]} of 50 against ONLYOFFICE’s {e["oo"]["pages"]}. '
                f'Document by document the picture is less flattering: ONLYOFFICE is closer to Word on {W["en_oo"][1]} of 50 and WordCraft on {W["en_wordcraft"][1]}, while Oxi beats LibreOffice on {W["en_libre"][0]} — Oxi’s mean holds up because its bad documents are less bad, not because it wins the typical one. '
                f'Charging every page an engine adds or drops as zero, the penalized means are {pens}. '
                f'Ground truth is Microsoft Word (Microsoft 365 {WORD}) at 150 DPI, one fresh Word process per document. Polaris Office was installed but its window-driven export timed out on every document, so it is reported as not measured.</p>')
        p_ja = (f'<p>Japanese is measured to the same standard. The current set is <strong>jaBlind-I50</strong> (frozen {DATE} <strong>before</strong> any measurement; the public manifest has no Japanese <em>technical</em> documents left, so those five slots went to the next entries across the other types by the rule declared in advance — all five are forms). '
                f'<strong>Oxi leads on pixels ({j["oxi"]["common"]:.3f})</strong> by {j["oxi"]["common"]-j["libre"]["common"]:.3f} over LibreOffice ({j["libre"]["common"]:.3f}), and is closer to Word than LibreOffice on {W["ja_libre"][0]} of 50 documents, than WordCraft on {W["ja_wordcraft"][0]} and than ONLYOFFICE on {W["ja_oo"][0]}; it also places the most page breaks where Word does ({j["oxi"]["pages"]}/50). '
                f'Engine rankings do not transfer between languages &mdash; ONLYOFFICE is first in English and eighth in Japanese &mdash; which is why a blind set is kept per language. '
                f'Earlier Japanese sets scored 0.842 / 0.802 / 0.802 / 0.857 on the engines of their day; this one is the first above 0.90 &mdash; one sample, so the usual &plusmn;0.02 applies.</p>')
        faq = (f'Oxi\'s layout is measured against Microsoft Word page by page with pixel-level SSIM, on frozen blind benchmarks of never-seen documents that are never used as fix targets. On the English blind-I50 set (frozen and measured {DATE}, every engine at its latest release) Oxi scores {e["oxi"]["common"]:.3f} per document and reproduces Word\'s page count on {e["oxi"]["pages"]} of 50, the most of any engine; ONLYOFFICE is slightly ahead on pixels ({e["oo"]["common"]:.3f}, {e["oo"]["pages"]}/50) and LibreOffice behind ({e["libre"]["common"]:.3f}, {e["libre"]["pages"]}/50). On the Japanese blind-I50 set Oxi leads at {j["oxi"]["common"]:.3f}, ahead of LibreOffice ({j["libre"]["common"]:.3f}) and WordCraft ({j["wordcraft"]["common"]:.3f}), and places the most page breaks correctly ({j["oxi"]["pages"]}/50). Engine rankings do not transfer between languages, which is why a blind set is kept per language.')
        return {"p_en": p_en, "p_ja": p_ja, "faq": faq,
                "h4_en": f"English blind set — 50 never-touched documents, per-document mean SSIM vs Word ({DATE})",
                "h4_ja": f"Japanese blind set — 50 never-touched documents, per-document mean SSIM vs Word ({DATE})"}
    pens = "、".join(f"{n} {p:.3f}" for n, p in en_pen_list)
    p_en = (f'<p>英語は最初から一番厳しい測り方をしています — 公開 docx-corpus から、修正のターゲットには一切使わない<strong>凍結ブラインドベンチ</strong> 50 文書。現行セットは <strong>blind-I50</strong>（種別ごとに未使用の次の SHA-256 順位、{DATE} に<strong>測定前</strong>凍結、規則は <code>pipeline_data/FROZEN_BLIND_I.md</code>）で、これがその唯一の測定です。比較エンジンはすべて測定日の最新公開版です。'
            f'<strong>英語で Oxi はピクセル 2 位（{e["oxi"]["common"]:.3f}）</strong> — ONLYOFFICE（{e["oo"]["common"]:.3f}）に {e["oo"]["common"]-e["oxi"]["common"]:.3f} 差で次ぎ、WordCraft（{e["wordcraft"]["common"]:.3f}）と LibreOffice（{e["libre"]["common"]:.3f}）を上回ります。Word とのページ数一致は Oxi が最多の {e["oxi"]["pages"]}/50（ONLYOFFICE {e["oo"]["pages"]}/50）。'
            f'ただし文書ごとに見ると、ONLYOFFICE は 50 文書中 {W["en_oo"][1]}、WordCraft は {W["en_wordcraft"][1]} 文書で Oxi より Word に近く、Oxi が LibreOffice に勝つのは {W["en_libre"][0]} 文書です — Oxi の平均が高いのは悪い文書が少ないからで、典型的な文書で勝っているからではありません。'
            f'増減したページを 0 点で数える penalized 平均は {pens}。'
            f'真値は Microsoft Word（Microsoft 365 {WORD}）、150 DPI、文書ごとに新しい Word プロセスで出力。Polaris Office はインストールしたものの、ウィンドウ操作による書き出しが全文書で時間切れになったため未測定としています。</p>')
    p_ja = (f'<p>日本語も同じ厳しさで測ります。現行セットは <strong>jaBlind-I50</strong>（{DATE} に<strong>測定前</strong>凍結。公開マニフェストの日本語 <em>technical</em> は枯渇しているため、その 5 枠は事前に宣言した規則で他種別の次の順位から補充 — 5 本とも forms）。'
            f'<strong>Oxi はピクセルで首位（{j["oxi"]["common"]:.3f}）</strong>で、2 位の LibreOffice（{j["libre"]["common"]:.3f}）に {j["oxi"]["common"]-j["libre"]["common"]:.3f} の差。文書ごとでも LibreOffice に {W["ja_libre"][0]}、WordCraft に {W["ja_wordcraft"][0]}、ONLYOFFICE に {W["ja_oo"][0]} 文書で勝ち、ページ数一致も最多です（{j["oxi"]["pages"]}/50）。'
            f'エンジンの順位は言語で入れ替わります — ONLYOFFICE は英語 1 位・日本語 8 位 — これが言語ごとにブラインドを持つ理由です。'
            f'過去の日本語セットは当時のエンジンで 0.842 / 0.802 / 0.802 / 0.857。本セットが初めて 0.90 を超えました — ただし 1 サンプルなので、例の ±0.02 がそのまま付きます。</p>')
    faq = (f'修正のターゲットには使わない初見文書だけの凍結ブラインドベンチマークで、Word とピクセル単位の SSIM を測定しています。英語 blind-I50（{DATE} 凍結・測定、比較エンジンはすべて最新版）で Oxi は文書平均 {e["oxi"]["common"]:.3f}、Word とのページ数一致は 50 文書中 {e["oxi"]["pages"]} で最多。ピクセルでは ONLYOFFICE がわずかに上（{e["oo"]["common"]:.3f}、{e["oo"]["pages"]}/50）、LibreOffice が下（{e["libre"]["common"]:.3f}、{e["libre"]["pages"]}/50）。日本語 blind-I50 では Oxi が {j["oxi"]["common"]:.3f} で LibreOffice（{j["libre"]["common"]:.3f}）と WordCraft（{j["wordcraft"]["common"]:.3f}）を上回り、ページ数一致も最多（{j["oxi"]["pages"]}/50）です。エンジンの順位は言語によって入れ替わるため、言語ごとにブラインドセットを持っています。')
    return {"p_en": p_en, "p_ja": p_ja, "faq": faq,
            "h4_en": f"英語ブラインドセット — 初見の 50 文書、Word との文書平均 SSIM（{DATE}）",
            "h4_ja": f"日本語ブラインドセット — 初見の 50 文書、Word との文書平均 SSIM（{DATE}）"}


def fix_scatter_direction(html: str, rows: dict, lang_page: str) -> tuple[str, int]:
    """aria-label 'most points sit above/below the diagonal' follows the per-document count."""
    n = 0

    def repl(m: re.Match) -> str:
        nonlocal n
        label = m.group(0)
        ja_set = ("Japanese" in label) or ("日本語" in label)
        eng = next((k for name, k in scat.ENGINES.items() if f" {name} " in label or f"と {name} の" in label), None)
        if eng is None:
            return label
        w, l = wins_vs(rows["ja" if ja_set else "en"], eng)
        above = w >= l  # x = other engine, y = Oxi: Oxi ahead -> above the diagonal
        n += 1
        if lang_page == "en":
            return re.sub(r"most points sit (above|below) the diagonal", f"most points sit {'above' if above else 'below'} the diagonal", label)
        return re.sub(r"多くの点が対角線(より上|の上|の下|より下)", "多くの点が対角線の上" if above else "多くの点が対角線の下", label)

    html = re.sub(r'aria-label="(?:Scatter plot of 50 blind|散布図: )[^"]*"', repl, html)
    return html, n


def main() -> None:
    check = "--check" in sys.argv
    rows = {k: json.loads(p.read_text(encoding="utf-8"))["docs"] for k, p in RESULTS.items()}
    H.NAMES = NAMES
    H.DATE = DATE
    H.prose = lambda lang, en, ja, wins: prose(lang, en, ja, W)
    cols = {k: H.columns(v) for k, v in rows.items()}
    W = {f"{l}_{e}": wins_vs(rows[l], e) for l in ("en", "ja") for e in ("oo", "libre", "wordcraft")}
    for l, cs in cols.items():
        print(l, [(c["name"], round(c["common"], 4), round(c["pen"], 4), c["pages"], c["n"]) for c in cs])
    print("wins (oxi, other):", W)
    scat.RESULTS = RESULTS
    clean = {l: [{k: (v if v is not None else {}) for k, v in r.items()} for r in rs] for l, rs in rows.items()}
    for lang, page in PAGES.items():
        html = page.read_text(encoding="utf-8")
        new, snotes = scat.rewrite(html, clean)
        new, notes = H.rewrite_page(lang, new, cols, {})
        new, k = fix_scatter_direction(new, clean, lang)
        print(f"--- {page.relative_to(REPO)}"); [print(n) for n in notes + snotes]; print(f"  scatter labels checked: {k}")
        if not check and new != html:
            page.write_text(new, encoding="utf-8"); print("  written")


if __name__ == "__main__":
    main()
