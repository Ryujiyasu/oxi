# -*- coding: utf-8 -*-
"""Rewrite the docx blind-benchmark sections of docs/index.html and
docs/ja/index.html from the blind-J result files (frozen 2026-11-03).

Same surgery as update_site_blind_i.py: the two docx bar charts (rows regenerated
from the engine columns), the two prose paragraphs under the h3 headings, the h4
dates, the figcaption measurement date, the stat tiles, the FAQ answer (JSON-LD
and the visible copy), and -- through update_site_blind_numbers -- every scatter
plot's points and "Oxi ahead on" caption, plus each scatter's aria-label
direction ("most points sit above / below the diagonal"), which follows the
document-by-document count. The PPTX chart is untouched.

  python tools/metrics/update_site_blind_j.py           # write both pages
  python tools/metrics/update_site_blind_j.py --check   # report only
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "tools" / "metrics"))
import update_site_blind_numbers as scat  # noqa: E402
import update_site_blind_h as H  # noqa: E402  (bar_rows / columns / rewrite surgery)
import update_site_blind_i as I  # noqa: E402  (wins_vs / fix_scatter_direction)

RESULTS = {
    "en": REPO / "pipeline_data/en_benchmark/ssim_blindJ50/_result.json",
    "ja": REPO / "pipeline_data/ja_benchmark/ssim_blindJ50/_result.json",
}
PAGES = {"en": REPO / "docs/index.html", "ja": REPO / "docs/ja/index.html"}
DATE = "2026-11-03"
WORD = "16.0.20527.20088"
NAMES = {
    "oxi": ("Oxi", "0.9.0"), "oo": ("ONLYOFFICE", "10.0.0"), "wordcraft": ("WordCraft", "0.15.0"),
    "libre": ("LibreOffice", "26.8.2.2"), "silurus": ("SILURUS", "0.89.7"),
    "genoffice": ("GenOffice", "e1f50e90"), "ibis": ("ibisOffice", "0.8.5"),
    "officecli": ("OfficeCLI", "1.0.160"), "betteroffice": ("BetterOffice", "0.5.1"),
    "eigenpal": ("eigenpal", "1.9.0"), "polaris": ("Polaris Office", "11.115.233"),
}


def prose(lang: str, en: list[dict], ja: list[dict], W: dict) -> dict[str, str]:
    e = {c["key"]: c for c in en}; j = {c["key"]: c for c in ja}
    en_pen_list = [(c["name"], c["pen"]) for c in sorted(en, key=lambda c: -c["pen"])]
    if lang == "en":
        pens = ", ".join(f"{n} {p:.3f}" for n, p in en_pen_list)
        p_en = (f'<p>English is measured the hard way: a <strong>frozen blind benchmark</strong> of 50 documents from the public docx-corpus that are never used as fix targets. '
                f'The current set is <strong>blind-J50</strong> (the next unclaimed SHA-256 ranks per type, frozen {DATE} before any measurement, rule in <code>pipeline_data/FROZEN_BLIND_J.md</code>); this is its first and only measurement, and every engine was at its latest public release that day. '
                f'<strong>Oxi is first on pixels at {e["oxi"]["common"]:.3f}</strong>, {e["oxi"]["common"]-e["oo"]["common"]:.3f} ahead of ONLYOFFICE ({e["oo"]["common"]:.3f}) &mdash; inside the &plusmn;0.02 one sample carries &mdash; and ahead of LibreOffice ({e["libre"]["common"]:.3f}) and WordCraft ({e["wordcraft"]["common"]:.3f}), and reproduces Word’s page count on the most documents: {e["oxi"]["pages"]} of 50 against ONLYOFFICE’s {e["oo"]["pages"]}. '
                f'Document by document it is closer than the mean: ONLYOFFICE is nearer to Word on {W["en_oo"][1]} of 50 and Oxi on {W["en_oo"][0]}; Oxi beats WordCraft on {W["en_wordcraft"][0]} and LibreOffice on {W["en_libre"][0]}. '
                f'Charging every page an engine adds or drops as zero, the penalized means are {pens}. '
                f'Ground truth is Microsoft Word (Microsoft 365 {WORD}) at 150 DPI, one fresh Word process per document.</p>')
        p_ja = (f'<p>Japanese is measured to the same standard. The current set is <strong>jaBlind-J50</strong> (frozen {DATE} <strong>before</strong> any measurement; the public manifest has no Japanese <em>technical</em> documents left, so those five slots went to the next entries across the other types by the rule declared in advance — all five are forms). '
                f'<strong>Oxi leads on pixels ({j["oxi"]["common"]:.3f})</strong> by {j["oxi"]["common"]-j["ibis"]["common"]:.3f} over ibisOffice ({j["ibis"]["common"]:.3f}) and {j["oxi"]["common"]-j["libre"]["common"]:.3f} over LibreOffice ({j["libre"]["common"]:.3f}), and is closer to Word than ibisOffice on {W["ja_ibis"][0]} of 50 documents, than LibreOffice on {W["ja_libre"][0]}, than ONLYOFFICE on {W["ja_oo"][0]} and than Polaris Office on {W["ja_polaris"][0]}; it also places the most page breaks where Word does ({j["oxi"]["pages"]}/50). '
                f'Engine rankings do not transfer between languages &mdash; ONLYOFFICE is second in English and eighth in Japanese, ibisOffice eighth in English and second in Japanese &mdash; which is why a blind set is kept per language. '
                f'Earlier Japanese sets scored 0.842 / 0.802 / 0.802 / 0.857 / 0.905 on the engines of their day; one sample carries the usual &plusmn;0.02.</p>')
        faq = (f'Oxi\'s layout is measured against Microsoft Word page by page with pixel-level SSIM, on frozen blind benchmarks of never-seen documents that are never used as fix targets. On the English blind-J50 set (frozen and measured {DATE}, every engine at its latest release) Oxi scores {e["oxi"]["common"]:.3f} per document, first of eleven engines, and reproduces Word\'s page count on {e["oxi"]["pages"]} of 50, the most of any engine; ONLYOFFICE is close behind ({e["oo"]["common"]:.3f}, {e["oo"]["pages"]}/50) and LibreOffice further ({e["libre"]["common"]:.3f}, {e["libre"]["pages"]}/50). On the Japanese blind-J50 set Oxi leads at {j["oxi"]["common"]:.3f}, ahead of ibisOffice ({j["ibis"]["common"]:.3f}) and LibreOffice ({j["libre"]["common"]:.3f}), and places the most page breaks correctly ({j["oxi"]["pages"]}/50). Engine rankings do not transfer between languages, which is why a blind set is kept per language.')
        return {"p_en": p_en, "p_ja": p_ja, "faq": faq,
                "h4_en": f"English blind set — 50 never-touched documents, per-document mean SSIM vs Word ({DATE})",
                "h4_ja": f"Japanese blind set — 50 never-touched documents, per-document mean SSIM vs Word ({DATE})"}
    pens = "、".join(f"{n} {p:.3f}" for n, p in en_pen_list)
    p_en = (f'<p>英語は最初から一番厳しい測り方をしています — 公開 docx-corpus から、修正のターゲットには一切使わない<strong>凍結ブラインドベンチ</strong> 50 文書。現行セットは <strong>blind-J50</strong>（種別ごとに未使用の次の SHA-256 順位、{DATE} に<strong>測定前</strong>凍結、規則は <code>pipeline_data/FROZEN_BLIND_J.md</code>）で、これがその唯一の測定です。比較エンジンはすべて測定日の最新公開版です。'
            f'<strong>英語で Oxi はピクセル首位（{e["oxi"]["common"]:.3f}）</strong> — ONLYOFFICE（{e["oo"]["common"]:.3f}）に {e["oxi"]["common"]-e["oo"]["common"]:.3f} 差（1 サンプルの ±0.02 の内側）、LibreOffice（{e["libre"]["common"]:.3f}）と WordCraft（{e["wordcraft"]["common"]:.3f}）を上回ります。Word とのページ数一致も Oxi が最多の {e["oxi"]["pages"]}/50（ONLYOFFICE {e["oo"]["pages"]}/50）。'
            f'文書ごとに見ると接戦で、ONLYOFFICE は 50 文書中 {W["en_oo"][1]}、Oxi は {W["en_oo"][0]} 文書で Word に近く、Oxi は WordCraft に {W["en_wordcraft"][0]}、LibreOffice に {W["en_libre"][0]} 文書で勝ちます。'
            f'増減したページを 0 点で数える penalized 平均は {pens}。'
            f'真値は Microsoft Word（Microsoft 365 {WORD}）、150 DPI、文書ごとに新しい Word プロセスで出力。</p>')
    p_ja = (f'<p>日本語も同じ厳しさで測ります。現行セットは <strong>jaBlind-J50</strong>（{DATE} に<strong>測定前</strong>凍結。公開マニフェストの日本語 <em>technical</em> は枯渇しているため、その 5 枠は事前に宣言した規則で他種別の次の順位から補充 — 5 本とも forms）。'
            f'<strong>Oxi はピクセルで首位（{j["oxi"]["common"]:.3f}）</strong>で、2 位の ibisOffice（{j["ibis"]["common"]:.3f}）に {j["oxi"]["common"]-j["ibis"]["common"]:.3f}、LibreOffice（{j["libre"]["common"]:.3f}）に {j["oxi"]["common"]-j["libre"]["common"]:.3f} の差。文書ごとでも ibisOffice に {W["ja_ibis"][0]}、LibreOffice に {W["ja_libre"][0]}、ONLYOFFICE に {W["ja_oo"][0]}、Polaris Office に {W["ja_polaris"][0]} 文書で勝ち、ページ数一致も最多です（{j["oxi"]["pages"]}/50）。'
            f'エンジンの順位は言語で入れ替わります — ONLYOFFICE は英語 2 位・日本語 8 位、ibisOffice は英語 8 位・日本語 2 位 — これが言語ごとにブラインドを持つ理由です。'
            f'過去の日本語セットは当時のエンジンで 0.842 / 0.802 / 0.802 / 0.857 / 0.905。1 サンプルなので、例の ±0.02 がそのまま付きます。</p>')
    faq = (f'修正のターゲットには使わない初見文書だけの凍結ブラインドベンチマークで、Word とピクセル単位の SSIM を測定しています。英語 blind-J50（{DATE} 凍結・測定、比較エンジンはすべて最新版）で Oxi は文書平均 {e["oxi"]["common"]:.3f} で 11 エンジン中首位、Word とのページ数一致は 50 文書中 {e["oxi"]["pages"]} で最多。ONLYOFFICE が僅差で続き（{e["oo"]["common"]:.3f}、{e["oo"]["pages"]}/50）、LibreOffice は {e["libre"]["common"]:.3f}（{e["libre"]["pages"]}/50）。日本語 blind-J50 では Oxi が {j["oxi"]["common"]:.3f} で ibisOffice（{j["ibis"]["common"]:.3f}）と LibreOffice（{j["libre"]["common"]:.3f}）を上回り、ページ数一致も最多（{j["oxi"]["pages"]}/50）です。エンジンの順位は言語によって入れ替わるため、言語ごとにブラインドセットを持っています。')
    return {"p_en": p_en, "p_ja": p_ja, "faq": faq,
            "h4_en": f"英語ブラインドセット — 初見の 50 文書、Word との文書平均 SSIM（{DATE}）",
            "h4_ja": f"日本語ブラインドセット — 初見の 50 文書、Word との文書平均 SSIM（{DATE}）"}


def main() -> None:
    check = "--check" in sys.argv
    rows = {k: json.loads(p.read_text(encoding="utf-8"))["docs"] for k, p in RESULTS.items()}
    H.NAMES = NAMES
    H.DATE = DATE
    W = {f"{l}_{e}": I.wins_vs(rows[l], e) for l in ("en", "ja") for e in ("oo", "libre", "wordcraft", "polaris", "ibis")}
    H.prose = lambda lang, en, ja, wins: prose(lang, en, ja, W)
    cols = {k: H.columns(v) for k, v in rows.items()}
    for l, cs in cols.items():
        print(l, [(c["name"], round(c["common"], 4), round(c["pen"], 4), c["pages"], c["n"]) for c in cs])
    print("wins (oxi, other):", W)
    scat.RESULTS = RESULTS
    clean = {l: [{k: (v if v is not None else {}) for k, v in r.items()} for r in rs] for l, rs in rows.items()}
    for lang, page in PAGES.items():
        html = page.read_text(encoding="utf-8")
        new, snotes = scat.rewrite(html, clean)
        new, notes = H.rewrite_page(lang, new, cols, {})
        new, k = I.fix_scatter_direction(new, clean, lang)
        print(f"--- {page.relative_to(REPO)}"); [print(n) for n in notes + snotes]; print(f"  scatter labels checked: {k}")
        if not check and new != html:
            page.write_text(new, encoding="utf-8"); print("  written")


if __name__ == "__main__":
    main()
