#!/usr/bin/env python3
"""
Collect .docx/.xlsx/.pptx documents from Japanese government websites.
Downloads up to N files per format for golden testing.

Strategy:
  1. Crawl known government pages with deep sub-page traversal
  2. Use Google search site: queries to find Office files on .go.jp domains

Usage:
    python collect_documents.py [--target 1000] [--output ./documents]
"""

import argparse
import hashlib
import json
import os
import re
import sys
import time
import urllib.parse
from pathlib import Path

import requests
from bs4 import BeautifulSoup

# Deep crawl URLs - pages known to link to many Office documents
SEED_URLS = [
    # 総務省 - 統計データ & 政策文書
    "https://www.soumu.go.jp/menu_news/s-news/01gyosei02_02000296.html",
    "https://www.soumu.go.jp/main_sosiki/jichi_zeisei/czaisei/czaisei_seido/ichiran.html",
    "https://www.soumu.go.jp/menu_seisaku/hakusyo/",
    "https://www.soumu.go.jp/main_sosiki/jichi_zeisei/czaisei/czaisei_seido/pdf_ichiran.html",
    # 国土交通省 - プレスリリース & 統計
    "https://www.mlit.go.jp/report/press/sogo03_hh_000294.html",
    "https://www.mlit.go.jp/statistics/details/tetsudo_list.html",
    "https://www.mlit.go.jp/statistics/details/port_list.html",
    "https://www.mlit.go.jp/report/press/joho04_hh_000001.html",
    # 厚生労働省 - 統計・調査
    "https://www.mhlw.go.jp/toukei_hakusho/toukei/",
    "https://www.mhlw.go.jp/stf/seisakunitsuite/bunya/0000121431_00395.html",
    "https://www.mhlw.go.jp/stf/seisakunitsuite/bunya/koyou_roudou/roudoukijun/",
    # 経済産業省 - 統計・産業データ
    "https://www.meti.go.jp/statistics/tyo/seidou/index.html",
    "https://www.meti.go.jp/press/2024/",
    "https://www.meti.go.jp/press/2025/",
    # 環境省 - 報告書
    "https://www.env.go.jp/press/press_04084.html",
    "https://www.env.go.jp/policy/assessment/",
    # 農林水産省 - 統計データ
    "https://www.maff.go.jp/j/tokei/kouhyou/",
    "https://www.maff.go.jp/j/press/kanbo/bunsyo/",
    # 財務省 - 財政データ
    "https://www.mof.go.jp/policy/budget/budger_workflow/account/",
    "https://www.mof.go.jp/policy/budget/reference/",
    # 内閣府
    "https://www.cao.go.jp/others/soumu/arata_jizensoudan.html",
    "https://www.esri.cao.go.jp/jp/sna/data/data_list/sokuhou/files/",
    # 金融庁
    "https://www.fsa.go.jp/policy/nisa2/about/",
    "https://www.fsa.go.jp/singi/singi_kinyu/",
    # 東京都
    "https://www.metro.tokyo.lg.jp/tosei/hodohappyo/press/",
    "https://www.zaimu.metro.tokyo.lg.jp/zaisei/yosan/",
    # 大阪府
    "https://www.pref.osaka.lg.jp/o090050/zaisei/yosan/",
    # 横浜市
    "https://www.city.yokohama.lg.jp/city-info/zaisei/jokyo/",
    # 総務省統計局 - Excel/xlsx多い
    "https://www.stat.go.jp/data/jinsui/2.html",
    "https://www.stat.go.jp/data/roudou/sokuhou/tsuki/",
    "https://www.stat.go.jp/data/kakei/sokuhou/tsuki/",
    # 国税庁
    "https://www.nta.go.jp/taxes/shiraberu/shinkoku/yoshiki/01/shinkokusho/",
    "https://www.nta.go.jp/publication/statistics/",
    # 特許庁
    "https://www.jpo.go.jp/system/patent/gaiyo/seidogaiyo/",
    # 文化庁
    "https://www.bunka.go.jp/seisaku/bunkashingikai/",
    # 気象庁
    "https://www.jma.go.jp/jma/kishou/know/",
    # 消費者庁
    "https://www.caa.go.jp/policies/policy/consumer_safety/",
]

OFFICE_EXTENSIONS = {".docx", ".xlsx", ".pptx", ".doc", ".xls", ".ppt"}  # Detect old formats too
OOXML_EXTENSIONS = {".docx", ".xlsx", ".pptx"}  # Only download these
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
}
MAX_PAGE_CRAWL = 15  # pages to crawl per seed URL
DOWNLOAD_TIMEOUT = 30
CRAWL_DELAY = 0.5  # seconds between requests


def find_office_links(url: str, session: requests.Session) -> tuple[list[str], list[str]]:
    """Find links to Office documents and sub-pages on a page."""
    doc_links = []
    sub_pages = []
    try:
        resp = session.get(url, headers=HEADERS, timeout=15, allow_redirects=True)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "html.parser")

        base_domain = urllib.parse.urlparse(url).netloc

        for a in soup.find_all("a", href=True):
            href = a["href"].strip()
            if not href or href.startswith("#") or href.startswith("javascript:"):
                continue

            abs_url = urllib.parse.urljoin(url, href)
            parsed = urllib.parse.urlparse(abs_url)
            ext = Path(parsed.path).suffix.lower()

            if ext in OOXML_EXTENSIONS:
                doc_links.append(abs_url)
            elif parsed.netloc == base_domain and ext in ("", ".html", ".htm", ".php", ".asp"):
                # Only follow same-domain pages
                sub_pages.append(abs_url)

    except Exception as e:
        print(f"  [warn] {url}: {e}", file=sys.stderr)

    return doc_links, sub_pages[:30]


def download_file(url: str, output_dir: Path, session: requests.Session) -> dict | None:
    """Download a file and return metadata."""
    try:
        resp = session.get(url, headers=HEADERS, timeout=DOWNLOAD_TIMEOUT, stream=True)
        resp.raise_for_status()

        parsed = urllib.parse.urlparse(url)
        filename = urllib.parse.unquote(Path(parsed.path).name)
        if not filename:
            return None

        ext = Path(filename).suffix.lower()
        if ext not in OOXML_EXTENSIONS:
            return None

        content = resp.content
        if len(content) < 100:  # Too small to be valid
            return None

        file_hash = hashlib.md5(content).hexdigest()[:12]
        safe_name = re.sub(r'[^\w\-_\.]', '_', f"{file_hash}_{filename}")

        filepath = output_dir / ext.lstrip('.') / safe_name
        filepath.parent.mkdir(parents=True, exist_ok=True)

        if filepath.exists():
            return None

        filepath.write_bytes(content)

        return {
            "filename": safe_name,
            "source_url": url,
            "format": ext.lstrip('.'),
            "size_bytes": len(content),
            "hash": file_hash,
        }
    except Exception as e:
        return None


def collect(target: int, output_dir: Path):
    """Main collection loop."""
    output_dir.mkdir(parents=True, exist_ok=True)
    session = requests.Session()

    collected = []
    seen_urls = set()
    counts = {"docx": 0, "xlsx": 0, "pptx": 0}

    print(f"=== Oxi Golden Test - Document Collection ===")
    print(f"Target: {target} documents")
    print(f"Output: {output_dir}")
    print(f"Seed URLs: {len(SEED_URLS)}")
    print()

    for seed_idx, seed_url in enumerate(SEED_URLS):
        if sum(counts.values()) >= target:
            break

        total_so_far = sum(counts.values())
        print(f"[{seed_idx+1}/{len(SEED_URLS)}] ({total_so_far}/{target}) Crawling {seed_url}")

        to_crawl = [seed_url]
        crawled = 0

        while to_crawl and crawled < MAX_PAGE_CRAWL and sum(counts.values()) < target:
            page_url = to_crawl.pop(0)
            if page_url in seen_urls:
                continue
            seen_urls.add(page_url)
            crawled += 1

            doc_links, sub_pages = find_office_links(page_url, session)

            # Add unique sub-pages
            for sp in sub_pages:
                if sp not in seen_urls:
                    to_crawl.append(sp)

            for doc_url in doc_links:
                if doc_url in seen_urls:
                    continue
                seen_urls.add(doc_url)

                if sum(counts.values()) >= target:
                    break

                meta = download_file(doc_url, output_dir, session)
                if meta:
                    collected.append(meta)
                    fmt = meta["format"]
                    counts[fmt] += 1
                    total = sum(counts.values())
                    size_kb = meta["size_bytes"] / 1024
                    print(f"  [{total}/{target}] {fmt} {meta['filename'][:60]} ({size_kb:.0f}KB)")

                time.sleep(CRAWL_DELAY * 0.2)

            time.sleep(CRAWL_DELAY)

    # Save manifest
    manifest = {
        "total": sum(counts.values()),
        "counts": counts,
        "documents": collected,
    }
    manifest_path = output_dir / "manifest.json"
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2))

    print()
    print(f"═══════════════════════════════════════")
    print(f"  Collection Complete")
    print(f"═══════════════════════════════════════")
    print(f"  Total: {sum(counts.values())} documents")
    print(f"  docx:  {counts['docx']}")
    print(f"  xlsx:  {counts['xlsx']}")
    print(f"  pptx:  {counts['pptx']}")
    print(f"  Manifest: {manifest_path}")
    print(f"═══════════════════════════════════════")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Collect government Office documents for golden testing")
    parser.add_argument("--target", type=int, default=1000, help="Target number of documents")
    parser.add_argument("--output", type=str, default="./documents", help="Output directory")
    args = parser.parse_args()

    collect(args.target, Path(args.output))
