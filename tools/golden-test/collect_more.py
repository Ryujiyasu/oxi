#!/usr/bin/env python3
"""
Additional document collection - target deeper pages and more diverse formats.
Appends to existing ./documents/ directory.
"""
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

OOXML_EXTENSIONS = {".docx", ".xlsx", ".pptx"}
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
}

# More targeted URLs with known Office doc links
DEEP_URLS = [
    # 厚労省 - 各種様式 (docx多い)
    "https://www.mhlw.go.jp/bunya/roudoukijun/roudoujouken01/",
    "https://www.mhlw.go.jp/stf/seisakunitsuite/bunya/koyou_roudou/roudoukijun/zigyonushi/model/index.html",
    "https://www.mhlw.go.jp/stf/seisakunitsuite/bunya/0000099482.html",
    "https://www.mhlw.go.jp/stf/newpage_08517.html",
    "https://www.mhlw.go.jp/stf/seisakunitsuite/bunya/koyou_roudou/koyou/kyufukin/pageL07.html",
    "https://www.mhlw.go.jp/stf/seisakunitsuite/bunya/koyou_roudou/koyou/kyufukin/pageL07_00001.html",
    # 国交省 - 書式ダウンロード
    "https://www.mlit.go.jp/jutakukentiku/build/jutakukentiku_house_tk_000145.html",
    "https://www.mlit.go.jp/sogoseisaku/jouhouka/sosei_jouhouka_tk4_000002.html",
    "https://www.mlit.go.jp/tochi_fudousan_kensetsugyo/const/tochi_fudousan_kensetsugyo_const_tk1_000001.html",
    # 法務省 - 各種申請書 (docx)
    "https://www.moj.go.jp/MINJI/minji06_00118.html",
    "https://www.moj.go.jp/MINJI/minji06_00001.html",
    "https://www.moj.go.jp/MINJI/MINJI79/minji79.html",
    "https://www.moj.go.jp/hisho/kouhou/hisho06_00842.html",
    # 文科省 - 公募・様式
    "https://www.mext.go.jp/a_menu/hyouka/kekka/1421055_00015.htm",
    "https://www.mext.go.jp/a_menu/shinkou/hojyo/1235400.htm",
    # 経産省 - 統計データ (xlsx)
    "https://www.meti.go.jp/statistics/tyo/syoudou/result-2.html",
    "https://www.meti.go.jp/statistics/tyo/kougyo/index.html",
    "https://www.meti.go.jp/statistics/tyo/tokusabizi/result/result_1.html",
    # 環境省 - 審議会資料 (pptx可能性あり)
    "https://www.env.go.jp/council/",
    "https://www.env.go.jp/policy/hakusyo/",
    # 消費者庁
    "https://www.caa.go.jp/policies/policy/consumer_safety/food_safety/food_safety_portal/",
    "https://www.caa.go.jp/notice/entry/",
    # 特許庁 - 各種様式
    "https://www.jpo.go.jp/system/process/toroku/shintouki/index.html",
    "https://www.jpo.go.jp/system/laws/rule/guideline/",
    # 横浜市 - 各種書式
    "https://www.city.yokohama.lg.jp/kurashi/koseki-zei-hoken/todokede/",
    # 東京都 - プレスリリース (xlsx/docx)
    "https://www.metro.tokyo.lg.jp/tosei/hodohappyo/press/2025/01/",
    "https://www.metro.tokyo.lg.jp/tosei/hodohappyo/press/2025/02/",
    "https://www.metro.tokyo.lg.jp/tosei/hodohappyo/press/2025/03/",
    # 自治体オープンデータ
    "https://www.city.chiba.jp/somu/joho/kaikaku/opendata_top.html",
    "https://opendata.pref.aichi.jp/",
    # 総務省統計局
    "https://www.stat.go.jp/data/jinsui/new.html",
    "https://www.stat.go.jp/data/cpi/sokuhou/tsuki/index-z.html",
    "https://www.stat.go.jp/data/kouri/sokuhou/tsuki/index.html",
    "https://www.stat.go.jp/data/service/2019/index2.html",
    # 日本銀行 (xlsx多い)
    "https://www.boj.or.jp/statistics/money/zandaka/index.htm",
    "https://www.boj.or.jp/statistics/boj/other/acmai/index.htm",
    # 年金機構
    "https://www.nenkin.go.jp/service/kounen/todokesho/",
    # 国土地理院
    "https://www.gsi.go.jp/REPORT/",
    # 防衛省
    "https://www.mod.go.jp/j/press/",
]

def find_office_links(url, session):
    doc_links, sub_pages = [], []
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
            elif parsed.netloc == base_domain and ext in ("", ".html", ".htm"):
                sub_pages.append(abs_url)
    except Exception as e:
        pass
    return doc_links, sub_pages[:20]


def download_file(url, output_dir, session, existing_hashes):
    try:
        resp = session.get(url, headers=HEADERS, timeout=30, stream=True)
        resp.raise_for_status()
        parsed = urllib.parse.urlparse(url)
        filename = urllib.parse.unquote(Path(parsed.path).name)
        if not filename:
            return None
        ext = Path(filename).suffix.lower()
        if ext not in OOXML_EXTENSIONS:
            return None
        content = resp.content
        if len(content) < 100:
            return None
        file_hash = hashlib.md5(content).hexdigest()[:12]
        if file_hash in existing_hashes:
            return None
        existing_hashes.add(file_hash)
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
    except:
        return None


def main():
    output_dir = Path("./documents")
    output_dir.mkdir(parents=True, exist_ok=True)
    session = requests.Session()

    # Load existing manifest
    manifest_path = output_dir / "manifest.json"
    existing = []
    existing_hashes = set()
    if manifest_path.exists():
        data = json.loads(manifest_path.read_text())
        existing = data.get("documents", [])
        existing_hashes = {d["hash"] for d in existing}

    collected = list(existing)
    counts = {"docx": 0, "xlsx": 0, "pptx": 0}
    for d in existing:
        counts[d["format"]] = counts.get(d["format"], 0) + 1

    initial = sum(counts.values())
    target = 1000
    print(f"Existing: {initial} documents")
    print(f"Target: {target}")
    print()

    seen_urls = set()

    for idx, seed_url in enumerate(DEEP_URLS):
        if sum(counts.values()) >= target:
            break

        total_so_far = sum(counts.values())
        print(f"[{idx+1}/{len(DEEP_URLS)}] ({total_so_far}/{target}) {seed_url}")

        to_crawl = [seed_url]
        crawled = 0

        while to_crawl and crawled < 10 and sum(counts.values()) < target:
            page_url = to_crawl.pop(0)
            if page_url in seen_urls:
                continue
            seen_urls.add(page_url)
            crawled += 1

            doc_links, sub_pages = find_office_links(page_url, session)
            for sp in sub_pages:
                if sp not in seen_urls:
                    to_crawl.append(sp)

            for doc_url in doc_links:
                if doc_url in seen_urls:
                    continue
                seen_urls.add(doc_url)
                if sum(counts.values()) >= target:
                    break

                meta = download_file(doc_url, output_dir, session, existing_hashes)
                if meta:
                    collected.append(meta)
                    fmt = meta["format"]
                    counts[fmt] = counts.get(fmt, 0) + 1
                    total = sum(counts.values())
                    size_kb = meta["size_bytes"] / 1024
                    print(f"  [{total}/{target}] {fmt} {meta['filename'][:55]} ({size_kb:.0f}KB)")
                time.sleep(0.15)
            time.sleep(0.4)

    # Update manifest
    manifest = {
        "total": sum(counts.values()),
        "counts": counts,
        "documents": collected,
    }
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2))

    added = sum(counts.values()) - initial
    print(f"\nAdded: {added} new documents")
    print(f"Total: {sum(counts.values())} (docx:{counts['docx']} xlsx:{counts['xlsx']} pptx:{counts['pptx']})")


if __name__ == "__main__":
    main()
