#!/usr/bin/env python3
"""
Collect from verified sources that actually serve OOXML files.
Uses direct file URLs and pages confirmed to have downloadable Office documents.
"""
import hashlib, json, os, re, sys, time, urllib.parse
from pathlib import Path
import requests
from bs4 import BeautifulSoup

OOXML_EXTENSIONS = {".docx", ".xlsx", ".pptx"}
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"}

# Direct xlsx file URLs from stat.go.jp (confirmed working)
DIRECT_FILES = [
    "https://www.stat.go.jp/data/jinsui/zuhyou/05k2-1.xlsx",
    "https://www.stat.go.jp/data/jinsui/zuhyou/05k2-2.xlsx",
    "https://www.stat.go.jp/data/jinsui/zuhyou/05k2-3.xlsx",
    "https://www.stat.go.jp/data/jinsui/zuhyou/05k2-4.xlsx",
    "https://www.stat.go.jp/data/jinsui/zuhyou/05k2-5.xlsx",
    "https://www.stat.go.jp/data/jinsui/zuhyou/05k2-6.xlsx",
    "https://www.stat.go.jp/data/jinsui/zuhyou/05k5-1.xlsx",
    "https://www.stat.go.jp/data/jinsui/zuhyou/05k5-2.xlsx",
    "https://www.stat.go.jp/data/jinsui/zuhyou/05k5-3.xlsx",
    "https://www.stat.go.jp/data/jinsui/zuhyou/05k5-4.xlsx",
    "https://www.stat.go.jp/data/jinsui/zuhyou/05k5-5.xlsx",
    "https://www.stat.go.jp/data/jinsui/zuhyou/05k5-6.xlsx",
    "https://www.stat.go.jp/data/jinsui/zuhyou/05k2n-1.xlsx",
    "https://www.stat.go.jp/data/jinsui/zuhyou/05k2n-2.xlsx",
    "https://www.stat.go.jp/data/jinsui/zuhyou/05k2n-3.xlsx",
    "https://www.stat.go.jp/data/jinsui/zuhyou/05k2n-4.xlsx",
]

# Pages that are known to link to xlsx/docx/pptx downloads
CRAWL_PAGES = [
    # Stat.go.jp - monthly economic stats (xlsx heavy, confirmed)
    "https://www.stat.go.jp/data/jinsui/2.html",
    "https://www.stat.go.jp/data/roudou/sokuhou/nen/dt/index.html",
    "https://www.stat.go.jp/data/roudou/sokuhou/nen/ft/index.html",
    "https://www.stat.go.jp/data/kakei/sokuhou/nen/index.html",
    "https://www.stat.go.jp/data/kouri/doukou/index.html",
    "https://www.stat.go.jp/data/cpi/sokuhou/nen/index-z.html",
    # METI industry stats (xlsx, confirmed working)
    "https://www.meti.go.jp/statistics/tyo/seidou/result/ichiran/08_seidou.html",
    "https://www.meti.go.jp/statistics/tyo/seidou/result/ichiran/02_seidou.html",
    "https://www.meti.go.jp/statistics/tyo/seidou/result/ichiran/01_seidou.html",
    "https://www.meti.go.jp/statistics/tyo/seidou/result/ichiran/03_seidou.html",
    "https://www.meti.go.jp/statistics/tyo/seidou/result/ichiran/04_seidou.html",
    "https://www.meti.go.jp/statistics/tyo/seidou/result/ichiran/05_seidou.html",
    "https://www.meti.go.jp/statistics/tyo/seidou/result/ichiran/06_seidou.html",
    "https://www.meti.go.jp/statistics/tyo/seidou/result/ichiran/07_seidou.html",
    # MOJ docx forms (confirmed)
    "https://www.moj.go.jp/MINJI/minji06_00108.html",
    "https://www.moj.go.jp/MINJI/minji06_00107.html",
    "https://www.moj.go.jp/MINJI/minji06_00104.html",
    "https://www.moj.go.jp/MINJI/minji06_00106.html",
    "https://www.moj.go.jp/MINJI/minji06_00105.html",
    "https://www.moj.go.jp/MINJI/minji05_00343.html",
    "https://www.moj.go.jp/MINJI/minji05_00344.html",
    "https://www.moj.go.jp/MINJI/minji05_00345.html",
    "https://www.moj.go.jp/MINJI/minji05_00346.html",
    "https://www.moj.go.jp/MINJI/minji05_00347.html",
    "https://www.moj.go.jp/MINJI/minji05_00355.html",
    "https://www.moj.go.jp/MINJI/minji05_00356.html",
    # BOJ data files (xlsx)
    "https://www.boj.or.jp/statistics/money/zandaka/zand2501.htm",
    "https://www.boj.or.jp/statistics/money/zandaka/zand2502.htm",
    "https://www.boj.or.jp/statistics/money/zandaka/zand2503.htm",
    "https://www.boj.or.jp/statistics/money/zandaka/zand2504.htm",
    "https://www.boj.or.jp/statistics/money/zandaka/zand2505.htm",
    "https://www.boj.or.jp/statistics/boj/other/acmai/release/2025/ac250228.htm/",
    "https://www.boj.or.jp/statistics/boj/other/acmai/release/2025/ac250131.htm/",
    "https://www.boj.or.jp/statistics/boj/other/acmai/release/2024/ac241231.htm/",
    "https://www.boj.or.jp/statistics/boj/other/acmai/release/2024/ac241130.htm/",
    "https://www.boj.or.jp/statistics/boj/other/acmai/release/2024/ac241031.htm/",
    # MHLW wage stats (xlsx confirmed)
    "https://www.mhlw.go.jp/toukei/itiran/roudou/chingin/kouzou/z2023/",
    "https://www.mhlw.go.jp/toukei/itiran/roudou/chingin/kouzou/z2022/",
    "https://www.mhlw.go.jp/toukei/itiran/roudou/chingin/kouzou/z2021/",
    # MLIT transport data (xlsx)
    "https://www.mlit.go.jp/statistics/details/tetsudo_list.html",
    "https://www.mlit.go.jp/statistics/details/port_list.html",
    "https://www.mlit.go.jp/statistics/details/kensetu_list.html",
    # NTA tax forms (docx)
    "https://www.nta.go.jp/taxes/tetsuzuki/shinsei/annai/hojin/shinkoku/itiran2024/01.htm",
    "https://www.nta.go.jp/taxes/tetsuzuki/shinsei/annai/hojin/shinkoku/itiran2023/01.htm",
    "https://www.nta.go.jp/taxes/tetsuzuki/shinsei/annai/hojin/shinkoku/itiran2022/01.htm",
    "https://www.nta.go.jp/taxes/tetsuzuki/shinsei/annai/gensen/annai/1648_01.htm",
    "https://www.nta.go.jp/taxes/tetsuzuki/shinsei/annai/gensen/annai/1648_73.htm",
    "https://www.nta.go.jp/taxes/tetsuzuki/shinsei/annai/gensen/annai/1648_74.htm",
    # Digital Agency data (xlsx/docx)
    "https://www.digital.go.jp/resources/open_data",
    "https://www.digital.go.jp/policies/mynumber/faq-document",
]

def find_links(url, session):
    doc_links, sub_pages = [], []
    try:
        resp = session.get(url, headers=HEADERS, timeout=15, allow_redirects=True)
        resp.raise_for_status()
        if any(url.lower().endswith(ext) for ext in OOXML_EXTENSIONS):
            return [url], []
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
        print(f"  [err] {url[:60]}: {e}")
    return doc_links, sub_pages[:30]

def download(url, output_dir, session, existing_hashes):
    try:
        resp = session.get(url, headers=HEADERS, timeout=30)
        resp.raise_for_status()
        parsed = urllib.parse.urlparse(url)
        filename = urllib.parse.unquote(Path(parsed.path).name)
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
        return {"filename": safe_name, "source_url": url, "format": ext.lstrip('.'),
                "size_bytes": len(content), "hash": file_hash}
    except:
        return None

def main():
    output_dir = Path("./documents")
    output_dir.mkdir(parents=True, exist_ok=True)
    session = requests.Session()
    manifest_path = output_dir / "manifest.json"
    existing = []
    existing_hashes = set()
    if manifest_path.exists():
        data = json.loads(manifest_path.read_text())
        existing = data.get("documents", [])
        existing_hashes = {d["hash"] for d in existing}
    collected = list(existing)
    counts = {}
    for d in existing:
        counts[d["format"]] = counts.get(d["format"], 0) + 1
    initial = sum(counts.values())
    target = 500
    print(f"Existing: {initial} (docx:{counts.get('docx',0)} xlsx:{counts.get('xlsx',0)} pptx:{counts.get('pptx',0)})")

    # Phase 1: Direct file downloads
    print(f"\n=== Phase 1: Direct file downloads ({len(DIRECT_FILES)} files) ===")
    for url in DIRECT_FILES:
        if sum(counts.values()) >= target:
            break
        meta = download(url, output_dir, session, existing_hashes)
        if meta:
            collected.append(meta)
            fmt = meta["format"]
            counts[fmt] = counts.get(fmt, 0) + 1
            total = sum(counts.values())
            size_kb = meta["size_bytes"] / 1024
            print(f"  [{total}/{target}] {fmt} {meta['filename'][:55]} ({size_kb:.0f}KB)")
        time.sleep(0.1)

    # Phase 2: Crawl pages for links
    print(f"\n=== Phase 2: Crawl pages ({len(CRAWL_PAGES)} seeds) ===")
    seen = set()
    for idx, seed in enumerate(CRAWL_PAGES):
        if sum(counts.values()) >= target:
            break
        total = sum(counts.values())
        print(f"[{idx+1}/{len(CRAWL_PAGES)}] ({total}/{target}) {seed[:70]}")
        to_crawl = [seed]
        crawled = 0
        while to_crawl and crawled < 12 and sum(counts.values()) < target:
            page = to_crawl.pop(0)
            if page in seen:
                continue
            seen.add(page)
            crawled += 1
            doc_links, sub_pages = find_links(page, session)
            for sp in sub_pages:
                if sp not in seen:
                    to_crawl.append(sp)
            for doc_url in doc_links:
                if doc_url in seen:
                    continue
                seen.add(doc_url)
                if sum(counts.values()) >= target:
                    break
                meta = download(doc_url, output_dir, session, existing_hashes)
                if meta:
                    collected.append(meta)
                    fmt = meta["format"]
                    counts[fmt] = counts.get(fmt, 0) + 1
                    total = sum(counts.values())
                    size_kb = meta["size_bytes"] / 1024
                    print(f"  [{total}/{target}] {fmt} {meta['filename'][:55]} ({size_kb:.0f}KB)")
                time.sleep(0.08)
            time.sleep(0.2)

    manifest = {"total": sum(counts.values()), "counts": counts, "documents": collected}
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2))
    added = sum(counts.values()) - initial
    print(f"\nAdded: {added}")
    print(f"Total: {sum(counts.values())} (docx:{counts.get('docx',0)} xlsx:{counts.get('xlsx',0)} pptx:{counts.get('pptx',0)})")

if __name__ == "__main__":
    main()
