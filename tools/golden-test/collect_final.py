#!/usr/bin/env python3
"""
Final push to 500 documents.
Uses known-working URL patterns and deeper sub-page crawling.
"""
import hashlib, json, os, re, sys, time, urllib.parse
from pathlib import Path
import requests
from bs4 import BeautifulSoup

OOXML_EXTENSIONS = {".docx", ".xlsx", ".pptx"}
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"}

# METI confirmed working pattern - seidou stats have xlsx
METI_SEIDOU_URLS = [
    f"https://www.meti.go.jp/statistics/tyo/seidou/result/{year}/08_seidou.html"
    for year in ["r05k", "r04k", "r03k", "r02k", "r01k", "h30k", "h29k", "h28k"]
] + [
    f"https://www.meti.go.jp/statistics/tyo/seidou/result/{year}/02_seidou.html"
    for year in ["r05k", "r04k", "r03k", "r02k", "r01k", "h30k", "h29k"]
]

# METI syoudou (commercial stats)
METI_SYOUDOU_URLS = [
    f"https://www.meti.go.jp/statistics/tyo/syoudou/result/result_{n}.html"
    for n in range(1, 10)
]

# Stat.go.jp - direct xlsx file patterns
STAT_DIRECT_XLSX = []
for prefix in ["05k2", "05k5", "05k2n"]:
    for n in range(1, 20):
        STAT_DIRECT_XLSX.append(f"https://www.stat.go.jp/data/jinsui/zuhyou/{prefix}-{n}.xlsx")

# MHLW chingin - known xlsx source
MHLW_CHINGIN = [
    f"https://www.mhlw.go.jp/toukei/itiran/roudou/chingin/kouzou/z{year}/"
    for year in range(2015, 2024)
]

# Digital Agency data catalogue - yielded pptx and docx
DIGITAL_URLS = [
    "https://www.digital.go.jp/resources/open_data",
    "https://www.digital.go.jp/resources/data_strategy",
    "https://www.digital.go.jp/resources/govcloud",
    "https://www.digital.go.jp/resources/standard_guidelines",
    "https://www.digital.go.jp/resources",
]

# MIC (Soumu) stat pages - xlsx
SOUMU_URLS = [
    f"https://www.soumu.go.jp/menu_news/s-news/01toukei03_01000{n:03d}.html"
    for n in range(100, 130)
]

# MEXT shingi (council) - pptx/docx in meeting materials
MEXT_URLS = [
    f"https://www.mext.go.jp/b_menu/shingi/chousa/shotou/174/shiryo/1422686_000{n:02d}.htm"
    for n in range(1, 20)
] + [
    f"https://www.mext.go.jp/b_menu/shingi/chousa/koutou/116/siryo/mext_000{n:02d}.html"
    for n in range(1, 10)
]

ALL_URLS = METI_SEIDOU_URLS + METI_SYOUDOU_URLS + STAT_DIRECT_XLSX + MHLW_CHINGIN + DIGITAL_URLS + SOUMU_URLS + MEXT_URLS

def find_links(url, session):
    doc_links, sub_pages = [], []
    try:
        resp = session.get(url, headers=HEADERS, timeout=15, allow_redirects=True)
        resp.raise_for_status()
        ct = resp.headers.get("content-type", "")
        if "application/" in ct and "html" not in ct:
            return [url], []
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
    except:
        pass
    return doc_links, sub_pages[:20]

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
    print(f"Target: {target}")
    print(f"URLs to try: {len(ALL_URLS)}")

    seen = set()
    for idx, seed in enumerate(ALL_URLS):
        if sum(counts.values()) >= target:
            break
        total = sum(counts.values())
        if idx % 20 == 0:
            print(f"[{idx+1}/{len(ALL_URLS)}] ({total}/{target})")

        # Direct file URL?
        if any(seed.lower().endswith(ext) for ext in OOXML_EXTENSIONS):
            if seed not in seen:
                seen.add(seed)
                meta = download(seed, output_dir, session, existing_hashes)
                if meta:
                    collected.append(meta)
                    fmt = meta["format"]
                    counts[fmt] = counts.get(fmt, 0) + 1
                    total = sum(counts.values())
                    size_kb = meta["size_bytes"] / 1024
                    print(f"  [{total}/{target}] {fmt} {meta['filename'][:55]} ({size_kb:.0f}KB)")
                time.sleep(0.05)
            continue

        # Crawl page
        to_crawl = [seed]
        crawled = 0
        while to_crawl and crawled < 8 and sum(counts.values()) < target:
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
                time.sleep(0.05)
            time.sleep(0.15)

    manifest = {"total": sum(counts.values()), "counts": counts, "documents": collected}
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2))
    added = sum(counts.values()) - initial
    print(f"\nAdded: {added}")
    print(f"Total: {sum(counts.values())} (docx:{counts.get('docx',0)} xlsx:{counts.get('xlsx',0)} pptx:{counts.get('pptx',0)})")

if __name__ == "__main__":
    main()
