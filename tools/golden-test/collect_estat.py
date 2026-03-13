#!/usr/bin/env python3
"""Bulk collect xlsx from e-Stat (government statistics portal) and other data portals."""
import hashlib, json, os, re, sys, time, urllib.parse
from pathlib import Path
import requests
from bs4 import BeautifulSoup

OOXML_EXTENSIONS = {".docx", ".xlsx", ".pptx"}
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}

# e-Stat category pages with xlsx downloads
ESTAT_PAGES = [
    f"https://www.e-stat.go.jp/stat-search/files?page=1&layout=datalist&toukei=00200521&tstat=000001136464&cycle=0&tclass1val=0&stat_infid={id_}"
    for id_ in range(800001, 800050)
]

# More targeted URLs
TARGETED_URLS = [
    # METI industry statistics (xlsx rich)
    "https://www.meti.go.jp/statistics/tyo/syoudou/result/result_1.html",
    "https://www.meti.go.jp/statistics/tyo/syoudou/result/result_2.html",
    "https://www.meti.go.jp/statistics/tyo/syoudou/result/result_3.html",
    "https://www.meti.go.jp/statistics/tyo/seidou/result/ichiran/08_seidou.html",
    "https://www.meti.go.jp/statistics/tyo/seidou/result/ichiran/02_seidou.html",
    # Stat.go.jp (xlsx heavy)
    "https://www.stat.go.jp/data/jinsui/tsuki/index.html",
    "https://www.stat.go.jp/data/roudou/sokuhou/nen/dt/index.html",
    "https://www.stat.go.jp/data/roudou/sokuhou/nen/ft/index.html",
    "https://www.stat.go.jp/data/kakei/sokuhou/nen/index.html",
    "https://www.stat.go.jp/data/kouri/doukou/index.html",
    "https://www.stat.go.jp/data/cpi/sokuhou/nen/index-z.html",
    "https://www.stat.go.jp/data/service/2019/index.html",
    "https://www.stat.go.jp/data/service/2019/zuhyou.html",
    "https://www.stat.go.jp/data/jinsui/zuhyou/05k2-1.xlsx",
    # MOF budget data (xlsx)
    "https://www.mof.go.jp/policy/budget/budger_workflow/account/fy2023/",
    "https://www.mof.go.jp/policy/budget/budger_workflow/account/fy2022/",
    "https://www.mof.go.jp/policy/budget/budger_workflow/account/fy2024/",
    # BOJ data (xlsx heavy)
    "https://www.boj.or.jp/statistics/money/zandaka/zand2501.htm",
    "https://www.boj.or.jp/statistics/money/zandaka/zand2502.htm",
    "https://www.boj.or.jp/statistics/money/zandaka/zand2503.htm",
    "https://www.boj.or.jp/statistics/boj/other/acmai/release/2025/ac250228.htm/",
    "https://www.boj.or.jp/statistics/boj/other/acmai/release/2025/ac250131.htm/",
    # NTA forms (docx)
    "https://www.nta.go.jp/taxes/tetsuzuki/shinsei/annai/hojin/shinkoku/itiran2024/01.htm",
    "https://www.nta.go.jp/taxes/tetsuzuki/shinsei/annai/hojin/shinkoku/itiran2023/01.htm",
    "https://www.nta.go.jp/taxes/tetsuzuki/shinsei/annai/hojin/shinkoku/itiran2022/01.htm",
    "https://www.nta.go.jp/taxes/tetsuzuki/shinsei/annai/gensen/annai/1648_01.htm",
    "https://www.nta.go.jp/taxes/tetsuzuki/shinsei/annai/gensen/annai/1648_73.htm",
    # MHLW forms and data
    "https://www.mhlw.go.jp/toukei/list/chinginkouzou.html",
    "https://www.mhlw.go.jp/toukei/itiran/roudou/chingin/kouzou/z2023/",
    "https://www.mhlw.go.jp/toukei/itiran/roudou/chingin/kouzou/z2022/",
    "https://www.mhlw.go.jp/toukei/list/h22-46-50.html",
    # MOJ forms (docx)
    "https://www.moj.go.jp/MINJI/minji06_00108.html",
    "https://www.moj.go.jp/MINJI/minji06_00107.html",
    "https://www.moj.go.jp/MINJI/minji06_00104.html",
    "https://www.moj.go.jp/MINJI/minji06_00106.html",
    "https://www.moj.go.jp/MINJI/minji06_00105.html",
    "https://www.moj.go.jp/MINJI/minji05_00343.html",
    "https://www.moj.go.jp/MINJI/minji05_00344.html",
    "https://www.moj.go.jp/MINJI/minji05_00345.html",
    # MLIT data downloads
    "https://www.mlit.go.jp/statistics/details/tetsudo_list.html",
    "https://www.mlit.go.jp/statistics/details/kensetu_list.html",
    "https://www.mlit.go.jp/statistics/details/port_list.html",
    # Digital agency
    "https://www.digital.go.jp/policies/mynumber/",
    "https://www.digital.go.jp/policies/data_strategy/",
]

def find_links(url, session):
    doc_links, sub_pages = [], []
    try:
        resp = session.get(url, headers=HEADERS, timeout=15, allow_redirects=True)
        resp.raise_for_status()
        # Check if URL itself is an xlsx
        if url.endswith(('.xlsx', '.docx', '.pptx')):
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

    all_urls = TARGETED_URLS
    seen = set()
    for idx, seed in enumerate(all_urls):
        if sum(counts.values()) >= target:
            break
        total = sum(counts.values())
        if idx % 10 == 0:
            print(f"[{idx+1}/{len(all_urls)}] ({total}/{target})")
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
