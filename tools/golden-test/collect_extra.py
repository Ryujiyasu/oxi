#!/usr/bin/env python3
"""
Extra collection - deeper crawl with more sub-page exploration.
Focus on pages that actually serve xlsx/docx/pptx files.
"""
import hashlib, json, os, re, sys, time, urllib.parse
from pathlib import Path
import requests
from bs4 import BeautifulSoup

OOXML_EXTENSIONS = {".docx", ".xlsx", ".pptx"}
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}

# Verified working URLs from previous runs + new deep pages
EXTRA_URLS = [
    # METI - confirmed xlsx source
    "https://www.meti.go.jp/statistics/tyo/syoudou/result/result_1.html",
    "https://www.meti.go.jp/statistics/tyo/syoudou/result/result_2.html",
    "https://www.meti.go.jp/statistics/tyo/syoudou/result/result_3.html",
    "https://www.meti.go.jp/statistics/tyo/seidou/result/ichiran/08_seidou.html",
    "https://www.meti.go.jp/statistics/tyo/seidou/result/ichiran/02_seidou.html",
    "https://www.meti.go.jp/statistics/tyo/seidou/result/ichiran/01_seidou.html",
    "https://www.meti.go.jp/statistics/tyo/seidou/result/ichiran/03_seidou.html",
    "https://www.meti.go.jp/statistics/tyo/seidou/result/ichiran/04_seidou.html",
    "https://www.meti.go.jp/statistics/tyo/seidou/result/ichiran/05_seidou.html",
    "https://www.meti.go.jp/statistics/tyo/seidou/result/ichiran/06_seidou.html",
    "https://www.meti.go.jp/statistics/tyo/seidou/result/ichiran/07_seidou.html",
    "https://www.meti.go.jp/statistics/tyo/tokusabido/result/result_1.html",
    "https://www.meti.go.jp/statistics/tyo/tokusabido/result/result_2.html",
    "https://www.meti.go.jp/statistics/tyo/tokusabido/result/result_3.html",
    # Stat.go.jp - confirmed xlsx
    "https://www.stat.go.jp/data/jinsui/2.html",
    "https://www.stat.go.jp/data/jinsui/tsuki/index.html",
    "https://www.stat.go.jp/data/roudou/sokuhou/tsuki/index.html",
    "https://www.stat.go.jp/data/kakei/sokuhou/tsuki/index.html",
    "https://www.stat.go.jp/data/cpi/sokuhou/tsuki/index-z.html",
    "https://www.stat.go.jp/data/cpi/sokuhou/tsuki/index-t.html",
    "https://www.stat.go.jp/data/koukei/index.html",
    "https://www.stat.go.jp/data/jinsui/new.html",
    "https://www.stat.go.jp/data/kouri/doukou/index.html",
    "https://www.stat.go.jp/data/service/2019/index.html",
    "https://www.stat.go.jp/data/service/2019/zuhyou.html",
    # BOJ - xlsx confirmed
    "https://www.boj.or.jp/statistics/money/ms/index.htm",
    "https://www.boj.or.jp/statistics/money/zandaka/index.htm",
    "https://www.boj.or.jp/statistics/tk/gaiyo/index.htm",
    "https://www.boj.or.jp/statistics/boj/other/acmai/index.htm",
    "https://www.boj.or.jp/statistics/dl/depo/tento/index.htm",
    "https://www.boj.or.jp/statistics/dl/loan/ldo/index.htm",
    # NTA - docx forms confirmed
    "https://www.nta.go.jp/taxes/tetsuzuki/shinsei/annai/hojin/shinkoku/itiran2024/01.htm",
    "https://www.nta.go.jp/taxes/tetsuzuki/shinsei/annai/hojin/shinkoku/itiran2023/01.htm",
    "https://www.nta.go.jp/taxes/tetsuzuki/shinsei/annai/hojin/shinkoku/itiran2022/01.htm",
    "https://www.nta.go.jp/taxes/tetsuzuki/shinsei/annai/gensen/annai/1648_01.htm",
    "https://www.nta.go.jp/taxes/tetsuzuki/shinsei/annai/gensen/annai/1648_73.htm",
    "https://www.nta.go.jp/taxes/tetsuzuki/shinsei/annai/hojin/annai/1554_2.htm",
    "https://www.nta.go.jp/taxes/tetsuzuki/shinsei/annai/hojin/annai/5100.htm",
    "https://www.nta.go.jp/taxes/tetsuzuki/shinsei/annai/hojin/annai/01.htm",
    "https://www.nta.go.jp/taxes/tetsuzuki/shinsei/annai/hojin/annai/5101.htm",
    "https://www.nta.go.jp/taxes/tetsuzuki/shinsei/annai/hojin/annai/5102.htm",
    # MOJ - docx forms confirmed
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
    # MHLW - xlsx wage data
    "https://www.mhlw.go.jp/toukei/itiran/roudou/chingin/kouzou/z2023/",
    "https://www.mhlw.go.jp/toukei/itiran/roudou/chingin/kouzou/z2022/",
    "https://www.mhlw.go.jp/toukei/itiran/roudou/chingin/kouzou/z2021/",
    "https://www.mhlw.go.jp/toukei/itiran/roudou/chingin/kouzou/z2020/",
    "https://www.mhlw.go.jp/toukei/list/chinginkouzou.html",
    "https://www.mhlw.go.jp/toukei/list/h22-46-50.html",
    # MLIT - confirmed xlsx source
    "https://www.mlit.go.jp/statistics/details/tetsudo_list.html",
    "https://www.mlit.go.jp/statistics/details/port_list.html",
    "https://www.mlit.go.jp/statistics/details/kensetu_list.html",
    "https://www.mlit.go.jp/statistics/details/tochi_fudousan_list.html",
    "https://www.mlit.go.jp/statistics/details/seibi_list.html",
    # Soumu - statistical data
    "https://www.soumu.go.jp/menu_seisaku/hakusyo/",
    "https://www.soumu.go.jp/menu_news/s-news/01toukei03_01000123.html",
    "https://www.soumu.go.jp/menu_news/s-news/01toukei03_01000124.html",
    # Digital Agency
    "https://www.digital.go.jp/policies/mynumber/faq-document",
    "https://www.digital.go.jp/resources/open_data",
    # MAFF - deeper pages with xlsx
    "https://www.maff.go.jp/j/tokei/kouhyou/sakumotu/sakkyou_kome/",
    "https://www.maff.go.jp/j/tokei/kouhyou/sakumotu/menseki/",
    "https://www.maff.go.jp/j/tokei/kouhyou/noukei/",
    "https://www.maff.go.jp/j/tokei/kouhyou/kensaku/bunya1.html",
    "https://www.maff.go.jp/j/tokei/kouhyou/kensaku/bunya2.html",
    "https://www.maff.go.jp/j/tokei/kouhyou/kensaku/bunya3.html",
    "https://www.maff.go.jp/j/tokei/kouhyou/kensaku/bunya4.html",
    "https://www.maff.go.jp/j/tokei/kouhyou/kensaku/bunya5.html",
    # Fukuoka (confirmed docx source)
    "https://www.pref.fukuoka.lg.jp/life/1/",
    "https://www.pref.fukuoka.lg.jp/life/2/",
    "https://www.pref.fukuoka.lg.jp/life/3/",
    "https://www.pref.fukuoka.lg.jp/life/4/",
    "https://www.pref.fukuoka.lg.jp/life/5/",
    "https://www.pref.fukuoka.lg.jp/life/6/",
    "https://www.pref.fukuoka.lg.jp/life/7/",
    "https://www.pref.fukuoka.lg.jp/life/8/",
    # e-Stat direct xlsx downloads
    "https://www.e-stat.go.jp/stat-search/files?page=1&layout=datalist&toukei=00200521",
    "https://www.e-stat.go.jp/stat-search/files?page=1&layout=datalist&toukei=00200522",
    "https://www.e-stat.go.jp/stat-search/files?page=1&layout=datalist&toukei=00200524",
    "https://www.e-stat.go.jp/stat-search/files?page=1&layout=datalist&toukei=00200531",
    "https://www.e-stat.go.jp/stat-search/files?page=1&layout=datalist&toukei=00200532",
    "https://www.e-stat.go.jp/stat-search/files?page=1&layout=datalist&toukei=00200541",
    "https://www.e-stat.go.jp/stat-search/files?page=1&layout=datalist&toukei=00200543",
    "https://www.e-stat.go.jp/stat-search/files?page=1&layout=datalist&toukei=00200544",
    "https://www.e-stat.go.jp/stat-search/files?page=1&layout=datalist&toukei=00200551",
    "https://www.e-stat.go.jp/stat-search/files?page=1&layout=datalist&toukei=00200552",
    "https://www.e-stat.go.jp/stat-search/files?page=1&layout=datalist&toukei=00200561",
    "https://www.e-stat.go.jp/stat-search/files?page=1&layout=datalist&toukei=00200571",
    "https://www.e-stat.go.jp/stat-search/files?page=1&layout=datalist&toukei=00200573",
]

def find_links(url, session):
    doc_links, sub_pages = [], []
    try:
        resp = session.get(url, headers=HEADERS, timeout=15, allow_redirects=True)
        resp.raise_for_status()
        if url.lower().endswith(('.xlsx', '.docx', '.pptx')):
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
    return doc_links, sub_pages[:35]

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
    seen = set()
    for idx, seed in enumerate(EXTRA_URLS):
        if sum(counts.values()) >= target:
            break
        total = sum(counts.values())
        if idx % 10 == 0:
            print(f"[{idx+1}/{len(EXTRA_URLS)}] ({total}/{target})")
        to_crawl = [seed]
        crawled = 0
        while to_crawl and crawled < 15 and sum(counts.values()) < target:
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
