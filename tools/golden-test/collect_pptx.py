#!/usr/bin/env python3
"""Collect .pptx files specifically from government sites."""
import hashlib, json, os, re, sys, time, urllib.parse
from pathlib import Path
import requests
from bs4 import BeautifulSoup

OOXML_EXTENSIONS = {".pptx", ".docx", ".xlsx"}
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}

# URLs known to have PowerPoint files
PPTX_URLS = [
    # Government councils/committees often use PPTX
    "https://www.meti.go.jp/shingikai/mono_info_service/digital_jinzai/",
    "https://www.meti.go.jp/shingikai/sankoshin/shin_kijun/",
    "https://www.meti.go.jp/shingikai/economy_and_industry/",
    "https://www.soumu.go.jp/main_sosiki/kenkyu/",
    "https://www.soumu.go.jp/main_sosiki/joho_tsusin/policyreports/",
    "https://www.mhlw.go.jp/stf/shingi/shingi-hosho_126706.html",
    "https://www.mhlw.go.jp/stf/shingi/shingi-rousei_126734.html",
    "https://www.mhlw.go.jp/stf/shingi/other-syokuhin_128957.html",
    "https://www.mlit.go.jp/policy/shingikai/",
    "https://www.env.go.jp/council/05haikibunkai/",
    "https://www.digital.go.jp/councils/",
    "https://www.cao.go.jp/council/",
    "https://www.fsa.go.jp/singi/",
    "https://www.nta.go.jp/about/council/",
    "https://www.mext.go.jp/b_menu/shingi/",
    "https://www.maff.go.jp/j/council/",
    # Research institutes
    "https://www.nistep.go.jp/archives/category/research-report",
    "https://www.ipa.go.jp/security/reports/",
    "https://www.ipa.go.jp/jinzai/",
    # Prefectures
    "https://www.pref.kanagawa.jp/docs/r5k/",
    "https://www.pref.saitama.lg.jp/a0001/",
    "https://www.pref.chiba.lg.jp/seisaku/",
    "https://www.city.nagoya.jp/shisei/",
    "https://www.city.sapporo.jp/kikaku/",
]

def find_links(url, session):
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
    print(f"Existing: {initial} (docx:{counts.get('docx',0)} xlsx:{counts.get('xlsx',0)} pptx:{counts.get('pptx',0)})")
    seen = set()
    for idx, seed in enumerate(PPTX_URLS):
        total = sum(counts.values())
        print(f"[{idx+1}/{len(PPTX_URLS)}] ({total}) {seed}")
        to_crawl = [seed]
        crawled = 0
        while to_crawl and crawled < 12:
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
                meta = download(doc_url, output_dir, session, existing_hashes)
                if meta:
                    collected.append(meta)
                    fmt = meta["format"]
                    counts[fmt] = counts.get(fmt, 0) + 1
                    total = sum(counts.values())
                    size_kb = meta["size_bytes"] / 1024
                    print(f"  [{total}] {fmt} {meta['filename'][:55]} ({size_kb:.0f}KB)")
                time.sleep(0.15)
            time.sleep(0.4)
    manifest = {"total": sum(counts.values()), "counts": counts, "documents": collected}
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2))
    added = sum(counts.values()) - initial
    print(f"\nAdded: {added} new documents")
    print(f"Total: {sum(counts.values())} (docx:{counts.get('docx',0)} xlsx:{counts.get('xlsx',0)} pptx:{counts.get('pptx',0)})")

if __name__ == "__main__":
    main()
