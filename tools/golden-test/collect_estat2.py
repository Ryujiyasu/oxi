#!/usr/bin/env python3
"""
Collect more xlsx from e-Stat API and various stat pages.
e-Stat provides direct download links for statistical tables.
"""
import hashlib, json, os, re, sys, time, urllib.parse
from pathlib import Path
import requests
from bs4 import BeautifulSoup

OOXML_EXTENSIONS = {".docx", ".xlsx", ".pptx"}
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"}

# e-Stat file search pages - these list downloadable xlsx files
ESTAT_SEARCH_PAGES = [
    # Population census
    f"https://www.e-stat.go.jp/stat-search/files?page={p}&layout=datalist&toukei=00200521&tstat=000001136464&cycle=0&tclass1val=0"
    for p in range(1, 20)
] + [
    # Labour force survey
    f"https://www.e-stat.go.jp/stat-search/files?page={p}&layout=datalist&toukei=00200531&tstat=000000110001&cycle=1"
    for p in range(1, 10)
] + [
    # CPI
    f"https://www.e-stat.go.jp/stat-search/files?page={p}&layout=datalist&toukei=00200573&tstat=000001084976&cycle=1"
    for p in range(1, 10)
] + [
    # Family income
    f"https://www.e-stat.go.jp/stat-search/files?page={p}&layout=datalist&toukei=00200561&tstat=000000330001&cycle=1"
    for p in range(1, 10)
] + [
    # Housing and land survey
    f"https://www.e-stat.go.jp/stat-search/files?page={p}&layout=datalist&toukei=00200522&tstat=000001127155&cycle=0"
    for p in range(1, 10)
]

def find_xlsx_links(url, session):
    """Find xlsx download links on e-Stat pages."""
    links = []
    try:
        resp = session.get(url, headers=HEADERS, timeout=15, allow_redirects=True)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "html.parser")
        for a in soup.find_all("a", href=True):
            href = a["href"].strip()
            abs_url = urllib.parse.urljoin(url, href)
            parsed = urllib.parse.urlparse(abs_url)
            ext = Path(parsed.path).suffix.lower()
            if ext in OOXML_EXTENSIONS:
                links.append(abs_url)
    except:
        pass
    return links

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
    for idx, page_url in enumerate(ESTAT_SEARCH_PAGES):
        if sum(counts.values()) >= target:
            break
        total = sum(counts.values())
        print(f"[{idx+1}/{len(ESTAT_SEARCH_PAGES)}] ({total}/{target})")
        links = find_xlsx_links(page_url, session)
        print(f"  Found {len(links)} file links")
        for doc_url in links:
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
        time.sleep(0.3)

    manifest = {"total": sum(counts.values()), "counts": counts, "documents": collected}
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2))
    added = sum(counts.values()) - initial
    print(f"\nAdded: {added}")
    print(f"Total: {sum(counts.values())} (docx:{counts.get('docx',0)} xlsx:{counts.get('xlsx',0)} pptx:{counts.get('pptx',0)})")

if __name__ == "__main__":
    main()
