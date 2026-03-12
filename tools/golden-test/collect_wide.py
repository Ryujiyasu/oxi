#!/usr/bin/env python3
"""
Collect Office documents from diverse sources (not limited to government).
Targets: universities, research institutes, companies, international orgs.
"""
import hashlib, json, os, re, sys, time, urllib.parse
from pathlib import Path
import requests
from bs4 import BeautifulSoup

OOXML_EXTENSIONS = {".docx", ".xlsx", ".pptx"}
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}

WIDE_URLS = [
    # Universities (often have pptx in course materials)
    "https://www.u-tokyo.ac.jp/ja/research/",
    "https://www.kyoto-u.ac.jp/ja/research/",
    "https://www.waseda.jp/top/news",
    "https://www.keio.ac.jp/ja/news/",
    # IPA (Information-technology Promotion Agency) - many pptx
    "https://www.ipa.go.jp/security/vuln/",
    "https://www.ipa.go.jp/security/10threats/10threats2025.html",
    "https://www.ipa.go.jp/jinzai/skill-standard/",
    "https://www.ipa.go.jp/publish/wp-ict/",
    # JICA - international cooperation reports
    "https://www.jica.go.jp/Resource/activities/schemes/finance_co/loan/standard/",
    # JETRO - trade reports
    "https://www.jetro.go.jp/world/reports/",
    "https://www.jetro.go.jp/ext_images/",
    # NEDO - research reports
    "https://www.nedo.go.jp/activities/",
    # JST - science reports
    "https://www.jst.go.jp/report/",
    # BOJ - Bank of Japan (xlsx heavy)
    "https://www.boj.or.jp/statistics/dl/",
    "https://www.boj.or.jp/research/wps_rev/",
    # Ministry of Defense
    "https://www.mod.go.jp/j/approach/agenda/meeting/",
    # National Diet Library
    "https://www.ndl.go.jp/jp/diet/publication/",
    # Corporations - IR materials (pptx!)
    "https://www.softbank.jp/corp/ir/documents/presentations/",
    "https://www.toyota-global.com/investors/ir_library/",
    "https://www.ntt.co.jp/ir/library/",
    "https://www.sony.com/ja/SonyInfo/IR/library/",
    # Open data portals
    "https://catalog.data.metro.tokyo.lg.jp/dataset/",
    "https://data.e-stat.go.jp/",
    # NHK open data
    "https://www.nhk.or.jp/bunken/research/",
    # Prefectures - deeper crawl for pptx
    "https://www.pref.hokkaido.lg.jp/ss/tkk/",
    "https://www.pref.fukuoka.lg.jp/life/",
    "https://www.pref.hiroshima.lg.jp/soshiki/",
    "https://www.pref.miyagi.jp/soshiki/",
    # Local government open data
    "https://opendata.pref.shizuoka.jp/",
    "https://www.city.kobe.lg.jp/a89138/shise/opendata/",
    # Academic societies
    "https://www.ipsj.or.jp/",
    "https://www.ieice.org/",
    # International orgs (Japan offices)
    "https://www.worldbank.org/ja/country/japan/publication/",
    "https://www.oecd.org/japan/",
    # MEXT - education research (pptx in meetings)
    "https://www.mext.go.jp/b_menu/shingi/chousa/shotou/",
    "https://www.mext.go.jp/b_menu/shingi/gijyutu/",
    # Cabinet Secretariat
    "https://www.cas.go.jp/jp/seisaku/",
    # Reconstruction Agency
    "https://www.reconstruction.go.jp/topics/main-cat1/",
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
            elif parsed.netloc == base_domain and ext in ("", ".html", ".htm", ".php"):
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
    target = 1000
    print(f"Existing: {initial} (docx:{counts.get('docx',0)} xlsx:{counts.get('xlsx',0)} pptx:{counts.get('pptx',0)})")
    print(f"Target: {target}")
    seen = set()
    for idx, seed in enumerate(WIDE_URLS):
        if sum(counts.values()) >= target:
            break
        total = sum(counts.values())
        print(f"[{idx+1}/{len(WIDE_URLS)}] ({total}/{target}) {seed}")
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
                time.sleep(0.1)
            time.sleep(0.3)
    manifest = {"total": sum(counts.values()), "counts": counts, "documents": collected}
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2))
    added = sum(counts.values()) - initial
    print(f"\nAdded: {added}")
    print(f"Total: {sum(counts.values())} (docx:{counts.get('docx',0)} xlsx:{counts.get('xlsx',0)} pptx:{counts.get('pptx',0)})")

if __name__ == "__main__":
    main()
