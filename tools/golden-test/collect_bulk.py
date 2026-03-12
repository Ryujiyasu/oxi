#!/usr/bin/env python3
"""Bulk collect by crawling deeper into sites that already yielded results."""
import hashlib, json, os, re, sys, time, urllib.parse
from pathlib import Path
import requests
from bs4 import BeautifulSoup

OOXML_EXTENSIONS = {".docx", ".xlsx", ".pptx"}
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}

# Sites that yielded results + deeper pages
BULK_URLS = [
    # IPA - more threat reports (pptx!)
    "https://www.ipa.go.jp/security/10threats/ps6vr700000080kf-att/kaisetsu_2025.pdf",
    "https://www.ipa.go.jp/security/10threats/10threats2024.html",
    "https://www.ipa.go.jp/security/10threats/10threats2023.html",
    "https://www.ipa.go.jp/security/guide/sme/",
    "https://www.ipa.go.jp/shiken/",
    "https://www.ipa.go.jp/digital/chousa/wps/",
    # Fukuoka pref - yielded docx
    "https://www.pref.fukuoka.lg.jp/life/1/",
    "https://www.pref.fukuoka.lg.jp/life/2/",
    "https://www.pref.fukuoka.lg.jp/life/3/",
    "https://www.pref.fukuoka.lg.jp/life/4/",
    "https://www.pref.fukuoka.lg.jp/life/5/",
    "https://www.pref.fukuoka.lg.jp/life/6/",
    "https://www.pref.fukuoka.lg.jp/life/7/",
    "https://www.pref.fukuoka.lg.jp/life/8/",
    # MAFF deeper
    "https://www.maff.go.jp/j/press/kanbo/anpo/",
    "https://www.maff.go.jp/j/press/shokuhin/kaigai/",
    "https://www.maff.go.jp/j/press/nousin/noukei/",
    "https://www.maff.go.jp/j/press/seisan/engei/",
    "https://www.maff.go.jp/j/budget/",
    "https://www.maff.go.jp/j/kanbo/joho/",
    # MHLW - deeper shingi pages
    "https://www.mhlw.go.jp/stf/shingi/shingi-hosho_126714.html",
    "https://www.mhlw.go.jp/stf/shingi/shingi-rousei_126748.html",
    "https://www.mhlw.go.jp/stf/shingi/other-roudou_128790.html",
    "https://www.mhlw.go.jp/stf/seisakunitsuite/bunya/0000121431_00396.html",
    "https://www.mhlw.go.jp/stf/newpage_20412.html",
    "https://www.mhlw.go.jp/stf/seisakunitsuite/bunya/hukushi_kaigo/kaigo_koureisha/",
    # METI deeper
    "https://www.meti.go.jp/statistics/tyo/syoudou/result/result_1.html",
    "https://www.meti.go.jp/shingikai/enecho/denryoku_gas/",
    "https://www.meti.go.jp/shingikai/sankoshin/seizo_sangyo/",
    "https://www.meti.go.jp/statistics/tyo/kougyo/result-2.html",
    # NTA deeper
    "https://www.nta.go.jp/taxes/tetsuzuki/shinsei/",
    "https://www.nta.go.jp/taxes/shiraberu/taxanswer/",
    "https://www.nta.go.jp/about/organization/ntc/kenkyu/",
    # MLIT deeper
    "https://www.mlit.go.jp/jutakukentiku/house/jutakukentiku_house_tk3_000015.html",
    "https://www.mlit.go.jp/road/ir/",
    "https://www.mlit.go.jp/sogoseisaku/transport/",
    "https://www.mlit.go.jp/statistics/details/kensetu_list.html",
    # Cabinet Secretariat deeper
    "https://www.cas.go.jp/jp/seisaku/atarashii_sihonshugi/",
    "https://www.cas.go.jp/jp/seisaku/digital_denen/",
    # MOF deeper
    "https://www.mof.go.jp/policy/budget/budger_workflow/budget/",
    "https://www.mof.go.jp/policy/international_policy/",
    # Hokkaido deeper
    "https://www.pref.hokkaido.lg.jp/ss/tkk/kaikaku/",
    "https://www.pref.hokkaido.lg.jp/kz/ssg/",
    # Tokyo metro deeper
    "https://www.metro.tokyo.lg.jp/tosei/hodohappyo/press/2024/01/",
    "https://www.metro.tokyo.lg.jp/tosei/hodohappyo/press/2024/02/",
    "https://www.metro.tokyo.lg.jp/tosei/hodohappyo/press/2024/03/",
    "https://www.metro.tokyo.lg.jp/tosei/hodohappyo/press/2024/04/",
    "https://www.metro.tokyo.lg.jp/tosei/hodohappyo/press/2024/05/",
    "https://www.metro.tokyo.lg.jp/tosei/hodohappyo/press/2024/06/",
    # Saitama deeper
    "https://www.pref.saitama.lg.jp/a0001/news/",
    "https://www.pref.saitama.lg.jp/a0301/",
    # Chiba deeper
    "https://www.pref.chiba.lg.jp/seisaku/shingikai/",
    "https://www.pref.chiba.lg.jp/seihou/",
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
    return doc_links, sub_pages[:40]

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
    print(f"Target: {target}")
    seen = set()
    for idx, seed in enumerate(BULK_URLS):
        if sum(counts.values()) >= target:
            break
        total = sum(counts.values())
        print(f"[{idx+1}/{len(BULK_URLS)}] ({total}/{target}) {seed}")
        to_crawl = [seed]
        crawled = 0
        while to_crawl and crawled < 20 and sum(counts.values()) < target:
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
            time.sleep(0.25)
    manifest = {"total": sum(counts.values()), "counts": counts, "documents": collected}
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2))
    added = sum(counts.values()) - initial
    print(f"\nAdded: {added}")
    print(f"Total: {sum(counts.values())} (docx:{counts.get('docx',0)} xlsx:{counts.get('xlsx',0)} pptx:{counts.get('pptx',0)})")

if __name__ == "__main__":
    main()
