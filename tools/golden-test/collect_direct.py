#!/usr/bin/env python3
"""
Direct download of known Office document URLs.
Focuses on docx and pptx which are underrepresented.
"""
import hashlib, json, os, re, sys, time, urllib.parse
from pathlib import Path
import requests
from bs4 import BeautifulSoup

OOXML_EXTENSIONS = {".docx", ".xlsx", ".pptx"}
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}

# Pages known to have many docx/pptx links
PAGES_WITH_DOCS = [
    # MEXT - council meetings (pptx heavy)
    "https://www.mext.go.jp/b_menu/shingi/chousa/shotou/174/shiryo/1422686_00003.htm",
    "https://www.mext.go.jp/b_menu/shingi/chousa/shotou/174/shiryo/1422686_00004.htm",
    "https://www.mext.go.jp/b_menu/shingi/chousa/shotou/174/shiryo/1422686_00005.htm",
    "https://www.mext.go.jp/b_menu/shingi/chousa/koutou/116/siryo/mext_00001.html",
    "https://www.mext.go.jp/b_menu/shingi/chousa/koutou/116/siryo/mext_00002.html",
    "https://www.mext.go.jp/b_menu/shingi/chousa/koutou/116/siryo/mext_00003.html",
    "https://www.mext.go.jp/b_menu/shingi/chousa/koutou/117/siryo/mext_00001.html",
    "https://www.mext.go.jp/b_menu/shingi/chousa/koutou/117/siryo/mext_00002.html",
    # MHLW - forms and templates (docx)
    "https://www.mhlw.go.jp/stf/seisakunitsuite/bunya/koyou_roudou/roudoukijun/zigyonushi/model/index.html",
    "https://www.mhlw.go.jp/bunya/roudoukijun/roudoujouken01/",
    "https://www.mhlw.go.jp/bunya/roudoukijun/roudoujouken02/",
    "https://www.mhlw.go.jp/stf/seisakunitsuite/bunya/0000187842.html",
    "https://www.mhlw.go.jp/stf/newpage_13432.html",
    "https://www.mhlw.go.jp/stf/newpage_20412.html",
    # Cabinet Office - council materials (pptx)
    "https://www.cas.go.jp/jp/seisaku/atarashii_sihonshugi/kaigi/dai26/gijisidai.html",
    "https://www.cas.go.jp/jp/seisaku/atarashii_sihonshugi/kaigi/dai25/gijisidai.html",
    "https://www.cas.go.jp/jp/seisaku/atarashii_sihonshugi/kaigi/dai24/gijisidai.html",
    "https://www.cas.go.jp/jp/seisaku/digital_denen/dai18/gijisidai.html",
    "https://www.cas.go.jp/jp/seisaku/digital_denen/dai17/gijisidai.html",
    "https://www.cas.go.jp/jp/seisaku/digital_denen/dai16/gijisidai.html",
    # FSA - financial data (xlsx heavy)
    "https://www.fsa.go.jp/singi/singi_kinyu/tosin/",
    "https://www.fsa.go.jp/policy/nisa2/about/tsumitate/target_fund_230630.html",
    # MOE - environment reports
    "https://www.env.go.jp/policy/hakusyo/",
    "https://www.env.go.jp/press/list/",
    "https://www.env.go.jp/council/09water/",
    # Reconstruction Agency - recovery reports
    "https://www.reconstruction.go.jp/topics/main-cat1/sub-cat1-1/",
    "https://www.reconstruction.go.jp/topics/main-cat1/sub-cat1-2/",
    "https://www.reconstruction.go.jp/topics/main-cat1/sub-cat1-3/",
    # Prefectures - forms and docs
    "https://www.pref.aichi.jp/soshiki/zeimu/0000002069.html",
    "https://www.pref.aichi.jp/soshiki/zeimu/0000008180.html",
    "https://www.pref.kanagawa.jp/osirase/0602/form/",
    "https://www.pref.saitama.lg.jp/a0301/yoshikishu/",
    "https://www.pref.chiba.lg.jp/kenfuku/kenko/",
    # City governments
    "https://www.city.nagoya.jp/zaisei/page/0000004097.html",
    "https://www.city.osaka.lg.jp/zaisei/page/0000006918.html",
    "https://www.city.sapporo.jp/zaisei/kohyo/yosan/",
    "https://www.city.fukuoka.lg.jp/zaisei/zaisei/shisei/",
    # MAFF deeper - agriculture data (xlsx + docx)
    "https://www.maff.go.jp/j/tokei/kouhyou/sakumotu/sakkyou_kome/",
    "https://www.maff.go.jp/j/tokei/kouhyou/sakumotu/menseki/",
    "https://www.maff.go.jp/j/tokei/kouhyou/noukei/",
    "https://www.maff.go.jp/j/tokei/kouhyou/kensaku/bunya1.html",
    "https://www.maff.go.jp/j/tokei/kouhyou/kensaku/bunya2.html",
    "https://www.maff.go.jp/j/tokei/kouhyou/kensaku/bunya3.html",
    # BOJ data (xlsx)
    "https://www.boj.or.jp/statistics/money/ms/index.htm",
    "https://www.boj.or.jp/statistics/money/zandaka/index.htm",
    "https://www.boj.or.jp/statistics/tk/gaiyo/index.htm",
    # Stat.go.jp - more xlsx
    "https://www.stat.go.jp/data/jinsui/tsuki/index.html",
    "https://www.stat.go.jp/data/roudou/sokuhou/tsuki/index.html",
    "https://www.stat.go.jp/data/kakei/sokuhou/tsuki/index.html",
    "https://www.stat.go.jp/data/cpi/sokuhou/tsuki/index-z.html",
    "https://www.stat.go.jp/data/cpi/sokuhou/tsuki/index-t.html",
    "https://www.stat.go.jp/data/koukei/index.html",
    "https://www.stat.go.jp/data/jinsui/new.html",
    # More METI
    "https://www.meti.go.jp/statistics/tyo/syoudou/result/result_1.html",
    "https://www.meti.go.jp/statistics/tyo/syoudou/result/result_2.html",
    "https://www.meti.go.jp/statistics/tyo/syoudou/result/result_3.html",
    "https://www.meti.go.jp/statistics/tyo/kougyo/result-2/r04/kakuho/index.html",
    "https://www.meti.go.jp/statistics/tyo/tokusabido/result/result_1.html",
    # NTA - tax forms (docx)
    "https://www.nta.go.jp/taxes/tetsuzuki/shinsei/annai/hojin/annai/1554_2.htm",
    "https://www.nta.go.jp/taxes/tetsuzuki/shinsei/annai/hojin/annai/01.htm",
    "https://www.nta.go.jp/taxes/tetsuzuki/shinsei/annai/hojin/annai/5100.htm",
    "https://www.nta.go.jp/taxes/tetsuzuki/shinsei/annai/gensen/annai/1648_01.htm",
    "https://www.nta.go.jp/taxes/tetsuzuki/shinsei/annai/gensen/annai/1648_73.htm",
    "https://www.nta.go.jp/taxes/tetsuzuki/shinsei/annai/gensen/annai/1648_74.htm",
    # MLIT - more transport data
    "https://www.mlit.go.jp/statistics/details/tetsudo_list.html",
    "https://www.mlit.go.jp/statistics/details/port_list.html",
    "https://www.mlit.go.jp/statistics/details/kensetu_list.html",
    "https://www.mlit.go.jp/jutakukentiku/house/jutakukentiku_house_tk3_000015.html",
    "https://www.mlit.go.jp/road/road_fr4_000055.html",
    # MOJ - legal forms (docx)
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
    # More docx from MIC
    "https://www.soumu.go.jp/main_sosiki/jichi_zeisei/czaisei/czaisei_seido/ichiran.html",
    "https://www.soumu.go.jp/main_sosiki/jichi_gyousei/bunken/index.html",
    # Digital Agency
    "https://www.digital.go.jp/policies/mynumber/faq-document",
    "https://www.digital.go.jp/policies/posts/mynumber-application",
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
    return doc_links, sub_pages[:25]

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
    for idx, seed in enumerate(PAGES_WITH_DOCS):
        if sum(counts.values()) >= target:
            break
        total = sum(counts.values())
        if idx % 5 == 0:
            print(f"[{idx+1}/{len(PAGES_WITH_DOCS)}] ({total}/{target})")
        to_crawl = [seed]
        crawled = 0
        while to_crawl and crawled < 10 and sum(counts.values()) < target:
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
