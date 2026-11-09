"""Freeze a Chinese (Simplified) blind set — rule in pipeline_data/FROZEN_BLIND_ZH_<L>.md,
declared before this runs.

Walks https://api.docxcorp.us/manifest?type=<type>&lang=zh in SHA-256 ascending order,
skips every document already on disk under pipeline_data/docx_corpus/zh/<type>/, and takes
the next 5 per type that pass the mechanical quarantine (zip + word/document.xml, no VBA,
<= 4 MB) AND are Simplified (GB2312-only ideographs > Big5-only ideographs). Top-up across
types in SHA order if a type runs out.

    python tools/metrics/freeze_zh_blind.py A
"""
import io, json, sys, zipfile, re, urllib.request
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')
REPO = Path(__file__).resolve().parents[2]
CORPUS = REPO / 'pipeline_data' / 'docx_corpus' / 'zh'
OUT = REPO / 'pipeline_data' / 'zh_benchmark'
TYPES = ['legal', 'forms', 'reports', 'policies', 'educational',
         'correspondence', 'technical', 'administrative', 'creative', 'reference']
PER_TYPE, N, MAX_MB = 5, 50, 4.0
UA = {'User-Agent': 'oxi-corpus-fetch/1.0 (+https://github.com/Ryujiyasu/oxi)'}


def get(url, t=60):
    return urllib.request.urlopen(urllib.request.Request(url, headers=UA), timeout=t).read()


def script(data):
    z = zipfile.ZipFile(io.BytesIO(data))
    text = ''.join(re.findall(r'<w:t(?: [^>]*)?>([^<]*)</w:t>', z.read('word/document.xml').decode('utf8', 'replace')))
    s = t = 0
    for ch in text:
        if '一' <= ch <= '鿿':
            try:
                ch.encode('gb2312'); g = True
            except UnicodeEncodeError:
                g = False
            try:
                ch.encode('big5'); b = True
            except UnicodeEncodeError:
                b = False
            s += g and not b
            t += b and not g
    return 'hans' if s > t else ('hant' if t > s else 'none')


def check(data):
    if len(data) > MAX_MB * 1024 * 1024:
        return 'size'
    try:
        z = zipfile.ZipFile(io.BytesIO(data))
        names = z.namelist()
    except Exception:
        return 'zip'
    if 'word/document.xml' not in names:
        return 'no document.xml'
    if any(n.endswith('vbaProject.bin') for n in names):
        return 'vba'
    k = script(data)
    return None if k == 'hans' else 'script ' + k


def main(letter):
    log, picked, leftovers = [], [], []
    for t in TYPES:
        have = {p.stem for p in (CORPUS / t).glob('*.docx')}
        manifest = get(f'https://api.docxcorp.us/manifest?type={t}&lang=zh').decode().split()
        got = 0
        for url in manifest:
            sha = url.rsplit('/', 1)[-1].replace('.docx', '')
            if sha[:16] in have:
                continue
            if got >= PER_TYPE:
                leftovers.append((sha, t, url))
                if len([l for l in leftovers if l[1] == t]) > 20:
                    break
                continue
            try:
                data = get(url, 120)
            except Exception as e:
                log.append(f'{t} {sha[:16]} fetch-fail {e}'); continue
            why = check(data)
            if why:
                log.append(f'{t} {sha[:16]} skip {why}'); continue
            (CORPUS / t).mkdir(parents=True, exist_ok=True)
            (CORPUS / t / f'{sha[:16]}.docx').write_bytes(data)
            picked.append(f'{t}/{sha[:16]}'); got += 1
            log.append(f'{t} {sha[:16]} TAKE')
        print(t, got)
    for sha, t, url in sorted(leftovers):
        if len(picked) >= N:
            break
        data = get(url, 120)
        why = check(data)
        if why:
            log.append(f'topup {t} {sha[:16]} skip {why}'); continue
        (CORPUS / t / f'{sha[:16]}.docx').write_bytes(data)
        picked.append(f'{t}/{sha[:16]}'); log.append(f'topup {t} {sha[:16]} TAKE')
    OUT.mkdir(parents=True, exist_ok=True)
    json.dump(picked, open(OUT / f'_final_zhBlind{letter}50.json', 'w', encoding='utf8'), indent=1)
    (OUT / f'zhBlind{letter}_draw.log').write_text('\n'.join(log), encoding='utf8')
    print('picked', len(picked))


if __name__ == '__main__':
    main(sys.argv[1])
