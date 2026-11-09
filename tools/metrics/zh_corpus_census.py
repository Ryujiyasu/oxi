"""Census of the zh slice of docx-corpus: Simplified vs Traditional, and the East Asian faces
each document asks for (so documents whose faces this machine lacks are not used as truth).

Script: characters encodable in GB2312 but not Big5 count as Simplified, the reverse as
Traditional; w:lang eastAsia (zh-CN / zh-TW / zh-HK ...) is reported beside it.
Faces: every w:rFonts eastAsia / ascii / hAnsi name in document.xml and styles.xml, plus theme
minor/major East Asian fonts.
Usage: python zh_corpus_census.py [corpus_root]   (default pipeline_data/docx_corpus/zh)
Writes <root>/_census.json and prints a summary.
"""
import sys, os, re, glob, json, zipfile, collections
sys.stdout.reconfigure(encoding="utf-8")

ROOT = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
    os.path.dirname(os.path.abspath(__file__)), '..', '..', 'pipeline_data', 'docx_corpus', 'zh')

# Faces this machine has (system + Office cloud cache), checked 2026-11-09
HAVE = {'SimSun', '宋体', 'NSimSun', '新宋体', 'Microsoft YaHei', '微软雅黑', 'Microsoft JhengHei',
        '微軟正黑體', 'DengXian', '等线', 'PMingLiU', '新細明體', 'MingLiU', '細明體',
        'SimSun-ExtB', 'Microsoft YaHei UI',
        # Windows Chinese supplemental fonts (Language.Fonts.Hans / Hant), installed 2026-11-09
        'SimHei', '黑体', 'KaiTi', '楷体', 'FangSong', '仿宋', 'DFKai-SB', '標楷體',
        'DengXian Light', '等线 Light'}
CJK_FACE = re.compile(r'[㐀-鿿]|SimSun|SimHei|YaHei|JhengHei|DengXian|FangSong|KaiTi|MingLiU|'
                      r'DFKai|STSong|STKaiti|STFangsong|STZhongsong|LiSu|YouYuan|Hei|Song|Kai|Ming', re.I)


def script(text):
    s = t = 0
    for ch in text:
        if not ('一' <= ch <= '鿿'):
            continue
        try:
            ch.encode('gb2312'); g = True
        except UnicodeEncodeError:
            g = False
        try:
            ch.encode('big5'); b = True
        except UnicodeEncodeError:
            b = False
        if g and not b:
            s += 1
        elif b and not g:
            t += 1
    return s, t


rows = []
for path in sorted(glob.glob(os.path.join(ROOT, '*', '*.docx'))):
    try:
        z = zipfile.ZipFile(path)
        doc = z.read('word/document.xml').decode('utf8', 'replace')
        sty = z.read('word/styles.xml').decode('utf8', 'replace') if 'word/styles.xml' in z.namelist() else ''
        thm = ''.join(z.read(n).decode('utf8', 'replace') for n in z.namelist() if n.startswith('word/theme/'))
    except Exception as e:
        rows.append({'doc': os.path.relpath(path, ROOT), 'error': str(e)})
        continue
    text = ''.join(re.findall(r'<w:t(?: [^>]*)?>([^<]*)</w:t>', doc))
    s, t = script(text)
    langs = collections.Counter(re.findall(r'w:eastAsia="(zh-[A-Za-z]+|ja-JP|ko-KR)"', doc + sty))
    faces = collections.Counter()
    for m in re.finditer(r'<w:rFonts ([^>]*)/?>', doc + sty):
        for a in ('eastAsia', 'ascii', 'hAnsi'):
            v = re.search(a + r'="([^"]+)"', m.group(1))
            if v:
                faces[v.group(1)] += 1
    for m in re.finditer(r'<a:font script="Hans" typeface="([^"]+)"|<a:font script="Hant" typeface="([^"]+)"|<a:ea typeface="([^"]+)"', thm):
        f = next(g for g in m.groups() if g)
        if f:
            faces[f] += 1
    cjk_faces = sorted(f for f in faces if CJK_FACE.search(f))
    missing = [f for f in cjk_faces if f not in HAVE]
    rows.append({'doc': os.path.relpath(path, ROOT).replace('\\', '/'), 'hans': s, 'hant': t,
                 'chars': len(text), 'lang': dict(langs), 'cjk_faces': cjk_faces, 'missing': missing,
                 'kind': 'hans' if s > t else ('hant' if t > s else 'none')})

json.dump(rows, open(os.path.join(ROOT, '_census.json'), 'w', encoding='utf8'), ensure_ascii=False, indent=1)
kinds = collections.Counter(r.get('kind') for r in rows)
print('docs', len(rows), dict(kinds))
ok = [r for r in rows if r.get('kind') in ('hans', 'hant') and not r.get('missing')]
print('usable now (no missing face):', collections.Counter(r['kind'] for r in ok))
mf = collections.Counter(f for r in rows for f in r.get('missing', []))
print('missing faces:', mf.most_common(15))
hf = collections.Counter(f for r in rows for f in r.get('cjk_faces', []))
print('all CJK faces:', hf.most_common(15))
