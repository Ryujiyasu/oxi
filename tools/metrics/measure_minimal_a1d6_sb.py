"""minimal_a1d6 variants の各 T5 セル先頭段落の Word 描画 y を COM 計測。

オリジナル a1d6 の T5 では R4/R11/R12/R13/R24/R25/R26 が APPLY (sb 適用).
variant ごとに各 cell first paragraph の y - top_margin と sb 設定を比較し,
Word が APPLY/SUPPRESS どちらかを判定する。

使用方法: 引数で docx パスを指定 (省略時はオリジナル).
"""
import sys, os, json
import win32com.client

NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

def measure(docx_path):
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False
    doc = word.Documents.Open(os.path.abspath(docx_path), ReadOnly=True)
    try:
        # トップ margin (a1d6 は section pgMar top=1247, ≈62.4pt) を pgSetup から取得
        ps = doc.PageSetup
        top_margin = float(ps.TopMargin)
        print(f'doc: {os.path.basename(docx_path)} top_margin={top_margin:.2f}pt')

        n = doc.Paragraphs.Count
        print(f'Total paragraphs: {n}')
        # 各段落の (Section数 = テーブル番号), inTable 状態, y, sb を取得
        rows = []
        for wi in range(1, n + 1):
            p = doc.Paragraphs(wi)
            rng = p.Range
            text = (rng.Text or '').replace('\r','').replace('\x07','').strip()[:40]
            start_rng = doc.Range(rng.Start, rng.Start)
            try:
                page = int(start_rng.Information(3))
                y    = float(start_rng.Information(6))
            except Exception:
                continue
            if y > 800 or y < 0: continue
            in_tbl = rng.Information(12)  # wdWithInTable
            sb = float(p.SpaceBefore)
            ls = float(p.LineSpacing)
            rows.append(dict(wi=wi, page=page, y=y, sb=sb, ls=ls,
                             in_tbl=bool(in_tbl), text=text))
        return rows, top_margin
    finally:
        doc.Close(SaveChanges=False)
        word.Quit()

def main():
    paths = sys.argv[1:] if len(sys.argv) > 1 else [
        'tools/golden-test/documents/docx/a1d6e4efa2e7_tokumei_08_01-4.docx',
        'c:/tmp/minimal_a1d6_v1.docx',
    ]
    for p in paths:
        if not os.path.exists(p):
            print(f'MISSING: {p}')
            continue
        print(f'\n=== {p} ===')
        rows, tm = measure(p)
        # T5 cells のだけ抽出: 段落 in_tbl=True で T5 範囲のもの
        # 単純化: 最初の T5 セル段落から最後の T5 セル段落まで（全 inTable のものを表示）
        # 各行・列の最初の段落 = vMerge restart/continue の先頭
        prev_wi = None
        for r in rows:
            if not r['in_tbl']: continue
            y_off = r['y'] - tm
            # 解釈: sb=N pt なのに y_off ≈ 0 → SUPPRESS, y_off ≈ sb or sb/2 → APPLY
            if r['sb'] < 0.1:
                v = 'sb=0'
            elif abs(y_off) < 1.5:
                v = 'SUPPRESS'
            elif abs(y_off - r['sb']) < 1.5:
                v = 'APPLY-full'
            elif abs(y_off - r['sb']/2) < 1.5:
                v = 'APPLY-half'
            else:
                v = f'partial({y_off:.2f}/{r["sb"]:.2f})'
            # 行・列推定（簡易: テキストプレフィックスで識別）
            tag = ''
            for marker in ['１　匿名', '名称', '２　匿名データの利用目的', '当該', '（１）直接の利用目的',
                            '（２）その他の利用目的', '（３）成果', '３　匿名データの利用場所', '（利用場所',
                            '４　匿名データの利用者の範囲', '氏　名',
                            '５　匿名データの提供を受ける方法', '（１）提供媒体',
                            '（２）提供方法', '（３）提供希望年月日',
                            '６　現に提供を受け',
                            '７　過去の提供履歴', '（１）過去に厚生労働省',
                            '（２）過去に他府省', '８　匿名データの利用場所が日本',
                            '（提供要件）', '９　その他必要な事項']:
                if r['text'].startswith(marker):
                    tag = f' [{marker[:8]}]'
                    break
            if tag and r['sb'] > 0.1:
                print(f"  wi={r['wi']:>4} p{r['page']:<2} y={r['y']:>7.2f} y_off={y_off:>+7.2f} sb={r['sb']:>5.2f} → {v:<14}{tag} {r['text']!r}")

if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    main()
