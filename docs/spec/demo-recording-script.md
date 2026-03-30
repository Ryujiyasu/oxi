# Oxi デモ録画スクリプト

## 準備
1. ブラウザで https://ryujiyasu.github.io/oxi/ を開く
2. ブラウザウィンドウを **1280x720** にリサイズ（Twitter推奨サイズ）
3. 日本語モードに切り替え（右上「日本語」ボタン）
4. `Win+G` → Game Bar を起動、録画準備

---

## シーン構成（目標: 35-40秒）

### シーン 1: ランディング（3秒）
- Oxi のトップページを見せる
- ドロップゾーン、サンプルボタンが見える状態

### シーン 2: docx レンダリング（8秒）
1. 「Word (.docx)」サンプルボタンをクリック
2. docx がパースされてレンダリングされる様子を見せる
3. 少しスクロールして中身を見せる

### シーン 3: テキスト編集 & ダウンロード（8秒）
1. テキストの一部をクリックして編集（例: 文字を追加）
2. 「ダウンロード .docx」ボタンをクリック
3. ファイルがダウンロードされる

### シーン 4: 判子 (Hanko)（10秒）  ← メインの見せ場
1. リボンの「Hanko」タブをクリック
2. 名前欄に「山田」→ 判子がリアルタイムプレビュー
3. 名前を「株式会社」に変更 → 角印プレビュー
4. スタイルを「Square」に切り替え → 2x2 グリッド表示

### シーン 5: PDF 生成（6秒）
1. メニュー「File」→「Export」をクリック
2. 「PDF を新規生成」をクリック
3. タイトル「Oxi Demo」、本文「Hello, World!」で生成
4. PDF がダウンロードされる

### シーン 6: エンドカード（5秒）
- ブラウザのアドレスバーを見せる（サーバー通信なし）
- 手動でトップに戻る or 新しいタブで GitHub を開く

---

## 録画後の編集

### テロップ（入れるなら）
- シーン2の開始時: 「.docx をブラウザで即座にレンダリング」
- シーン4の開始時: 「判子をリアルタイム生成」
- シーン6: 「No server. 100% browser. MIT License.」

### ツール
- **録画**: Win+G (Game Bar) or OBS Studio
- **トリミング**: Clipchamp (Windows 標準) or ffmpeg
- **GIF変換**（必要なら）: `ffmpeg -i demo.mp4 -vf "fps=15,scale=640:-1" -loop 0 demo.gif`

### mp4 → Twitter 最適化
```bash
# 720p, 30fps, Twitter向け圧縮
ffmpeg -i raw.mp4 -vf "scale=1280:720" -r 30 -c:v libx264 -crf 23 -preset medium -an oxi-demo.mp4
```

---

## ツイート案（日本語）

```
Rust + WASMで作ったOSSドキュメントスイート「Oxi」を公開しました

.docx / .xlsx / .pptx / PDF をブラウザだけで表示・編集・ダウンロード
サーバー不要、データはブラウザの外に出ません

- 官公庁ファイル 90件 パース成功率 100%
- 判子(印鑑)をSVG生成 → PDF電子署名
- 禁則処理(JIS X 4051)対応
- WASM 1.4MB

Demo: https://ryujiyasu.github.io/oxi/
GitHub: https://github.com/Ryujiyasu/oxi

#Rust #WebAssembly #OSS #OSSドキュメント
```

## ツイート案（英語 — 別ツイート or スレッド2つ目）

```
I built an open-source document suite in Rust + WebAssembly.

Oxi parses, renders, and edits .docx / .xlsx / .pptx / PDF entirely in the browser — no server, no uploads.

- 100% parse rate on 90 Japanese government files
- Digital stamp (Hanko) generator → PDF signatures
- Japanese typography (kinsoku shori)
- ~1.4 MB WASM binary

Demo: https://ryujiyasu.github.io/oxi/
GitHub: https://github.com/Ryujiyasu/oxi

#rustlang #webassembly #opensource
```

## ハッシュタグ戦略

**日本語**: #Rust #WebAssembly #OSS #個人開発 #ドキュメント
**英語**: #rustlang #webassembly #opensource #wasm
**投稿タイミング**: 日本語は平日昼(12-13時) or 夜(20-22時)、英語は日本時間の深夜(=US昼)
