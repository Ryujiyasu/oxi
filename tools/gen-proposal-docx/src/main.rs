use std::io::{Cursor, Write};
use zip::write::{SimpleFileOptions, ZipWriter};

fn main() {
    let output_path = std::env::args()
        .nth(1)
        .unwrap_or_else(|| "proposal.docx".to_string());

    let bytes = build_proposal_docx();
    std::fs::write(&output_path, &bytes).expect("failed to write docx");
    println!("Written {} bytes to {}", bytes.len(), output_path);
}

fn build_proposal_docx() -> Vec<u8> {
    let buf = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(buf);
    let opts = SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated);

    // [Content_Types].xml
    zip.start_file("[Content_Types].xml", opts).unwrap();
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
</Types>"#).unwrap();

    // _rels/.rels
    zip.start_file("_rels/.rels", opts).unwrap();
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#).unwrap();

    // word/_rels/document.xml.rels
    zip.start_file("word/_rels/document.xml.rels", opts).unwrap();
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
</Relationships>"#).unwrap();

    // word/numbering.xml - for bullet lists
    zip.start_file("word/numbering.xml", opts).unwrap();
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="・"/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="420" w:hanging="210"/></w:pPr>
      <w:rPr><w:rFonts w:ascii="Yu Gothic" w:hAnsi="Yu Gothic" w:eastAsia="Yu Gothic"/></w:rPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1">
    <w:abstractNumId w:val="0"/>
  </w:num>
</w:numbering>"#).unwrap();

    // word/styles.xml
    zip.start_file("word/styles.xml", opts).unwrap();
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Yu Gothic" w:hAnsi="Yu Gothic" w:eastAsia="Yu Gothic"/>
        <w:sz w:val="21"/>
        <w:szCs w:val="21"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:spacing w:after="60" w:line="320" w:lineRule="auto"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:styleId="Normal">
    <w:name w:val="Normal"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Title">
    <w:name w:val="Title"/>
    <w:pPr><w:jc w:val="center"/><w:spacing w:before="480" w:after="120"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Subtitle">
    <w:name w:val="Subtitle"/>
    <w:pPr><w:jc w:val="center"/><w:spacing w:after="60"/></w:pPr>
    <w:rPr><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:pPr>
      <w:spacing w:before="360" w:after="120"/>
      <w:pBdr><w:bottom w:val="single" w:sz="8" w:space="1" w:color="2E4057"/></w:pBdr>
    </w:pPr>
    <w:rPr><w:b/><w:sz w:val="28"/><w:szCs w:val="28"/><w:color w:val="2E4057"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
    <w:pPr><w:spacing w:before="240" w:after="80"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="24"/><w:szCs w:val="24"/><w:color w:val="2E4057"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading3">
    <w:name w:val="heading 3"/>
    <w:pPr><w:spacing w:before="200" w:after="60"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>
  </w:style>
</w:styles>"#).unwrap();

    // word/document.xml
    zip.start_file("word/document.xml", opts).unwrap();
    let body = build_document_body();
    zip.write_all(body.as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}

// Helper functions for building OOXML

fn xml_escape(s: &str) -> String {
    s.replace('&', "&amp;")
     .replace('<', "&lt;")
     .replace('>', "&gt;")
}

fn title(text: &str) -> String {
    format!(r#"<w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:r><w:t>{}</w:t></w:r></w:p>"#, xml_escape(text))
}

fn subtitle(text: &str) -> String {
    format!(r#"<w:p><w:pPr><w:pStyle w:val="Subtitle"/></w:pPr><w:r><w:t>{}</w:t></w:r></w:p>"#, xml_escape(text))
}

fn h1(text: &str) -> String {
    format!(r#"<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>{}</w:t></w:r></w:p>"#, xml_escape(text))
}

fn h2(text: &str) -> String {
    format!(r#"<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>{}</w:t></w:r></w:p>"#, xml_escape(text))
}

fn h3(text: &str) -> String {
    format!(r#"<w:p><w:pPr><w:pStyle w:val="Heading3"/></w:pPr><w:r><w:t>{}</w:t></w:r></w:p>"#, xml_escape(text))
}

fn para(text: &str) -> String {
    format!(r#"<w:p><w:r><w:t xml:space="preserve">{}</w:t></w:r></w:p>"#, xml_escape(text))
}

fn bullet(text: &str) -> String {
    format!(r#"<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t xml:space="preserve">{}</w:t></w:r></w:p>"#, xml_escape(text))
}

fn page_break() -> String {
    r#"<w:p><w:r><w:br w:type="page"/></w:r></w:p>"#.to_string()
}

/// Build a table with header row (blue bg) and data rows.
fn table(headers: &[&str], rows: &[Vec<&str>]) -> String {
    let mut xml = String::new();
    xml.push_str(r#"<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/><w:tblBorders>"#);
    for border in &["top", "left", "bottom", "right", "insideH", "insideV"] {
        xml.push_str(&format!(r#"<w:{} w:val="single" w:sz="4" w:space="0" w:color="808080"/>"#, border));
    }
    xml.push_str(r#"</w:tblBorders></w:tblPr>"#);

    // Header row
    xml.push_str("<w:tr>");
    for h in headers {
        xml.push_str(&format!(
            r#"<w:tc><w:tcPr><w:shd w:val="clear" w:fill="2E4057"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:color w:val="FFFFFF"/><w:sz w:val="20"/></w:rPr><w:t>{}</w:t></w:r></w:p></w:tc>"#,
            xml_escape(h)
        ));
    }
    xml.push_str("</w:tr>");

    // Data rows
    for row in rows {
        xml.push_str("<w:tr>");
        for cell in row {
            xml.push_str(&format!(
                r#"<w:tc><w:p><w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t xml:space="preserve">{}</w:t></w:r></w:p></w:tc>"#,
                xml_escape(cell)
            ));
        }
        xml.push_str("</w:tr>");
    }

    xml.push_str("</w:tbl>");
    xml
}

fn build_document_body() -> String {
    let mut body = String::new();

    body.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<w:body>"#);

    // ============================================
    // Page 1: Title + Section 1
    // ============================================
    body.push_str(&title("脱OfficeのためのOSS完全互換スイートOxiの開発"));
    body.push_str(&subtitle("一人で世界標準を作る時代が来た"));
    body.push_str(&subtitle("提案者：安河内 竜二　　2026年3月"));

    body.push_str(&h1("① プロジェクトの背景、目的、目標"));

    body.push_str(&h2("1-1. 背景"));

    body.push_str(&h3("空白の30年と、逆転の機会"));
    body.push_str(&para("空白の30年で、日本は周りの国に抜かれていった。IT産業においてその差は特に顕著であり、ソフトウェア開発の主導権を海外企業に譲り続けた結果、日本の官公庁・企業は文書処理の基盤すらMicrosoft（米国）に依存する構造が固定化された。"));
    body.push_str(&para("財政難の自治体は古いソフトを使い続けるしかない。生産性は上がらず、職員の負担は増え、人手不足が加速する悪循環に陥っている。LibreOfficeへの移行を試みた自治体も多いが、日本語レイアウトが崩れるという根本的な問題により断念した事例が全国に存在する。しかし今、この状況を逆転できる歴史上初めてのタイミングが来ている。"));

    body.push_str(&h3("大企業は会議をしている間に、一人が世界を作る"));
    body.push_str(&para("大企業はリスク管理と組織の意思決定コストという負債を抱えている。新技術の採用には会議・承認・法務確認が必要で、身動きが取れない。日本の大企業はレガシーシステムという技術的負債も抱え、過去の投資を守るために新しいスタックへの移行が遅れる。LibreOfficeなど既存OSSもC++による巨大なコードベースという負債を持ち、日本語組版への対応が後回しにされ続けてきた。"));
    body.push_str(&para("身軽な個人開発者にはこれらの負債がすべてない。意思決定はゼロ秒、方向転換は即時、AIを100%全力で使い倒せる。実際、提案者はこの提案書の作成と並行して開発を開始し、わずか2日で.docx/.xlsx/.pptxの3形式すべてに対応したブラウザ動作デモを完成させた。これは理論ではなく、すでに証明された事実である。"));

    body.push_str(&h3("AIコーディングのウィンドウは今しか開いていない"));
    body.push_str(&para("Claude Codeをはじめとするエージェント型AI開発環境の登場により、かつては大規模チームにしか不可能だった複雑な実装が個人で現実的になった。このウィンドウは永遠に開いていない。大企業がAIコーディングを本格採用した瞬間、リソース量で逆転される。今動いた者だけが標準を作れる。"));

    body.push_str(&h3("世界を日本に合わせる"));
    body.push_str(&para("これまで日本はWordに合わせて文書を作り、英語UIを日本語化し、海外OSSに日本語対応パッチを当ててきた。常に「世界標準に合わせる」側だった。Oxiは逆の発想で作る。判子・縦書き・禁則処理・ルビ・公文書書式を「例外処理」ではなく「コア機能」として設計する。日本の文書文化を一等市民として扱い、世界がそれに合わせる基盤を作る。同じ問題を抱えるCJK文字圏（中国・韓国・台湾）のデファクトスタンダードになり得る。"));

    body.push_str(&h2("1-2. 目的"));
    body.push_str(&para("すでに動作するOSSオフィススイート「Oxi」（.docx/.xlsx/.pptx/.pdf対応・Rust+Wasm実装）をベースに、Microsoft Officeとの完全互換性を確立し、日本の官公庁・自治体・教育機関への実導入を実現する。ブラウザさえあればWindows・Mac・Linux・Chromebook どこでも動作し、インストール不要・ライセンス費用ゼロ。MITライセンスによる完全無償公開で、財政難の自治体が古いソフトを使い続けざるを得ない悪循環を断ち切る。"));

    body.push_str(&h2("1-3. 目標（事業期間：2026年7月〜2027年2月）"));
    body.push_str(&bullet(".docx/.xlsx/.pptx全形式のレイアウト精度向上（Microsoft Officeとのピクセル単位一致を目標とする自動テストスイート構築）"));
    body.push_str(&bullet("日本語組版の完全実装（JIS X 4051準拠：禁則処理・ルビ・縦書き・均等割り付け）"));
    body.push_str(&bullet("国内自治体・教育機関との実証実験（最低1団体での試験導入）"));
    body.push_str(&bullet("GitHubスター500以上・OSSコミュニティの形成"));
    body.push_str(&bullet("CJK文字圏（中国語・韓国語）対応の基盤整備"));

    // ============================================
    // Page 2: Section 1-4 (table)
    // ============================================
    body.push_str(&h2("1-4. 競合との差別化"));
    body.push_str(&table(
        &["比較項目", "Oxi（本提案）", "LibreOffice", "Google Docs", "OnlyOffice"],
        &[
            vec!["実装言語", "Rust + Wasm", "C++ / Java", "JavaScript", "JavaScript"],
            vec!["Word互換精度", "完全一致を目標", "差異あり", "差異あり", "差異あり"],
            vec!["日本語組版", "◎ コア設計", "△ 後付け対応", "△ 部分対応", "△ 部分対応"],
            vec!["ブラウザ・OS非依存", "◎ Wasmのみ", "× Windows依存強", "◎", "△"],
            vec!["判子・電子署名", "◎ PAdES対応", "×", "×", "×"],
            vec!["ライセンス", "MIT（完全無償）", "LGPL", "プロプライエタリ", "一部商用制限"],
        ],
    ));

    body.push_str(&page_break());

    // ============================================
    // Page 3: Section 2
    // ============================================
    body.push_str(&h1("② プロジェクトの内容"));

    body.push_str(&h2("2-1. 現在の実装状況（提案書提出時点）"));
    body.push_str(&para("提案書提出時点で以下の機能がすべて実装・動作済みである。"));
    body.push_str(&bullet(".docx / .xlsx / .pptx パーサー、言語非依存IR、レイアウトエンジン"));
    body.push_str(&bullet("段落・表・画像・ヘッダー/フッター・ページ罫線の描画"));
    body.push_str(&bullet("日本語禁則処理（JIS X 4051）、13フォント・約55KBのフォントメトリクス"));
    body.push_str(&bullet("3形式のラウンドトリップ編集（元のXMLを保持したままパッチ方式で編集・ダウンロード）"));
    body.push_str(&bullet("PDF 1.7 パーサー・テキスト抽出・PDF生成（oxipdf-core）"));
    body.push_str(&bullet("デジタル判子生成 + PAdES PDF電子署名（oxihanko）"));
    body.push_str(&bullet("Wasmビルド + ブラウザWebデモ（https://ryujiyasu.github.io/oxi/）"));
    body.push_str(&bullet("E2Eテストスイート・ゴールデンテスト基盤"));

    body.push_str(&h2("2-2. アーキテクチャ"));
    body.push_str(&bullet("oxi-common：共通OOXMLユーティリティ（ZIP・XML・relationships）"));
    body.push_str(&bullet("oxidocs-core：.docxエンジン（パーサー・IR・レイアウト・フォントメトリクス・エディタ）"));
    body.push_str(&bullet("oxicells-core：.xlsxエンジン（パーサー・IR・エディタ）"));
    body.push_str(&bullet("oxislides-core：.pptxエンジン（パーサー・IR・エディタ）"));
    body.push_str(&bullet("oxipdf-core：PDF 1.7エンジン（パーサー・テキスト抽出・PDF生成）"));
    body.push_str(&bullet("oxihanko：デジタル判子生成 + PAdES電子署名"));
    body.push_str(&bullet("oxi-wasm：WebAssemblyバインディング（wasm-bindgen）"));

    body.push_str(&h2("2-3. 事業期間中に実施する内容"));
    body.push_str(&bullet("【精度向上】Word/Excel/PowerPointとのピクセル単位レイアウト一致を目指す自動テストスイートの構築・公開。OSSコミュニティと協力して差分を継続的に解消する。"));
    body.push_str(&bullet("【日本語組版完全実装】縦書き・ルビ（振り仮名）・均等割り付けの実装。JIS X 4051に基づく完全対応。"));
    body.push_str(&bullet("【自治体実証実験】LibreOfficeへの移行を断念した自治体・教育機関へのアプローチと試験導入。PMネットワーク・Code for Japan等を活用。"));
    body.push_str(&bullet("【CJK展開基盤】中国語・韓国語の組版対応基盤を整備し、アジア全体のOSSコミュニティへのリーチを確立する。"));

    body.push_str(&page_break());

    // ============================================
    // Page 4-5: Section 3
    // ============================================
    body.push_str(&h1("③ プロジェクトの計画"));

    body.push_str(&h2("3-1. スケジュール"));
    body.push_str(&table(
        &["期間", "フェーズ", "主な作業内容"],
        &[
            vec!["7月〜8月", "精度基盤", "自動テストスイート構築、差分計測システム確立、禁則処理の精度向上"],
            vec!["9月〜10月", "日本語組版", "縦書き・ルビ・均等割り付けの完全実装、公文書書式への対応強化"],
            vec!["11月〜12月", "社会実験", "自治体・教育機関への実証実験開始、フィードバックによる改善、コミュニティ形成"],
            vec!["1月〜2月", "拡張・総括", "CJK展開基盤整備、実証実験結果報告、成果報告書作成"],
        ],
    ));

    body.push_str(&h2("3-2. 克服すべき課題と解決策"));
    body.push_str(&bullet("【課題1】Wordのレイアウトアルゴリズムは非公開：既存のWord文書を大量にサンプリングしレンダリング結果を比較する自動テストスイートを構築・公開。OSSコミュニティと協力して差分を継続的に解消する。"));
    body.push_str(&bullet("【課題2】フォントメトリクスの取得：GitHub ActionsのWindows runnerを用いてフォントをレンダリング・計測しテーブル化する独自システムをすでに構築済み（13フォント・約55KB JSON）。"));
    body.push_str(&bullet("【課題3】自治体との接点づくり：採択後のPMネットワーク活用、Code for Japan経由のアプローチ、公開済みWebデモを活用した直接アプローチを並行実施。"));
    body.push_str(&bullet("【課題4】普及の壁：提案書提出時点ですでに動くデモが公開済み。「LibreOfficeで日本語が崩れた」という体験を持つ担当者への具体的な代替として提示できる。"));

    body.push_str(&page_break());

    // ============================================
    // Page 5-6: Section 4
    // ============================================
    body.push_str(&h1("④ 提案者の実力・能力を示す実績"));

    body.push_str(&h2("4-1. 提案者プロフィール"));
    body.push_str(&para("氏名：安河内 竜二（ヤスコウチ リュウジ）"));
    body.push_str(&para("所属：株式会社エムスクエア・ラボ　取締役CTO"));

    body.push_str(&h2("4-2. 学歴・職歴"));
    body.push_str(&table(
        &["年月", "内容"],
        &[
            vec!["2013年3月", "京都大学農学部 卒業"],
            vec!["2015年3月", "京都大学大学院理学研究科 生物科学専攻 修了"],
            vec!["2016年4月", "月桂冠株式会社 入社（システム開発・IT業務）"],
            vec!["2020年〜", "株式会社エムスクエア・ラボ 入社、取締役CTO就任（現任）"],
        ],
    ));

    body.push_str(&h2("4-3. 技術スキル・実績"));
    body.push_str(&bullet("Rust + Wasm：提案書提出から2日で.docx/.xlsx/.pptx/.pdf対応・判子・PAdES署名・E2Eテスト基盤まで完成（https://ryujiyasu.github.io/oxi/）"));
    body.push_str(&bullet("OOXMLフォーマット解析：Word/Excel/PowerPointの内部仕様を独自に解析・実装。PDF 1.7パーサー・PAdES電子署名も実装済み"));
    body.push_str(&bullet("農地ナビ＋：農地情報・衛星データを統合した農業向けナビゲーションシステムの開発・運用"));
    body.push_str(&bullet("Cloudflare Workers/D1/KV：複数プロダクトの本番運用実績"));
    body.push_str(&bullet("React/TypeScript：フロントエンド開発の実務経験"));
    body.push_str(&bullet("AI活用開発：Claude APIを用いたプロダクト開発（SaaS・チャットボット等）の実績"));

    body.push_str(&h2("4-4. プロジェクト実現性の根拠"));
    body.push_str(&para("提案書提出から2日で全形式対応・判子機能・PAdES電子署名・E2Eテスト基盤まで含むブラウザ動作デモを完成させた。これは理論ではなく実績である。大企業が仕様検討をしている間に、個人がAIコーディングを全力投球することでここまで実装できることの証明でもある。"));
    body.push_str(&para("GitHubリポジトリ：https://github.com/Ryujiyasu/oxi（MITライセンス・公開済み）"));
    body.push_str(&para("Webデモ：https://ryujiyasu.github.io/oxi/"));

    body.push_str(&page_break());

    // ============================================
    // Section 5
    // ============================================
    body.push_str(&h1("⑤ 事業化・社会実装の新規性・優位性、想定するターゲットと規模"));

    body.push_str(&h2("5-1. 市場規模"));
    body.push_str(&para("国内のMicrosoft Officeライセンス市場は年間推計3,000〜5,000億円規模と言われる。官公庁・自治体（約1,800団体）および大学・教育機関（約1,200校）だけでも、年間数百億円のライセンス費用が発生している。OxiはMITライセンスで完全無償提供するため、財政難の自治体が古いソフトを使い続けざるを得ない悪循環を解決できる唯一のOSS基盤となり得る。Wasm＋ブラウザネイティブ設計によりWindowsへの依存も不要となるため、OS調達コストの削減も同時に実現できる。"));

    body.push_str(&h2("5-2. ターゲットユーザー"));
    body.push_str(&bullet("一次ターゲット：脱Office・脱Windows推進中の地方自治体・官公庁"));
    body.push_str(&bullet("二次ターゲット：教育機関（大学・高専・高校）"));
    body.push_str(&bullet("三次ターゲット：文書管理SaaSを開発するスタートアップ・IT企業（APIとして利用）"));
    body.push_str(&bullet("グローバル：CJK文字圏（中国・韓国・台湾）のOSSコミュニティ"));

    body.push_str(&h2("5-3. 競合サービスとの比較"));
    body.push_str(&para("LibreOffice：歴史あるOSSだがC++実装・デスクトップ前提・日本語レイアウト精度に根本的な問題がある。実際に多くの自治体がLibreOfficeへの移行を試みて断念した。Oxiはブラウザネイティブ・インストール不要・日本語を一等市民として設計している点で全方位的に差別化される。"));
    body.push_str(&para("Google Docs / OnlyOffice：独自形式への変換が主体で.docxの完全互換を目指していない。また前者はプロプライエタリ、後者は地政学的リスクを抱える。"));

    body.push_str(&page_break());

    // ============================================
    // Section 6
    // ============================================
    body.push_str(&h1("⑥ 事業化・社会実装の具体的な進め方"));

    body.push_str(&h2("6-1. ビジネスモデル"));
    body.push_str(&para("コアエンジンはMITライセンスで永続的に無償公開する。事業期間終了後は以下のモデルを組み合わせる。"));
    body.push_str(&bullet("【モデル1】OSSコア＋エンタープライズサポート（RedHatモデル）：基本機能を無償提供し、サポート契約・カスタマイズ開発を有償化。自治体1団体年間数百万円 × 全国1,800団体がポテンシャル。"));
    body.push_str(&bullet("【モデル2】クラウドSaaS：Oxiエンジンを用いたOffice互換文書エディタをSaaSとして提供。月額課金モデル。"));
    body.push_str(&bullet("【モデル3】APIライセンス：.docx変換・レンダリング・PDF生成APIを他社SaaSへBtoBで提供。"));

    body.push_str(&h2("6-2. 事業期間中の事業化に向けた作業"));
    body.push_str(&bullet("GitHubリポジトリのスター獲得・コミュニティ形成（目標：500スター）"));
    body.push_str(&bullet("公開済みWebデモを活用した自治体・教育機関へのアプローチ"));
    body.push_str(&bullet("採択後のPMネットワークを活用した自治体コネクション獲得"));
    body.push_str(&bullet("NLnet Foundation等の欧州OSSファンディングへの申請準備"));

    body.push_str(&page_break());

    // ============================================
    // Section 7
    // ============================================
    body.push_str(&h1("⑦ 事業期間終了後の事業化・社会実装に関する計画"));

    body.push_str(&h2("7-1. アウトプット形態"));
    body.push_str(&para("すでにGitHubでMITライセンスにて公開済み（https://github.com/Ryujiyasu/oxi）。事業期間を通じてコミュニティからのバグ報告・PRを取り込みながら精度を高める。完全互換はコミュニティと協力しながら育てていく。"));

    body.push_str(&h2("7-2. ロードマップ（事業期間終了後）"));
    body.push_str(&table(
        &["フェーズ", "時期", "内容"],
        &[
            vec!["v2", "〜2027年末", "CRDTによるリアルタイム共同編集、AIアシスト機能、エンドツーエンド暗号化、PWA/オフライン対応"],
            vec!["v3", "〜2028年末", "SaaS正式開始、CJK市場（中国・台湾）展開、プラグインシステム、Tauriデスクトップ・モバイルアプリ"],
            vec!["v4", "〜2029年末", "エンタープライズ向けコンプライアンス・監査証跡、業界特化（法務・医療・行政）、LibreOfficeに代わる次世代OSSスイートとして世界標準化"],
        ],
    ));

    body.push_str(&page_break());

    // ============================================
    // Section 8
    // ============================================
    body.push_str(&h1("⑧ 予算内訳のまとめ"));

    body.push_str(&h2("8-1. 稼働計画"));
    body.push_str(&table(
        &["メンバー", "月平均稼働時間", "事業期間（月数）", "総稼働時間", "金額（税抜）"],
        &[
            vec!["安河内 竜二（課税事業者）", "80時間/月", "8ヶ月", "640時間", "3,200,000円"],
            vec!["合計", "", "", "640時間", "3,200,000円"],
        ],
    ));

    body.push_str(&h2("8-2. 費用総額"));
    body.push_str(&table(
        &["項目", "金額"],
        &[
            vec!["作業費（税抜）：640時間 × 5,000円/時間", "3,200,000円"],
            vec!["消費税（10%）", "320,000円"],
            vec!["実施費用総額（税込・申請金額）", "3,520,000円"],
        ],
    ));

    body.push_str(&para("※時間単価：5,000円/時間（公募要領4.(4)①に基づく一律単価）"));
    body.push_str(&para("※1名プロジェクトのため上限800万円の範囲内"));

    body.push_str(&page_break());

    // ============================================
    // Section 9
    // ============================================
    body.push_str(&h1("⑨ 提案プロジェクトと所属企業のビジネスモデル・技術開発との差異"));

    body.push_str(&h2("9-1. 所属組織について"));
    body.push_str(&para("提案者は株式会社エムスクエア・ラボの取締役CTOとして在籍している。同社は農業ロボット・スマート農業システムの研究開発・販売を主事業としており、農業IoTセンサー・農機自動化システム・農業データ分析プラットフォームの開発に特化している。"));

    body.push_str(&h2("9-2. 本提案プロジェクトとの差異"));
    body.push_str(&para("本提案プロジェクトは文書処理・Officeフォーマット互換性というテーマであり、農業ロボット・スマート農業という所属組織の事業内容と技術的・ビジネス的に全く異なる領域である。所属組織の技術開発はROS・センサーフュージョン・農業機械制御を中心としており、RustやWebAssemblyを用いたブラウザ上の文書処理エンジンの開発とは接点がない。"));
    body.push_str(&para("本プロジェクトは提案者が個人として取り組む新規OSS開発であり、所属組織のビジネスモデルや技術開発と類似・重複する内容は一切含まない。"));

    body.push_str(&h2("9-3. 所属組織からの了解"));
    body.push_str(&para("所属組織（株式会社エムスクエア・ラボ）からは、本事業による支援措置を受けること、および開発成果がイノベータ個人に帰属することについて了解を得ている。契約時には書面による承諾書を提出する。"));

    // Footer note
    body.push_str(&para(""));
    body.push_str(&para("本提案書のPDFはOxi（oxipdf-core）で生成しました。"));

    // Section properties (A4, margins)
    body.push_str(r#"<w:sectPr>
  <w:pgSz w:w="11906" w:h="16838"/>
  <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
</w:sectPr>"#);

    body.push_str("</w:body></w:document>");
    body
}
