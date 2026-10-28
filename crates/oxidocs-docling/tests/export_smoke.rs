// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! A two-paragraph .docx built in memory: the export names the schema, has one
//! page, and every text item carries a provenance on that page whose bbox lies
//! inside it and whose charspan covers the item's text.

use std::io::{Cursor, Write};

fn docx(paragraphs: &[&str]) -> Vec<u8> {
    let body: String = paragraphs
        .iter()
        .map(|t| format!("<w:p><w:r><w:t>{t}</w:t></w:r></w:p>"))
        .collect();
    let files = [
        ("[Content_Types].xml", r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_string()),
        ("_rels/.rels", r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_string()),
        ("word/document.xml", format!(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>"#)),
    ];
    let mut out = Cursor::new(Vec::new());
    {
        let mut z = zip::ZipWriter::new(&mut out);
        let opts = zip::write::SimpleFileOptions::default();
        for (name, data) in files {
            z.start_file(name, opts).unwrap();
            z.write_all(data.as_bytes()).unwrap();
        }
        z.finish().unwrap();
    }
    out.into_inner()
}

#[test]
fn every_text_item_carries_provenance_on_its_page() {
    let data = docx(&["Hello provenance.", "Second paragraph."]);
    let (doc, stats) = oxidocs_docling::docx_to_docling(&data, "smoke.docx").expect("export");
    assert_eq!(doc["schema_name"], "DoclingDocument");
    assert_eq!(doc["version"], "1.10.0");
    assert_eq!(stats.pages, 1);
    let page = &doc["pages"]["1"]["size"];
    let (pw, ph) = (page["width"].as_f64().unwrap(), page["height"].as_f64().unwrap());
    assert!((pw - 612.0).abs() < 0.5 && (ph - 792.0).abs() < 0.5, "Letter page in points: {pw} x {ph}");
    let texts = doc["texts"].as_array().unwrap();
    let wanted = ["Hello provenance.", "Second paragraph."];
    for want in wanted {
        let item = texts.iter().find(|t| t["text"] == want).unwrap_or_else(|| panic!("no item {want:?}"));
        let prov = item["prov"].as_array().unwrap();
        assert_eq!(prov.len(), 1, "{want:?}");
        let p = &prov[0];
        assert_eq!(p["page_no"], 1);
        let b = &p["bbox"];
        let (l, t, r, bt) = (b["l"].as_f64().unwrap(), b["t"].as_f64().unwrap(), b["r"].as_f64().unwrap(), b["b"].as_f64().unwrap());
        assert_eq!(b["coord_origin"], "TOPLEFT");
        assert!(0.0 <= l && l < r && r <= pw && 0.0 <= t && t < bt && bt <= ph, "{want:?} bbox {b}");
        assert_eq!(p["charspan"], serde_json::json!([0, want.chars().count()]), "{want:?}");
    }
}
