// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! `OXI_IGNORE_LRPB` switches off Word's saved page-break hint
//! (`<w:lastRenderedPageBreak/>`) at the parser, so a benchmark scored against
//! Word cannot read Word's answer from the file.
//!
//! One test in this file on purpose: it sets a process-wide environment
//! variable, and integration test files run as separate processes.

use std::io::{Cursor, Write};

use oxidocs_core::ir::Block;
use oxidocs_core::parse_docx;

fn docx_with_lrpb() -> Vec<u8> {
    let mut buf = Cursor::new(Vec::new());
    {
        let mut z = zip::ZipWriter::new(&mut buf);
        let opt = zip::write::SimpleFileOptions::default();
        let files = [
            ("[Content_Types].xml", r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#),
            ("_rels/.rels", r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#),
            ("word/document.xml", r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>first</w:t></w:r></w:p><w:p><w:r><w:lastRenderedPageBreak/><w:t>second</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>"#),
        ];
        for (name, body) in files {
            z.start_file(name, opt).unwrap();
            z.write_all(body.as_bytes()).unwrap();
        }
        z.finish().unwrap();
    }
    buf.into_inner()
}

fn marker_count(data: &[u8]) -> (usize, usize) {
    let doc = parse_docx(data).expect("parse");
    let runs = doc.pages.iter().flat_map(|p| p.blocks.iter()).filter_map(|b| match b {
        Block::Paragraph(p) => Some(p.runs.iter().filter(|r| r.has_last_rendered_page_break).count()),
        _ => None,
    }).sum();
    let total = doc.pages.iter().map(|p| p.total_lrpb_count).sum();
    (runs, total)
}

#[test]
fn ignore_lrpb_env_drops_the_marker_and_its_count() {
    let data = docx_with_lrpb();

    std::env::remove_var("OXI_IGNORE_LRPB");
    assert_eq!(marker_count(&data), (1, 1), "default: the marker is read");

    std::env::set_var("OXI_IGNORE_LRPB", "0");
    assert_eq!(marker_count(&data), (1, 1), "\"0\" leaves it on");

    std::env::set_var("OXI_IGNORE_LRPB", "1");
    assert_eq!(marker_count(&data), (0, 0), "ignored: no run flag, no section count");

    std::env::remove_var("OXI_IGNORE_LRPB");
}
