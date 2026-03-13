//! PDF file generation.
//!
//! Builds a valid PDF file from IR types. Produces PDF 1.7 output
//! with FlateDecode-compressed content streams.

use crate::ir::*;
use flate2::write::ZlibEncoder;
use flate2::Compression;
use std::io::Write;

/// Build a PDF file from a `PdfDocument` and return the bytes.
pub fn write_pdf(doc: &PdfDocument) -> Vec<u8> {
    let mut writer = PdfWriter::new();
    writer.build(doc);
    writer.finish()
}

struct PdfWriter {
    buf: Vec<u8>,
    /// Byte offsets of each object (indexed by obj number, 0 is unused).
    offsets: Vec<u64>,
    /// Next object number to assign.
    next_obj: u32,
}

impl PdfWriter {
    fn new() -> Self {
        Self {
            buf: Vec::new(),
            offsets: vec![0], // obj 0 is reserved
            next_obj: 1,
        }
    }

    fn alloc_obj(&mut self) -> u32 {
        let num = self.next_obj;
        self.next_obj += 1;
        num
    }

    fn begin_obj(&mut self, num: u32) {
        while self.offsets.len() <= num as usize {
            self.offsets.push(0);
        }
        self.offsets[num as usize] = self.buf.len() as u64;
        write!(self.buf, "{num} 0 obj\n").unwrap();
    }

    fn end_obj(&mut self) {
        write!(self.buf, "endobj\n").unwrap();
    }

    fn build(&mut self, doc: &PdfDocument) {
        // Header
        write!(
            self.buf,
            "%PDF-{}.{}\n",
            doc.version.major, doc.version.minor
        )
        .unwrap();
        // Binary comment to indicate this PDF contains binary data.
        self.buf.extend_from_slice(&[b'%', 0xE2, 0xE3, 0xCF, 0xD3, b'\n']);

        // Pre-allocate object numbers.
        let catalog_num = self.alloc_obj(); // 1
        let pages_num = self.alloc_obj(); // 2
        let info_num = self.alloc_obj(); // 3

        // Collect all unique font names across all pages.
        let mut all_font_names: Vec<String> = Vec::new();
        for page in &doc.pages {
            for el in &page.contents {
                if let ContentElement::Text(span) = el {
                    let name = escape_name(&span.font_name);
                    if !all_font_names.contains(&name) {
                        all_font_names.push(name);
                    }
                }
            }
        }

        // Allocate font objects.
        let font_objs: Vec<(String, u32)> = all_font_names
            .iter()
            .map(|name| (name.clone(), self.alloc_obj()))
            .collect();

        // Collect images per page and allocate XObject references.
        let mut page_images: Vec<Vec<(usize, u32)>> = Vec::new(); // (image_index_in_contents, obj_num)
        for page in &doc.pages {
            let mut images = Vec::new();
            for (idx, el) in page.contents.iter().enumerate() {
                if let ContentElement::Image(_) = el {
                    let obj_num = self.alloc_obj();
                    images.push((idx, obj_num));
                }
            }
            page_images.push(images);
        }

        // Allocate per-page objects: each page needs a page obj + content stream obj.
        let mut page_objs = Vec::new();
        for _ in &doc.pages {
            let page_num = self.alloc_obj();
            let stream_num = self.alloc_obj();
            page_objs.push((page_num, stream_num));
        }

        // Write catalog.
        self.begin_obj(catalog_num);
        write!(
            self.buf,
            "<< /Type /Catalog /Pages {pages_num} 0 R >>\n"
        )
        .unwrap();
        self.end_obj();

        // Write pages tree.
        self.begin_obj(pages_num);
        write!(self.buf, "<< /Type /Pages /Kids [").unwrap();
        for (page_num, _) in &page_objs {
            write!(self.buf, "{page_num} 0 R ").unwrap();
        }
        write!(self.buf, "] /Count {} >>\n", doc.pages.len()).unwrap();
        self.end_obj();

        // Write info dictionary.
        self.begin_obj(info_num);
        write!(self.buf, "<<").unwrap();
        if let Some(ref title) = doc.info.title {
            write!(self.buf, " /Title ({})", escape_pdf_string(title)).unwrap();
        }
        if let Some(ref author) = doc.info.author {
            write!(self.buf, " /Author ({})", escape_pdf_string(author)).unwrap();
        }
        if let Some(ref subject) = doc.info.subject {
            write!(self.buf, " /Subject ({})", escape_pdf_string(subject)).unwrap();
        }
        write!(
            self.buf,
            " /Producer (oxipdf-core {}) >>\n",
            env!("CARGO_PKG_VERSION")
        )
        .unwrap();
        self.end_obj();

        // Determine which fonts need CIDFont (have non-ASCII text).
        let mut cid_font_chars: std::collections::HashMap<String, std::collections::BTreeSet<u16>> =
            std::collections::HashMap::new();
        for page in &doc.pages {
            for el in &page.contents {
                if let ContentElement::Text(span) = el {
                    let name = escape_name(&span.font_name);
                    if span.text.chars().any(|c| c as u32 > 0x7F) {
                        let entry = cid_font_chars.entry(name).or_default();
                        for ch in span.text.chars() {
                            entry.insert(ch as u16);
                        }
                    }
                }
            }
        }

        // Write font objects.
        for (name, obj_num) in &font_objs {
            if let Some(used_chars) = cid_font_chars.get(name) {
                // Type0 composite font for CJK/Unicode text.
                let cid_font_num = self.alloc_obj();
                let tounicode_num = self.alloc_obj();
                let descriptor_num = self.alloc_obj();

                // Type0 font dictionary
                self.begin_obj(*obj_num);
                write!(
                    self.buf,
                    "<< /Type /Font /Subtype /Type0 /BaseFont /{name} \
                     /Encoding /Identity-H \
                     /DescendantFonts [{cid_font_num} 0 R] \
                     /ToUnicode {tounicode_num} 0 R >>\n"
                )
                .unwrap();
                self.end_obj();

                // CIDFont dictionary
                self.begin_obj(cid_font_num);
                write!(
                    self.buf,
                    "<< /Type /Font /Subtype /CIDFontType2 /BaseFont /{name} \
                     /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> \
                     /FontDescriptor {descriptor_num} 0 R \
                     /DW 1000 >>\n"
                )
                .unwrap();
                self.end_obj();

                // Font descriptor (minimal, enough for PDF readers to look up system font)
                self.begin_obj(descriptor_num);
                write!(
                    self.buf,
                    "<< /Type /FontDescriptor /FontName /{name} \
                     /Flags 4 /ItalicAngle 0 \
                     /Ascent 880 /Descent -120 /CapHeight 740 \
                     /StemV 80 \
                     /FontBBox [-100 -250 1100 900] >>\n"
                )
                .unwrap();
                self.end_obj();

                // ToUnicode CMap stream
                let cmap_data = build_tounicode_cmap(used_chars);
                let cmap_compressed = compress(&cmap_data);
                self.begin_obj(tounicode_num);
                write!(
                    self.buf,
                    "<< /Length {} /Filter /FlateDecode >>\nstream\n",
                    cmap_compressed.len()
                )
                .unwrap();
                self.buf.extend_from_slice(&cmap_compressed);
                write!(self.buf, "\nendstream\n").unwrap();
                self.end_obj();
            } else {
                // Simple Type1 base font for ASCII-only text.
                self.begin_obj(*obj_num);
                write!(
                    self.buf,
                    "<< /Type /Font /Subtype /Type1 /BaseFont /{name} >>\n"
                )
                .unwrap();
                self.end_obj();
            }
        }

        // Write image XObjects.
        for (page_idx, images) in page_images.iter().enumerate() {
            for &(content_idx, obj_num) in images {
                if let ContentElement::Image(img) = &doc.pages[page_idx].contents[content_idx] {
                    let cs_name = match img.color_space {
                        ColorSpace::DeviceGray => "DeviceGray",
                        ColorSpace::DeviceRgb => "DeviceRGB",
                        ColorSpace::DeviceCmyk => "DeviceCMYK",
                    };
                    let compressed = compress(&img.data);
                    self.begin_obj(obj_num);
                    write!(
                        self.buf,
                        "<< /Type /XObject /Subtype /Image \
                         /Width {} /Height {} \
                         /ColorSpace /{cs_name} \
                         /BitsPerComponent {} \
                         /Length {} /Filter /FlateDecode >>\nstream\n",
                        img.width as u32,
                        img.height as u32,
                        img.bits_per_component,
                        compressed.len()
                    )
                    .unwrap();
                    self.buf.extend_from_slice(&compressed);
                    write!(self.buf, "\nendstream\n").unwrap();
                    self.end_obj();
                }
            }
        }

        // Write each page + content stream.
        for (i, page) in doc.pages.iter().enumerate() {
            let (page_num, stream_num) = page_objs[i];

            // Build content stream data (with image references).
            let content_data = build_content_stream_with_images(page, &page_images[i], &cid_font_chars);
            let compressed = compress(&content_data);

            // Write content stream object.
            self.begin_obj(stream_num);
            write!(
                self.buf,
                "<< /Length {} /Filter /FlateDecode >>\nstream\n",
                compressed.len()
            )
            .unwrap();
            self.buf.extend_from_slice(&compressed);
            write!(self.buf, "\nendstream\n").unwrap();
            self.end_obj();

            // Write page object.
            self.begin_obj(page_num);
            write!(
                self.buf,
                "<< /Type /Page /Parent {pages_num} 0 R \
                 /MediaBox [{} {} {} {}] \
                 /Contents {stream_num} 0 R",
                page.media_box.llx, page.media_box.lly, page.media_box.urx, page.media_box.ury,
            )
            .unwrap();

            // Resources dictionary.
            let has_fonts = !font_objs.is_empty();
            let has_images = !page_images[i].is_empty();
            if has_fonts || has_images {
                write!(self.buf, " /Resources <<").unwrap();
                if has_fonts {
                    write!(self.buf, " /Font <<").unwrap();
                    for (name, obj_num) in &font_objs {
                        write!(self.buf, " /{name} {obj_num} 0 R").unwrap();
                    }
                    write!(self.buf, " >>").unwrap();
                }
                if has_images {
                    write!(self.buf, " /XObject <<").unwrap();
                    for (img_idx, &(_, obj_num)) in page_images[i].iter().enumerate() {
                        write!(self.buf, " /Im{img_idx} {obj_num} 0 R").unwrap();
                    }
                    write!(self.buf, " >>").unwrap();
                }
                write!(self.buf, " >>").unwrap();
            }
            if page.rotation != 0 {
                write!(self.buf, " /Rotate {}", page.rotation).unwrap();
            }
            write!(self.buf, " >>\n").unwrap();
            self.end_obj();
        }
    }

    fn finish(mut self) -> Vec<u8> {
        // Xref table.
        let xref_offset = self.buf.len();
        write!(self.buf, "xref\n").unwrap();
        write!(self.buf, "0 {}\n", self.offsets.len()).unwrap();
        // Object 0: free entry.
        // Each xref entry must be exactly 20 bytes: "oooooooooo ggggg X\r\n"
        self.buf.extend_from_slice(b"0000000000 65535 f\r\n");
        for offset in &self.offsets[1..] {
            let entry = format!("{:010} 00000 n\r\n", offset);
            self.buf.extend_from_slice(entry.as_bytes());
        }

        // Trailer.
        write!(
            self.buf,
            "trailer\n<< /Size {} /Root 1 0 R /Info 3 0 R >>\n",
            self.offsets.len()
        )
        .unwrap();
        write!(self.buf, "startxref\n{xref_offset}\n%%EOF\n").unwrap();

        self.buf
    }
}

/// Build a content stream from a page's contents, including image references.
fn build_content_stream_with_images(
    page: &Page,
    images: &[(usize, u32)],
    cid_fonts: &std::collections::HashMap<String, std::collections::BTreeSet<u16>>,
) -> Vec<u8> {
    let mut buf = Vec::new();
    let mut img_counter = 0usize;

    for (idx, element) in page.contents.iter().enumerate() {
        match element {
            ContentElement::Text(span) => {
                let font_key = escape_name(&span.font_name);
                write_color_op(&mut buf, &span.fill_color, false);
                write!(buf, "BT\n").unwrap();
                write!(buf, "/{} {} Tf\n", font_key, span.font_size).unwrap();
                let pdf_y = page.height - span.y;
                write!(buf, "{} {} Td\n", span.x, pdf_y).unwrap();
                if cid_fonts.contains_key(&font_key) {
                    // CIDFont: encode as UTF-16BE hex string
                    write!(buf, "<{}> Tj\n", encode_utf16be_hex(&span.text)).unwrap();
                } else {
                    write!(buf, "({}) Tj\n", escape_pdf_string(&span.text)).unwrap();
                }
                write!(buf, "ET\n").unwrap();
            }
            ContentElement::Path(path) => {
                if let Some(ref fill) = path.fill {
                    write_color_op(&mut buf, fill, false);
                }
                if let Some(ref stroke) = path.stroke {
                    write_color_op(&mut buf, &stroke.color, true);
                    write!(buf, "{} w\n", stroke.width).unwrap();
                }
                for op in &path.operations {
                    match op {
                        PathOp::MoveTo(x, y) => {
                            let py = page.height - y;
                            write!(buf, "{x} {py} m\n").unwrap();
                        }
                        PathOp::LineTo(x, y) => {
                            let py = page.height - y;
                            write!(buf, "{x} {py} l\n").unwrap();
                        }
                        PathOp::CurveTo(x1, y1, x2, y2, x3, y3) => {
                            write!(
                                buf,
                                "{x1} {} {x2} {} {x3} {} c\n",
                                page.height - y1,
                                page.height - y2,
                                page.height - y3,
                            )
                            .unwrap();
                        }
                        PathOp::ClosePath => write!(buf, "h\n").unwrap(),
                    }
                }
                match (&path.fill, &path.stroke) {
                    (Some(_), Some(_)) => write!(buf, "B\n").unwrap(),
                    (Some(_), None) => write!(buf, "f\n").unwrap(),
                    (None, Some(_)) => write!(buf, "S\n").unwrap(),
                    (None, None) => write!(buf, "n\n").unwrap(),
                }
            }
            ContentElement::Image(img) => {
                // Check if this image has an allocated XObject.
                if images.iter().any(|&(ci, _)| ci == idx) {
                    // Place image using cm (transformation matrix) + Do.
                    let pdf_y = page.height - img.y - img.height;
                    write!(buf, "q\n").unwrap();
                    write!(
                        buf,
                        "{} 0 0 {} {} {} cm\n",
                        img.width, img.height, img.x, pdf_y
                    )
                    .unwrap();
                    write!(buf, "/Im{img_counter} Do\n").unwrap();
                    write!(buf, "Q\n").unwrap();
                    img_counter += 1;
                }
            }
        }
    }

    buf
}

fn write_color_op(buf: &mut Vec<u8>, color: &Color, stroke: bool) {
    match color {
        Color::Gray(g) => {
            let op = if stroke { "G" } else { "g" };
            write!(buf, "{g} {op}\n").unwrap();
        }
        Color::Rgb(r, g, b) => {
            let op = if stroke { "RG" } else { "rg" };
            write!(buf, "{r} {g} {b} {op}\n").unwrap();
        }
        Color::Cmyk(c, m, y, k) => {
            let op = if stroke { "K" } else { "k" };
            write!(buf, "{c} {m} {y} {k} {op}\n").unwrap();
        }
    }
}

fn escape_pdf_string(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    for c in s.chars() {
        match c {
            '(' => out.push_str("\\("),
            ')' => out.push_str("\\)"),
            '\\' => out.push_str("\\\\"),
            _ => out.push(c),
        }
    }
    out
}

fn escape_name(s: &str) -> String {
    // Simple name escaping — just strip the leading '/' if present.
    if let Some(stripped) = s.strip_prefix('/') {
        stripped.to_string()
    } else {
        s.to_string()
    }
}

/// Build a ToUnicode CMap that maps CIDs (= UTF-16 code units) back to Unicode.
/// This allows PDF readers to extract text when copying/searching.
fn build_tounicode_cmap(used_chars: &std::collections::BTreeSet<u16>) -> Vec<u8> {
    let mut buf = Vec::new();
    write!(buf, "/CIDInit /ProcSet findresource begin\n").unwrap();
    write!(buf, "12 dict begin\n").unwrap();
    write!(buf, "begincmap\n").unwrap();
    write!(buf, "/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n")
        .unwrap();
    write!(buf, "/CMapName /Adobe-Identity-UCS def\n").unwrap();
    write!(buf, "/CMapType 2 def\n").unwrap();
    write!(buf, "1 begincodespacerange\n").unwrap();
    write!(buf, "<0000> <FFFF>\n").unwrap();
    write!(buf, "endcodespacerange\n").unwrap();

    // Write bfchar entries in chunks of 100 (PDF spec limit).
    let chars: Vec<u16> = used_chars.iter().copied().collect();
    for chunk in chars.chunks(100) {
        write!(buf, "{} beginbfchar\n", chunk.len()).unwrap();
        for &code in chunk {
            write!(buf, "<{:04X}> <{:04X}>\n", code, code).unwrap();
        }
        write!(buf, "endbfchar\n").unwrap();
    }

    write!(buf, "endcmap\n").unwrap();
    write!(buf, "CMapName currentdict /CMap defineresource pop\n").unwrap();
    write!(buf, "end\nend\n").unwrap();
    buf
}

/// Encode a string as UTF-16BE hex for use with CIDFont (Identity-H encoding).
fn encode_utf16be_hex(s: &str) -> String {
    let mut hex = String::new();
    for code_unit in s.encode_utf16() {
        use std::fmt::Write;
        write!(hex, "{:04X}", code_unit).unwrap();
    }
    hex
}

fn compress(data: &[u8]) -> Vec<u8> {
    let mut encoder = ZlibEncoder::new(Vec::new(), Compression::default());
    encoder.write_all(data).unwrap();
    encoder.finish().unwrap()
}

#[cfg(test)]
mod tests {
    use super::*;

    fn make_simple_doc() -> PdfDocument {
        PdfDocument {
            version: PdfVersion::new(1, 7),
            info: DocumentInfo {
                title: Some("Test Document".into()),
                author: Some("Oxi".into()),
                ..Default::default()
            },
            pages: vec![Page {
                width: 612.0,
                height: 792.0,
                media_box: Rectangle {
                    llx: 0.0,
                    lly: 0.0,
                    urx: 612.0,
                    ury: 792.0,
                },
                crop_box: None,
                contents: vec![ContentElement::Text(TextSpan {
                    x: 72.0,
                    y: 72.0,
                    text: "Hello, PDF!".into(),
                    font_name: "F1".into(),
                    font_size: 12.0,
                    fill_color: Color::Gray(0.0),
                })],
                rotation: 0,
            }],
            outline: Vec::new(),
        }
    }

    #[test]
    fn test_write_pdf_header() {
        let doc = make_simple_doc();
        let bytes = write_pdf(&doc);
        let s = String::from_utf8_lossy(&bytes);
        assert!(s.starts_with("%PDF-1.7"));
    }

    #[test]
    fn test_write_pdf_has_eof() {
        let doc = make_simple_doc();
        let bytes = write_pdf(&doc);
        let s = String::from_utf8_lossy(&bytes);
        assert!(s.contains("%%EOF"));
    }

    #[test]
    fn test_write_pdf_has_xref() {
        let doc = make_simple_doc();
        let bytes = write_pdf(&doc);
        let s = String::from_utf8_lossy(&bytes);
        assert!(s.contains("xref"));
        assert!(s.contains("trailer"));
        assert!(s.contains("startxref"));
    }

    #[test]
    fn test_write_pdf_has_catalog() {
        let doc = make_simple_doc();
        let bytes = write_pdf(&doc);
        let s = String::from_utf8_lossy(&bytes);
        assert!(s.contains("/Type /Catalog"));
        assert!(s.contains("/Type /Pages"));
        assert!(s.contains("/Type /Page"));
    }

    #[test]
    fn test_write_pdf_info() {
        let doc = make_simple_doc();
        let bytes = write_pdf(&doc);
        let s = String::from_utf8_lossy(&bytes);
        assert!(s.contains("/Title (Test Document)"));
        assert!(s.contains("/Author (Oxi)"));
        assert!(s.contains("/Producer (oxipdf-core"));
    }

    #[test]
    fn test_roundtrip_structure() {
        let doc = make_simple_doc();
        let bytes = write_pdf(&doc);

        // The output should be parseable by our own parser.
        let parsed = crate::parse_pdf(&bytes).unwrap();
        assert_eq!(parsed.version, PdfVersion::new(1, 7));
        assert_eq!(parsed.pages.len(), 1);
        assert!((parsed.pages[0].width - 612.0).abs() < 0.01);
        assert!((parsed.pages[0].height - 792.0).abs() < 0.01);
    }

    #[test]
    fn test_roundtrip_text() {
        let doc = make_simple_doc();
        let bytes = write_pdf(&doc);

        let parsed = crate::parse_pdf(&bytes).unwrap();
        assert!(!parsed.pages[0].contents.is_empty());
        match &parsed.pages[0].contents[0] {
            ContentElement::Text(span) => {
                assert_eq!(span.text, "Hello, PDF!");
                assert_eq!(span.font_size, 12.0);
            }
            _ => panic!("expected text element"),
        }
    }

    #[test]
    fn test_escape_string() {
        assert_eq!(escape_pdf_string("hello"), "hello");
        assert_eq!(escape_pdf_string("a(b)c"), "a\\(b\\)c");
        assert_eq!(escape_pdf_string("a\\b"), "a\\\\b");
    }

    #[test]
    fn test_multi_page() {
        let doc = PdfDocument {
            version: PdfVersion::new(1, 7),
            info: DocumentInfo::default(),
            pages: vec![
                Page {
                    width: 612.0,
                    height: 792.0,
                    media_box: Rectangle { llx: 0.0, lly: 0.0, urx: 612.0, ury: 792.0 },
                    crop_box: None,
                    contents: vec![],
                    rotation: 0,
                },
                Page {
                    width: 595.0,
                    height: 842.0,
                    media_box: Rectangle { llx: 0.0, lly: 0.0, urx: 595.0, ury: 842.0 },
                    crop_box: None,
                    contents: vec![],
                    rotation: 0,
                },
            ],
            outline: Vec::new(),
        };
        let bytes = write_pdf(&doc);
        let parsed = crate::parse_pdf(&bytes).unwrap();
        assert_eq!(parsed.pages.len(), 2);
        assert!((parsed.pages[1].width - 595.0).abs() < 0.01);
    }

    #[test]
    fn test_write_pdf_with_image() {
        // 2x2 red pixel image (RGB, 8bpc).
        let pixel_data = vec![
            255, 0, 0, 255, 0, 0, // row 1: red, red
            255, 0, 0, 255, 0, 0, // row 2: red, red
        ];
        let doc = PdfDocument {
            version: PdfVersion::new(1, 7),
            info: DocumentInfo::default(),
            pages: vec![Page {
                width: 612.0,
                height: 792.0,
                media_box: Rectangle { llx: 0.0, lly: 0.0, urx: 612.0, ury: 792.0 },
                crop_box: None,
                contents: vec![ContentElement::Image(ImageData {
                    x: 100.0,
                    y: 100.0,
                    width: 200.0,
                    height: 200.0,
                    data: pixel_data,
                    color_space: ColorSpace::DeviceRgb,
                    bits_per_component: 8,
                })],
                rotation: 0,
            }],
            outline: Vec::new(),
        };
        let bytes = write_pdf(&doc);
        let s = String::from_utf8_lossy(&bytes);
        assert!(s.contains("/Type /XObject"));
        assert!(s.contains("/Subtype /Image"));
        assert!(s.contains("/ColorSpace /DeviceRGB"));
        assert!(s.contains("/Im0"));

        // Should be parseable.
        let parsed = crate::parse_pdf(&bytes).unwrap();
        assert_eq!(parsed.pages.len(), 1);
    }

    #[test]
    fn test_write_pdf_japanese_cidfont() {
        let doc = PdfDocument {
            version: PdfVersion::new(1, 7),
            info: DocumentInfo {
                title: Some("日本語テスト".into()),
                ..Default::default()
            },
            pages: vec![Page {
                width: 595.0,
                height: 842.0,
                media_box: Rectangle { llx: 0.0, lly: 0.0, urx: 595.0, ury: 842.0 },
                crop_box: None,
                contents: vec![ContentElement::Text(TextSpan {
                    x: 72.0,
                    y: 72.0,
                    text: "こんにちは世界".into(),
                    font_name: "MSGothic".into(),
                    font_size: 12.0,
                    fill_color: Color::Gray(0.0),
                })],
                rotation: 0,
            }],
            outline: Vec::new(),
        };
        let bytes = write_pdf(&doc);
        let s = String::from_utf8_lossy(&bytes);
        // Should use Type0/CIDFont, not Type1
        assert!(s.contains("/Subtype /Type0"), "expected Type0 font");
        assert!(s.contains("/Encoding /Identity-H"), "expected Identity-H encoding");
        assert!(s.contains("/Subtype /CIDFontType2"), "expected CIDFontType2");
        assert!(s.contains("/ToUnicode"), "expected ToUnicode reference");
        // Content stream should have hex string, not parenthesized string
        // Verify the PDF is structurally valid by parsing
        let parsed = crate::parse_pdf(&bytes).unwrap();
        assert_eq!(parsed.pages.len(), 1);
    }

    #[test]
    fn test_write_pdf_mixed_ascii_and_japanese() {
        let doc = PdfDocument {
            version: PdfVersion::new(1, 7),
            info: DocumentInfo::default(),
            pages: vec![Page {
                width: 612.0,
                height: 792.0,
                media_box: Rectangle { llx: 0.0, lly: 0.0, urx: 612.0, ury: 792.0 },
                crop_box: None,
                contents: vec![
                    ContentElement::Text(TextSpan {
                        x: 72.0, y: 72.0,
                        text: "Hello".into(),
                        font_name: "Helvetica".into(),
                        font_size: 12.0,
                        fill_color: Color::Gray(0.0),
                    }),
                    ContentElement::Text(TextSpan {
                        x: 72.0, y: 100.0,
                        text: "日本語テキスト".into(),
                        font_name: "MSGothic".into(),
                        font_size: 12.0,
                        fill_color: Color::Gray(0.0),
                    }),
                ],
                rotation: 0,
            }],
            outline: Vec::new(),
        };
        let bytes = write_pdf(&doc);
        let s = String::from_utf8_lossy(&bytes);
        // Helvetica should be Type1, MSGothic should be Type0
        assert!(s.contains("/Subtype /Type1"), "expected Type1 for ASCII font");
        assert!(s.contains("/Subtype /Type0"), "expected Type0 for CJK font");
    }

    #[test]
    fn test_encode_utf16be_hex() {
        // ASCII
        assert_eq!(encode_utf16be_hex("A"), "0041");
        // Japanese
        assert_eq!(encode_utf16be_hex("あ"), "3042");
        // Mixed
        assert_eq!(encode_utf16be_hex("Aあ"), "00413042");
    }

    #[test]
    fn test_write_pdf_font_resources() {
        let doc = make_simple_doc();
        let bytes = write_pdf(&doc);
        let s = String::from_utf8_lossy(&bytes);
        assert!(s.contains("/Font <<"));
        assert!(s.contains("/F1"));
        assert!(s.contains("/BaseFont /F1"));
    }
}
