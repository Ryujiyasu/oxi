use std::collections::HashMap;

use crate::error::PdfError;
use crate::ir::*;
use super::object::*;
use super::xref::*;
use super::cmap::{CMap, parse_cmap};
use super::content::{interpret_content_stream_with_resources, PageResources, XObjectData};
use super::encoding::FontEncoding;
use super::filter::decode_stream;

/// Parse a PDF file from bytes into a `PdfDocument`.
pub fn parse_pdf(data: &[u8]) -> Result<PdfDocument, PdfError> {
    let version = parse_header(data)?;
    let startxref = find_startxref(data)?;

    let (xref, trailer) = parse_xref_and_trailer(data, startxref)?;

    let catalog = read_object_at(data, &xref, trailer.root.num)?;
    let pages = parse_pages(data, &xref, &catalog)?;

    let info = if let Some(info_ref) = trailer.info {
        let info_obj = read_object_at(data, &xref, info_ref.num)?;
        parse_document_info(&info_obj)
    } else {
        DocumentInfo::default()
    };

    Ok(PdfDocument {
        version,
        info,
        pages,
        outline: Vec::new(),
        embedded_fonts: std::collections::HashMap::new(),
    })
}

/// Detect whether the xref at `offset` is a traditional table or an xref stream,
/// and parse accordingly.
fn parse_xref_and_trailer(
    data: &[u8],
    offset: u64,
) -> Result<(XrefTable, Trailer), PdfError> {
    let pos = offset as usize;

    // Check if this is a traditional xref table (starts with "xref")
    if pos + 4 <= data.len() && &data[pos..pos + 4] == b"xref" {
        let (xref, trailer_pos) = parse_xref_table(data, offset)?;
        let trailer = parse_trailer(data, trailer_pos)?;
        Ok((xref, trailer))
    } else {
        // Must be an xref stream (PDF 1.5+): an indirect object containing a stream
        // with /Type /XRef
        parse_xref_stream_at(data, pos)
    }
}

/// Parse an xref stream object at the given position.
fn parse_xref_stream_at(data: &[u8], pos: usize) -> Result<(XrefTable, Trailer), PdfError> {
    // Skip the "N G obj" header
    let mut p = pos;

    // object number
    while p < data.len() && data[p].is_ascii_digit() {
        p += 1;
    }
    p = skip_ws(data, p);
    // generation number
    while p < data.len() && data[p].is_ascii_digit() {
        p += 1;
    }
    p = skip_ws(data, p);
    // "obj"
    if p + 3 <= data.len() && &data[p..p + 3] == b"obj" {
        p += 3;
    }
    p = skip_ws(data, p);

    let (obj, _) = parse_raw_object(data, p)?;

    match obj {
        PdfObject::Stream { ref dict, data: ref stream_data } => {
            let decoded = decode_stream(dict, stream_data)?;
            parse_xref_stream(dict, &decoded)
        }
        _ => Err(PdfError::Parse("expected xref stream object".into())),
    }
}

/// Parse the PDF header line: `%PDF-X.Y`
fn parse_header(data: &[u8]) -> Result<PdfVersion, PdfError> {
    if data.len() < 8 || &data[0..5] != b"%PDF-" {
        return Err(PdfError::Parse("not a PDF file".into()));
    }
    let major = data[5]
        .checked_sub(b'0')
        .ok_or_else(|| PdfError::Parse("invalid PDF major version".into()))?;
    // data[6] should be '.'
    let minor = data[7]
        .checked_sub(b'0')
        .ok_or_else(|| PdfError::Parse("invalid PDF minor version".into()))?;
    Ok(PdfVersion::new(major, minor))
}

/// Parse the trailer dictionary after the xref table.
fn parse_trailer(data: &[u8], mut pos: usize) -> Result<Trailer, PdfError> {
    // Skip to "trailer" keyword.
    let marker = b"trailer";
    if pos + marker.len() > data.len() || &data[pos..pos + marker.len()] != marker {
        return Err(PdfError::Parse("expected 'trailer' keyword".into()));
    }
    pos += marker.len();
    pos = skip_ws(data, pos);

    let (dict, _) = parse_raw_object(data, pos)?;

    let root = dict
        .dict_get("Root")
        .and_then(|o| o.as_ref())
        .ok_or_else(|| PdfError::Parse("trailer missing /Root".into()))?;

    let size = dict
        .dict_get("Size")
        .and_then(|o| o.as_i64())
        .unwrap_or(0) as u32;

    let info = dict.dict_get("Info").and_then(|o| o.as_ref());

    let prev = dict
        .dict_get("Prev")
        .and_then(|o| o.as_i64())
        .map(|v| v as u64);

    Ok(Trailer {
        size,
        root,
        info,
        id: None,
        prev,
    })
}

/// Read an indirect object given its object number.
fn read_object_at(data: &[u8], xref: &XrefTable, obj_num: u32) -> Result<PdfObject, PdfError> {
    match xref.entries.get(&obj_num) {
        Some(XrefEntry::InUse { offset, .. }) => {
            read_object_at_offset(data, *offset as usize)
        }
        Some(XrefEntry::Compressed { stream_obj, index }) => {
            read_object_from_object_stream(data, xref, *stream_obj, *index)
        }
        _ => Err(PdfError::Parse(format!("object {obj_num} not in xref"))),
    }
}

/// Read an object at a direct byte offset (traditional in-use object).
fn read_object_at_offset(data: &[u8], mut pos: usize) -> Result<PdfObject, PdfError> {
    // Skip "N G obj" header.
    // object number
    while pos < data.len() && data[pos].is_ascii_digit() {
        pos += 1;
    }
    pos = skip_ws(data, pos);
    // generation number
    while pos < data.len() && data[pos].is_ascii_digit() {
        pos += 1;
    }
    pos = skip_ws(data, pos);
    // "obj"
    if pos + 3 <= data.len() && &data[pos..pos + 3] == b"obj" {
        pos += 3;
    }
    pos = skip_ws(data, pos);

    let (obj, _) = parse_raw_object(data, pos)?;
    Ok(obj)
}

/// Read an object from an object stream (PDF 1.5+, /Type /ObjStm).
fn read_object_from_object_stream(
    data: &[u8],
    xref: &XrefTable,
    stream_obj_num: u32,
    index: u16,
) -> Result<PdfObject, PdfError> {
    // The object stream itself must be an InUse entry
    let stream_offset = match xref.entries.get(&stream_obj_num) {
        Some(XrefEntry::InUse { offset, .. }) => *offset as usize,
        _ => {
            return Err(PdfError::Parse(format!(
                "object stream {stream_obj_num} not found"
            )))
        }
    };

    let stream_obj = read_object_at_offset(data, stream_offset)?;

    match stream_obj {
        PdfObject::Stream {
            ref dict,
            data: ref stream_data,
        } => {
            use super::filter::decode_stream;
            let decoded = decode_stream(dict, stream_data)?;

            // /N = number of objects, /First = byte offset of first object data
            let n = dict
                .iter()
                .find(|(k, _)| k == "N")
                .and_then(|(_, v)| v.as_i64())
                .unwrap_or(0) as usize;

            let first = dict
                .iter()
                .find(|(k, _)| k == "First")
                .and_then(|(_, v)| v.as_i64())
                .unwrap_or(0) as usize;

            if (index as usize) >= n {
                return Err(PdfError::Parse(format!(
                    "object stream index {index} >= N={n}"
                )));
            }

            // The first part of the decoded data contains N pairs of (obj_num, byte_offset)
            // as space-separated integers. Parse them.
            let header_part = &decoded[..first.min(decoded.len())];
            let header_str = String::from_utf8_lossy(header_part);
            let tokens: Vec<&str> = header_str.split_whitespace().collect();

            // Each object has 2 tokens: obj_num, offset_from_first
            let pair_idx = (index as usize) * 2 + 1; // +1 to get the offset (skip obj_num)
            if pair_idx >= tokens.len() {
                return Err(PdfError::Parse(
                    "object stream header truncated".into(),
                ));
            }

            let obj_offset: usize = tokens[pair_idx]
                .parse()
                .map_err(|_| PdfError::Parse("invalid object stream offset".into()))?;

            let abs_offset = first + obj_offset;
            if abs_offset >= decoded.len() {
                return Err(PdfError::Parse(
                    "object stream offset out of bounds".into(),
                ));
            }

            let (obj, _) = parse_raw_object(&decoded, abs_offset)?;
            Ok(obj)
        }
        _ => Err(PdfError::Parse(format!(
            "object {stream_obj_num} is not a stream"
        ))),
    }
}

/// Resolve a reference by following it through the xref table.
fn resolve(
    data: &[u8],
    xref: &XrefTable,
    obj: &PdfObject,
) -> Result<PdfObject, PdfError> {
    match obj {
        PdfObject::Reference(r) => read_object_at(data, xref, r.num),
        other => Ok(other.clone()),
    }
}

/// Parse the page tree starting from the catalog's /Pages reference.
fn parse_pages(
    data: &[u8],
    xref: &XrefTable,
    catalog: &PdfObject,
) -> Result<Vec<Page>, PdfError> {
    let pages_ref = catalog
        .dict_get("Pages")
        .ok_or_else(|| PdfError::Parse("catalog missing /Pages".into()))?;
    let pages_obj = resolve(data, xref, pages_ref)?;

    let mut result = Vec::new();
    collect_pages(data, xref, &pages_obj, &None, &mut result)?;
    Ok(result)
}

/// Recursively walk the page tree (Pages nodes can be nested).
fn collect_pages(
    data: &[u8],
    xref: &XrefTable,
    node: &PdfObject,
    inherited_media_box: &Option<Rectangle>,
    out: &mut Vec<Page>,
) -> Result<(), PdfError> {
    let node_type = node.dict_get("Type").and_then(|o| o.as_name());

    // Inherit MediaBox from parent if present.
    let media_box = parse_rectangle(node.dict_get("MediaBox")).or(*inherited_media_box);

    match node_type {
        Some("Pages") => {
            if let Some(kids) = node.dict_get("Kids").and_then(|o| o.as_array()) {
                for kid_ref in kids {
                    let kid = resolve(data, xref, kid_ref)?;
                    collect_pages(data, xref, &kid, &media_box, out)?;
                }
            }
        }
        Some("Page") | None => {
            let mbox = media_box.unwrap_or(Rectangle {
                llx: 0.0,
                lly: 0.0,
                urx: 612.0,
                ury: 792.0, // US Letter default
            });
            let crop_box = parse_rectangle(node.dict_get("CropBox"));
            let rotation = node
                .dict_get("Rotate")
                .and_then(|o| o.as_i64())
                .unwrap_or(0) as u16;

            // Parse content stream(s) for this page.
            let contents = parse_page_contents(data, xref, node, mbox.height());

            out.push(Page {
                width: mbox.width(),
                height: mbox.height(),
                media_box: mbox,
                crop_box,
                contents,
                rotation,
            });
        }
        _ => {}
    }

    Ok(())
}

/// Extract and decode the content stream(s) for a page, then interpret them.
fn parse_page_contents(
    data: &[u8],
    xref: &XrefTable,
    page_node: &PdfObject,
    _page_height: f64,
) -> Vec<ContentElement> {
    // Build full page resources (CMaps, encodings, XObjects).
    let resources = build_page_resources(data, xref, page_node);

    let contents_obj = match page_node.dict_get("Contents") {
        Some(obj) => obj,
        None => return Vec::new(),
    };

    let stream_data = match contents_obj {
        PdfObject::Reference(r) => {
            read_object_at(data, xref, r.num)
                .ok()
                .and_then(|obj| decode_stream_object(&obj))
        }
        PdfObject::Array(refs) => {
            let mut combined = Vec::new();
            for item in refs {
                let resolved = match item {
                    PdfObject::Reference(r) => read_object_at(data, xref, r.num).ok(),
                    other => Some(other.clone()),
                };
                if let Some(obj) = resolved {
                    if let Some(decoded) = decode_stream_object(&obj) {
                        combined.extend_from_slice(&decoded);
                        combined.push(b' ');
                    }
                }
            }
            Some(combined)
        }
        _ => None,
    };

    match stream_data {
        Some(raw) => interpret_content_stream_with_resources(&raw, &resources)
            .unwrap_or_default(),
        None => Vec::new(),
    }
}

/// Build full PageResources from a page/XObject node's Resources dictionary.
fn build_page_resources(
    data: &[u8],
    xref: &XrefTable,
    node: &PdfObject,
) -> PageResources {
    let resources_obj = match node.dict_get("Resources") {
        Some(PdfObject::Reference(r)) => read_object_at(data, xref, r.num).ok(),
        Some(obj) => Some(obj.clone()),
        None => None,
    };

    let resources_obj = match resources_obj {
        Some(r) => r,
        None => return PageResources::default(),
    };

    let font_cmaps = extract_font_cmaps_from(data, xref, &resources_obj);
    let font_encodings = extract_font_encodings(data, xref, &resources_obj);
    let font_base_names = extract_font_base_names(data, xref, &resources_obj);
    let xobject_streams = extract_xobject_streams(data, xref, &resources_obj);

    PageResources {
        font_cmaps,
        font_encodings,
        font_base_names,
        xobject_streams,
    }
}

/// Extract ToUnicode CMaps from a resolved Resources dictionary.
fn extract_font_cmaps_from(
    data: &[u8],
    xref: &XrefTable,
    resources: &PdfObject,
) -> HashMap<String, CMap> {
    let mut cmaps = HashMap::new();

    let font_dict = match resources.dict_get("Font") {
        Some(PdfObject::Reference(r)) => read_object_at(data, xref, r.num).ok(),
        Some(obj) => Some(obj.clone()),
        None => None,
    };

    let font_dict = match font_dict {
        Some(d) => d,
        None => return cmaps,
    };

    if let Some(entries) = font_dict.as_dict() {
        for (font_name, font_ref) in entries {
            let font_obj = match font_ref {
                PdfObject::Reference(r) => read_object_at(data, xref, r.num).ok(),
                obj => Some(obj.clone()),
            };

            if let Some(font_obj) = font_obj {
                if let Some(to_unicode_ref) = font_obj.dict_get("ToUnicode") {
                    let to_unicode = match to_unicode_ref {
                        PdfObject::Reference(r) => read_object_at(data, xref, r.num).ok(),
                        obj => Some(obj.clone()),
                    };

                    if let Some(to_unicode) = to_unicode {
                        if let Some(cmap_data) = decode_stream_object(&to_unicode) {
                            let cmap = parse_cmap(&cmap_data);
                            cmaps.insert(font_name.clone(), cmap);
                        }
                    }
                }
            }
        }
    }

    cmaps
}

/// Extract BaseFont names from Resources → Font dictionary.
/// Maps resource keys like "F1" to actual font names like "Helvetica-Bold".
fn extract_font_base_names(
    data: &[u8],
    xref: &XrefTable,
    resources: &PdfObject,
) -> HashMap<String, String> {
    let mut names = HashMap::new();

    let font_dict = match resources.dict_get("Font") {
        Some(PdfObject::Reference(r)) => read_object_at(data, xref, r.num).ok(),
        Some(obj) => Some(obj.clone()),
        None => None,
    };

    let font_dict = match font_dict {
        Some(d) => d,
        None => return names,
    };

    if let Some(entries) = font_dict.as_dict() {
        for (font_name, font_ref) in entries {
            let font_obj = match font_ref {
                PdfObject::Reference(r) => read_object_at(data, xref, r.num).ok(),
                obj => Some(obj.clone()),
            };

            if let Some(font_obj) = font_obj {
                // Try /BaseFont first
                if let Some(base_font) = font_obj.dict_get("BaseFont") {
                    if let Some(name) = base_font.as_name() {
                        names.insert(font_name.clone(), name.to_string());
                        continue;
                    }
                }
                // For Type0 (composite) fonts, check DescendantFonts
                if let Some(descendants) = font_obj.dict_get("DescendantFonts") {
                    let arr = match descendants {
                        PdfObject::Reference(r) => read_object_at(data, xref, r.num).ok(),
                        obj => Some(obj.clone()),
                    };
                    if let Some(PdfObject::Array(arr)) = arr {
                        if let Some(first) = arr.first() {
                            let desc_font = match first {
                                PdfObject::Reference(r) => read_object_at(data, xref, r.num).ok(),
                                obj => Some(obj.clone()),
                            };
                            if let Some(desc_font) = desc_font {
                                if let Some(base_font) = desc_font.dict_get("BaseFont") {
                                    if let Some(name) = base_font.as_name() {
                                        names.insert(font_name.clone(), name.to_string());
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    names
}

/// Extract font encodings from Resources → Font dictionary.
fn extract_font_encodings(
    data: &[u8],
    xref: &XrefTable,
    resources: &PdfObject,
) -> HashMap<String, FontEncoding> {
    let mut encodings = HashMap::new();

    let font_dict = match resources.dict_get("Font") {
        Some(PdfObject::Reference(r)) => read_object_at(data, xref, r.num).ok(),
        Some(obj) => Some(obj.clone()),
        None => None,
    };

    let font_dict = match font_dict {
        Some(d) => d,
        None => return encodings,
    };

    if let Some(entries) = font_dict.as_dict() {
        for (font_name, font_ref) in entries {
            let font_obj = match font_ref {
                PdfObject::Reference(r) => read_object_at(data, xref, r.num).ok(),
                obj => Some(obj.clone()),
            };

            if let Some(font_obj) = font_obj {
                if let Some(enc) = parse_font_encoding(data, xref, &font_obj) {
                    encodings.insert(font_name.clone(), enc);
                }
            }
        }
    }

    encodings
}

/// Parse the /Encoding entry of a font dictionary.
fn parse_font_encoding(
    data: &[u8],
    xref: &XrefTable,
    font_obj: &PdfObject,
) -> Option<FontEncoding> {
    let enc = font_obj.dict_get("Encoding")?;

    match enc {
        PdfObject::Name(name) => match name.as_str() {
            "WinAnsiEncoding" => Some(FontEncoding::WinAnsi),
            "MacRomanEncoding" => Some(FontEncoding::MacRoman),
            _ => None,
        },
        PdfObject::Reference(r) => {
            let enc_obj = read_object_at(data, xref, r.num).ok()?;
            let base = enc_obj
                .dict_get("BaseEncoding")
                .and_then(|o| o.as_name())
                .unwrap_or("");
            match base {
                "WinAnsiEncoding" => Some(FontEncoding::WinAnsi),
                "MacRomanEncoding" => Some(FontEncoding::MacRoman),
                _ => Some(FontEncoding::WinAnsi), // default
            }
        }
        PdfObject::Dictionary(_) => {
            let base = enc
                .dict_get("BaseEncoding")
                .and_then(|o| o.as_name())
                .unwrap_or("");
            match base {
                "WinAnsiEncoding" => Some(FontEncoding::WinAnsi),
                "MacRomanEncoding" => Some(FontEncoding::MacRoman),
                _ => Some(FontEncoding::WinAnsi),
            }
        }
        _ => None,
    }
}

/// Extract Form XObject streams from Resources → XObject dictionary.
fn extract_xobject_streams(
    data: &[u8],
    xref: &XrefTable,
    resources: &PdfObject,
) -> HashMap<String, XObjectData> {
    let mut xobjects = HashMap::new();

    let xobj_dict = match resources.dict_get("XObject") {
        Some(PdfObject::Reference(r)) => read_object_at(data, xref, r.num).ok(),
        Some(obj) => Some(obj.clone()),
        None => None,
    };

    let xobj_dict = match xobj_dict {
        Some(d) => d,
        None => return xobjects,
    };

    if let Some(entries) = xobj_dict.as_dict() {
        for (name, xobj_ref) in entries {
            let xobj = match xobj_ref {
                PdfObject::Reference(r) => read_object_at(data, xref, r.num).ok(),
                obj => Some(obj.clone()),
            };

            if let Some(xobj) = xobj {
                if let Some(xobj_data) = parse_form_xobject(data, xref, &xobj) {
                    xobjects.insert(name.clone(), xobj_data);
                }
            }
        }
    }

    xobjects
}

/// Parse a Form XObject (stream with /Subtype /Form).
fn parse_form_xobject(
    data: &[u8],
    xref: &XrefTable,
    obj: &PdfObject,
) -> Option<XObjectData> {
    match obj {
        PdfObject::Stream { dict, data: stream_data } => {
            // Only process Form XObjects (not Image XObjects)
            let subtype = dict
                .iter()
                .find(|(k, _)| k == "Subtype")
                .and_then(|(_, v)| v.as_name());

            if subtype != Some("Form") {
                return None;
            }

            let decoded = decode_stream(dict, stream_data).ok()?;

            // Extract /Matrix (default: identity)
            let matrix = dict
                .iter()
                .find(|(k, _)| k == "Matrix")
                .and_then(|(_, v)| v.as_array())
                .map(|arr| {
                    let mut m = [1.0, 0.0, 0.0, 1.0, 0.0, 0.0];
                    for (i, val) in arr.iter().take(6).enumerate() {
                        m[i] = val.as_f64().unwrap_or(if i == 0 || i == 3 { 1.0 } else { 0.0 });
                    }
                    m
                })
                .unwrap_or([1.0, 0.0, 0.0, 1.0, 0.0, 0.0]);

            // Extract /BBox
            let bbox = dict
                .iter()
                .find(|(k, _)| k == "BBox")
                .and_then(|(_, v)| v.as_array())
                .and_then(|arr| {
                    if arr.len() >= 4 {
                        Some(Rectangle {
                            llx: arr[0].as_f64()?,
                            lly: arr[1].as_f64()?,
                            urx: arr[2].as_f64()?,
                            ury: arr[3].as_f64()?,
                        })
                    } else {
                        None
                    }
                });

            // Extract sub-resources from the XObject's own /Resources
            let sub_resources = dict
                .iter()
                .find(|(k, _)| k == "Resources")
                .and_then(|(_, v)| {
                    let res_obj = match v {
                        PdfObject::Reference(r) => read_object_at(data, xref, r.num).ok(),
                        obj => Some(obj.clone()),
                    };
                    res_obj.map(|r| {
                        let cmaps = extract_font_cmaps_from(data, xref, &r);
                        let encs = extract_font_encodings(data, xref, &r);
                        let base_names = extract_font_base_names(data, xref, &r);
                        let xobjs = extract_xobject_streams(data, xref, &r);
                        Box::new(PageResources {
                            font_cmaps: cmaps,
                            font_encodings: encs,
                            font_base_names: base_names,
                            xobject_streams: xobjs,
                        })
                    })
                });

            Some(XObjectData {
                stream: decoded,
                matrix,
                bbox,
                resources: sub_resources,
            })
        }
        _ => None,
    }
}

/// Decode a stream object (apply filters) and return the raw bytes.
fn decode_stream_object(obj: &PdfObject) -> Option<Vec<u8>> {
    match obj {
        PdfObject::Stream { dict, data } => decode_stream(dict, data).ok(),
        _ => None,
    }
}

fn parse_rectangle(obj: Option<&PdfObject>) -> Option<Rectangle> {
    let arr = obj?.as_array()?;
    if arr.len() < 4 {
        return None;
    }
    Some(Rectangle {
        llx: arr[0].as_f64()?,
        lly: arr[1].as_f64()?,
        urx: arr[2].as_f64()?,
        ury: arr[3].as_f64()?,
    })
}

fn parse_document_info(obj: &PdfObject) -> DocumentInfo {
    let get_str = |key: &str| -> Option<String> {
        obj.dict_get(key)
            .and_then(|o| o.as_bytes())
            .and_then(|b| String::from_utf8(b.to_vec()).ok())
    };
    DocumentInfo {
        title: get_str("Title"),
        author: get_str("Author"),
        subject: get_str("Subject"),
        keywords: get_str("Keywords"),
        creator: get_str("Creator"),
        producer: get_str("Producer"),
        creation_date: get_str("CreationDate"),
        mod_date: get_str("ModDate"),
    }
}

// ---------------------------------------------------------------------------
// Minimal recursive-descent parser for PDF objects
// ---------------------------------------------------------------------------

/// Maximum nesting depth for PDF objects (dicts/arrays).
const MAX_PARSE_DEPTH: usize = 64;

/// Parse a single PDF object starting at `pos`, return (object, end_pos).
fn parse_raw_object(data: &[u8], pos: usize) -> Result<(PdfObject, usize), PdfError> {
    parse_raw_object_depth(data, pos, 0)
}

fn parse_raw_object_depth(data: &[u8], mut pos: usize, depth: usize) -> Result<(PdfObject, usize), PdfError> {
    if depth > MAX_PARSE_DEPTH {
        return Err(PdfError::Parse(format!(
            "object nesting too deep (>{MAX_PARSE_DEPTH})"
        )));
    }
    pos = skip_ws(data, pos);
    if pos >= data.len() {
        return Err(PdfError::Parse("unexpected end of data".into()));
    }

    match data[pos] {
        b'<' if pos + 1 < data.len() && data[pos + 1] == b'<' => parse_dict_depth(data, pos, depth),
        b'<' => parse_hex_string(data, pos),
        b'(' => parse_literal_string(data, pos),
        b'[' => parse_array_depth(data, pos, depth),
        b'/' => parse_name(data, pos),
        b't' if data[pos..].starts_with(b"true") => Ok((PdfObject::Boolean(true), pos + 4)),
        b'f' if data[pos..].starts_with(b"false") => Ok((PdfObject::Boolean(false), pos + 5)),
        b'n' if data[pos..].starts_with(b"null") => Ok((PdfObject::Null, pos + 4)),
        b'+' | b'-' | b'.' | b'0'..=b'9' => parse_number_or_ref(data, pos),
        _ => Err(PdfError::Parse(format!(
            "unexpected byte '{}' at offset {pos}",
            data[pos] as char
        ))),
    }
}

fn parse_dict_depth(data: &[u8], mut pos: usize, depth: usize) -> Result<(PdfObject, usize), PdfError> {
    pos += 2; // skip "<<"
    pos = skip_ws(data, pos);

    let mut entries = Vec::new();

    while pos + 1 < data.len() && !(data[pos] == b'>' && data[pos + 1] == b'>') {
        // Key must be a name.
        let (key_obj, new_pos) = parse_name(data, pos)?;
        let key = match key_obj {
            PdfObject::Name(s) => s,
            _ => return Err(PdfError::Parse("dict key must be a name".into())),
        };
        pos = skip_ws(data, new_pos);

        // Value.
        let (val, new_pos) = parse_raw_object_depth(data, pos, depth + 1)?;
        pos = skip_ws(data, new_pos);

        entries.push((key, val));
    }

    if pos + 1 < data.len() {
        pos += 2; // skip ">>"
    }
    pos = skip_ws(data, pos);

    // Check for stream.
    if pos + 6 <= data.len() && &data[pos..pos + 6] == b"stream" {
        pos += 6;
        // Skip CR, LF, or CRLF after "stream".
        if pos < data.len() && data[pos] == b'\r' {
            pos += 1;
        }
        if pos < data.len() && data[pos] == b'\n' {
            pos += 1;
        }

        let length = entries
            .iter()
            .find(|(k, _)| k == "Length")
            .and_then(|(_, v)| v.as_i64())
            .unwrap_or(0) as usize;

        let stream_data = if pos + length <= data.len() {
            data[pos..pos + length].to_vec()
        } else {
            Vec::new()
        };
        pos += length;

        // Skip "endstream".
        pos = skip_ws(data, pos);
        if pos + 9 <= data.len() && &data[pos..pos + 9] == b"endstream" {
            pos += 9;
        }

        return Ok((
            PdfObject::Stream {
                dict: entries,
                data: stream_data,
            },
            pos,
        ));
    }

    Ok((PdfObject::Dictionary(entries), pos))
}

fn parse_array_depth(data: &[u8], mut pos: usize, depth: usize) -> Result<(PdfObject, usize), PdfError> {
    pos += 1; // skip '['
    pos = skip_ws(data, pos);

    let mut items = Vec::new();
    while pos < data.len() && data[pos] != b']' {
        let (obj, new_pos) = parse_raw_object_depth(data, pos, depth + 1)?;
        items.push(obj);
        pos = skip_ws(data, new_pos);
    }
    if pos < data.len() {
        pos += 1; // skip ']'
    }
    Ok((PdfObject::Array(items), pos))
}

fn parse_name(data: &[u8], mut pos: usize) -> Result<(PdfObject, usize), PdfError> {
    if data[pos] != b'/' {
        return Err(PdfError::Parse("expected '/' for name".into()));
    }
    pos += 1;
    let start = pos;
    while pos < data.len()
        && !data[pos].is_ascii_whitespace()
        && !matches!(data[pos], b'/' | b'<' | b'>' | b'[' | b']' | b'(' | b')' | b'{' | b'}')
    {
        pos += 1;
    }
    let name = String::from_utf8_lossy(&data[start..pos]).into_owned();
    Ok((PdfObject::Name(name), pos))
}

fn parse_literal_string(data: &[u8], mut pos: usize) -> Result<(PdfObject, usize), PdfError> {
    pos += 1; // skip '('
    let mut result = Vec::new();
    let mut depth = 1u32;

    while pos < data.len() && depth > 0 {
        match data[pos] {
            b'(' => {
                depth += 1;
                result.push(b'(');
            }
            b')' => {
                depth -= 1;
                if depth > 0 {
                    result.push(b')');
                }
            }
            b'\\' => {
                pos += 1;
                if pos < data.len() {
                    match data[pos] {
                        b'n' => result.push(b'\n'),
                        b'r' => result.push(b'\r'),
                        b't' => result.push(b'\t'),
                        b'(' => result.push(b'('),
                        b')' => result.push(b')'),
                        b'\\' => result.push(b'\\'),
                        _ => result.push(data[pos]),
                    }
                }
            }
            b => result.push(b),
        }
        pos += 1;
    }
    Ok((PdfObject::String(result), pos))
}

fn parse_hex_string(data: &[u8], mut pos: usize) -> Result<(PdfObject, usize), PdfError> {
    pos += 1; // skip '<'
    let mut hex = Vec::new();
    while pos < data.len() && data[pos] != b'>' {
        if !data[pos].is_ascii_whitespace() {
            hex.push(data[pos]);
        }
        pos += 1;
    }
    if pos < data.len() {
        pos += 1; // skip '>'
    }

    // Pad with 0 if odd number of hex digits.
    if hex.len() % 2 != 0 {
        hex.push(b'0');
    }

    let bytes: Vec<u8> = hex
        .chunks(2)
        .filter_map(|pair| {
            let s = std::str::from_utf8(pair).ok()?;
            u8::from_str_radix(s, 16).ok()
        })
        .collect();

    Ok((PdfObject::HexString(bytes), pos))
}

fn parse_number_or_ref(data: &[u8], pos: usize) -> Result<(PdfObject, usize), PdfError> {
    let (num_str, mut end) = read_token(data, pos);

    // Check if this could be an indirect reference: "N G R"
    let saved = end;
    let ws_end = skip_ws(data, end);
    if ws_end < data.len() && data[ws_end].is_ascii_digit() {
        let (gen_str, gen_end) = read_token(data, ws_end);
        let ws_end2 = skip_ws(data, gen_end);
        if ws_end2 < data.len() && data[ws_end2] == b'R' {
            if let (Ok(num), Ok(gen)) = (num_str.parse::<u32>(), gen_str.parse::<u16>()) {
                return Ok((PdfObject::Reference(ObjRef { num, gen }), ws_end2 + 1));
            }
        }
    }
    end = saved;

    // It's a number.
    if num_str.contains('.') {
        let val: f64 = num_str
            .parse()
            .map_err(|_| PdfError::Parse(format!("invalid real: {num_str}")))?;
        Ok((PdfObject::Real(val), end))
    } else {
        let val: i64 = num_str
            .parse()
            .map_err(|_| PdfError::Parse(format!("invalid integer: {num_str}")))?;
        Ok((PdfObject::Integer(val), end))
    }
}

fn read_token(data: &[u8], pos: usize) -> (String, usize) {
    let mut end = pos;
    while end < data.len()
        && !data[end].is_ascii_whitespace()
        && !matches!(data[end], b'/' | b'<' | b'>' | b'[' | b']' | b'(' | b')')
    {
        end += 1;
    }
    let s = String::from_utf8_lossy(&data[pos..end]).into_owned();
    (s, end)
}

fn skip_ws(data: &[u8], mut pos: usize) -> usize {
    while pos < data.len() {
        match data[pos] {
            b' ' | b'\t' | b'\r' | b'\n' | 0x0C | 0x00 => pos += 1,
            b'%' => {
                // Skip comment line.
                while pos < data.len() && data[pos] != b'\n' && data[pos] != b'\r' {
                    pos += 1;
                }
            }
            _ => break,
        }
    }
    pos
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_header() {
        let data = b"%PDF-1.7\n";
        let ver = parse_header(data).unwrap();
        assert_eq!(ver, PdfVersion::new(1, 7));
    }

    #[test]
    fn test_parse_header_2_0() {
        let data = b"%PDF-2.0\n";
        let ver = parse_header(data).unwrap();
        assert_eq!(ver, PdfVersion::new(2, 0));
    }

    #[test]
    fn test_parse_header_invalid() {
        assert!(parse_header(b"not a pdf").is_err());
    }

    #[test]
    fn test_parse_name() {
        let data = b"/Type";
        let (obj, _) = parse_name(data, 0).unwrap();
        assert_eq!(obj, PdfObject::Name("Type".into()));
    }

    #[test]
    fn test_parse_literal_string() {
        let data = b"(hello world)";
        let (obj, _) = parse_literal_string(data, 0).unwrap();
        assert_eq!(obj, PdfObject::String(b"hello world".to_vec()));
    }

    #[test]
    fn test_parse_nested_parens() {
        let data = b"(hello (nested) world)";
        let (obj, _) = parse_literal_string(data, 0).unwrap();
        assert_eq!(obj, PdfObject::String(b"hello (nested) world".to_vec()));
    }

    #[test]
    fn test_parse_hex_string() {
        let data = b"<48656C6C6F>";
        let (obj, _) = parse_hex_string(data, 0).unwrap();
        assert_eq!(obj, PdfObject::HexString(b"Hello".to_vec()));
    }

    #[test]
    fn test_parse_array() {
        let data = b"[1 2 3]";
        let (obj, _) = parse_array_depth(data, 0, 0).unwrap();
        assert_eq!(
            obj,
            PdfObject::Array(vec![
                PdfObject::Integer(1),
                PdfObject::Integer(2),
                PdfObject::Integer(3),
            ])
        );
    }

    #[test]
    fn test_parse_dict() {
        let data = b"<< /Type /Catalog /Pages 3 0 R >>";
        let (obj, _) = parse_dict_depth(data, 0, 0).unwrap();
        assert_eq!(
            obj.dict_get("Type"),
            Some(&PdfObject::Name("Catalog".into()))
        );
        assert_eq!(
            obj.dict_get("Pages"),
            Some(&PdfObject::Reference(ObjRef { num: 3, gen: 0 }))
        );
    }

    #[test]
    fn test_parse_number() {
        let (obj, _) = parse_number_or_ref(b"42 ", 0).unwrap();
        assert_eq!(obj, PdfObject::Integer(42));

        let (obj, _) = parse_number_or_ref(b"3.14 ", 0).unwrap();
        assert_eq!(obj, PdfObject::Real(3.14));
    }

    #[test]
    fn test_parse_reference() {
        let (obj, _) = parse_number_or_ref(b"10 0 R", 0).unwrap();
        assert_eq!(obj, PdfObject::Reference(ObjRef { num: 10, gen: 0 }));
    }

    #[test]
    fn test_find_startxref() {
        let data = b"%PDF-1.4\nsome content\nstartxref\n12345\n%%EOF";
        let offset = find_startxref(data).unwrap();
        assert_eq!(offset, 12345);
    }
}
