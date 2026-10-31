// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Static census of content Oxi does not draw.
//!
//! The layout parser walks past a chart, a SmartArt diagram or an altChunk
//! without a trace, so the PDF comes out with a hole and nothing says so. This
//! scan reads the story parts on their own and lists every such construct, so
//! a caller can refuse the result (`oxi docx-to-pdf --strict`) and send the
//! file to another engine. It never changes what is parsed or laid out.
//!
//! The rule for `mc:AlternateContent` mirrors the parser: inside a run, a
//! Choice that yields nothing drawable falls to a VML `w:pict` in the
//! Fallback, so an unsupported Choice with such a Fallback is drawn and not
//! reported. Everywhere else the Fallback is ignored, as the parser ignores
//! it. A wrong guess here should err towards reporting: a false report costs
//! a fallback to another engine, a missed one costs a silent hole.

use std::collections::BTreeMap;

use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;

use crate::ir::UnsupportedContent;

/// DrawingML graphic types the parser draws.
const DRAWN_GRAPHICS: &[&str] = &[
    "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
    "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
];

/// Whether `name` is a story part: one whose content Word renders on a page.
pub(crate) fn is_story_part(name: &str) -> bool {
    let story = |prefix: &str| {
        name.strip_prefix(prefix)
            .and_then(|rest| rest.strip_suffix(".xml"))
            .is_some_and(|n| n.chars().all(|c| c.is_ascii_digit()))
    };
    name == "word/document.xml"
        || name == "word/footnotes.xml"
        || name == "word/endnotes.xml"
        || story("word/header")
        || story("word/footer")
}

/// Census of the given story parts, as (part name, XML) pairs. Entries come
/// out in part order, then by kind.
pub(crate) fn census<'a>(parts: impl IntoIterator<Item = (&'a str, &'a str)>) -> Vec<UnsupportedContent> {
    let mut out = Vec::new();
    for (part, xml) in parts {
        for (kind, count) in scan_part(xml) {
            out.push(UnsupportedContent { kind, part: part.to_string(), count });
        }
    }
    out
}

fn graphic_kind(uri: &str) -> Option<String> {
    if DRAWN_GRAPHICS.contains(&uri) {
        return None;
    }
    Some(match uri {
        "http://schemas.openxmlformats.org/drawingml/2006/chart"
        | "http://schemas.microsoft.com/office/drawing/2014/chartex" => "chart".to_string(),
        "http://schemas.openxmlformats.org/drawingml/2006/diagram" => "smartart".to_string(),
        _ => format!("graphic:{uri}"),
    })
}

/// One open `mc:AlternateContent`.
struct AltFrame {
    /// Element depth of the AlternateContent itself.
    depth: usize,
    /// Its parent is a run, so the parser can fall to a VML Fallback.
    in_run: bool,
    /// Findings inside the Choice, held until the frame closes.
    choice: Vec<String>,
    fallback_has_pict: bool,
}

/// One open `w:object` (OLE): drawn only through its preview picture.
struct ObjectFrame {
    depth: usize,
    has_preview: bool,
}

fn scan_part(xml: &str) -> BTreeMap<String, usize> {
    let mut found: BTreeMap<String, usize> = BTreeMap::new();
    let mut reader = Reader::from_str(xml);
    // Local names of the open elements.
    let mut stack: Vec<String> = Vec::new();
    let mut alts: Vec<AltFrame> = Vec::new();
    let mut objects: Vec<ObjectFrame> = Vec::new();
    // Depth of the open Choice / Fallback of the innermost AlternateContent.
    let mut choice_depth: Option<usize> = None;
    let mut fallback_depth: Option<usize> = None;

    loop {
        let (e, empty) = match reader.read_event() {
            Ok(Event::Start(e)) => (e, false),
            Ok(Event::Empty(e)) => (e, true),
            Ok(Event::End(_)) => {
                let depth = stack.len();
                if objects.last().is_some_and(|o| o.depth == depth) {
                    let o = objects.pop().unwrap_or(ObjectFrame { depth, has_preview: true });
                    if !o.has_preview && fallback_depth.is_none() {
                        report(&mut found, &mut alts, choice_depth, "ole-without-preview".to_string());
                    }
                }
                if choice_depth == Some(depth) {
                    choice_depth = None;
                }
                if fallback_depth == Some(depth) {
                    fallback_depth = None;
                }
                if alts.last().is_some_and(|a| a.depth == depth) {
                    if let Some(a) = alts.pop() {
                        if !(a.in_run && a.fallback_has_pict) {
                            for kind in a.choice {
                                report(&mut found, &mut alts, None, kind);
                            }
                        }
                    }
                }
                stack.pop();
                continue;
            }
            Ok(Event::Eof) | Err(_) => break,
            Ok(_) => continue,
        };
        let name = local(&e);
        let depth = stack.len() + 1;
        let parent_is_run = stack.last().is_some_and(|p| p == "r");

        match name.as_str() {
            "AlternateContent" if !empty => alts.push(AltFrame {
                depth,
                in_run: parent_is_run,
                choice: Vec::new(),
                fallback_has_pict: false,
            }),
            "Choice" if !empty && alts.last().is_some_and(|a| a.depth + 1 == depth) => {
                choice_depth = Some(depth);
            }
            "Fallback" if !empty && alts.last().is_some_and(|a| a.depth + 1 == depth) => {
                fallback_depth = Some(depth);
            }
            "pict" if fallback_depth.is_some() => {
                if let Some(a) = alts.last_mut() {
                    a.fallback_has_pict = true;
                }
            }
            _ => {}
        }

        // The parser reads nothing inside a Fallback but a run-level VML pict,
        // so nothing found there is a hole.
        if fallback_depth.is_none() {
            let kind = match name.as_str() {
                "graphicData" => attr(&e, "uri").and_then(|uri| graphic_kind(&uri)),
                "altChunk" => Some("altChunk".to_string()),
                "subDoc" => Some("subDoc".to_string()),
                "contentPart" => Some("ink".to_string()),
                _ => None,
            };
            if let Some(kind) = kind {
                report(&mut found, &mut alts, choice_depth, kind);
            }
            match name.as_str() {
                "object" if !empty && stack.last().is_some_and(|p| p == "r") => {
                    objects.push(ObjectFrame { depth, has_preview: false });
                }
                "imagedata" | "drawing" | "shape" | "rect" | "line" => {
                    if let Some(o) = objects.last_mut() {
                        o.has_preview = true;
                    }
                }
                _ => {}
            }
        }

        if !empty {
            stack.push(name);
        }
    }
    found
}

/// Count `kind`, or hold it in the open Choice until its AlternateContent
/// shows whether a Fallback covers it.
fn report(found: &mut BTreeMap<String, usize>, alts: &mut [AltFrame], choice_depth: Option<usize>, kind: String) {
    match (choice_depth, alts.last_mut()) {
        (Some(_), Some(a)) => a.choice.push(kind),
        _ => *found.entry(kind).or_insert(0) += 1,
    }
}

fn local(e: &BytesStart) -> String {
    String::from_utf8_lossy(e.local_name().as_ref()).into_owned()
}

fn attr(e: &BytesStart, key: &str) -> Option<String> {
    e.attributes()
        .flatten()
        .find(|a| a.key.local_name().as_ref() == key.as_bytes())
        .map(|a| String::from_utf8_lossy(&a.value).into_owned())
}

#[cfg(test)]
mod tests {
    use super::*;

    const NS: &str = r#"xmlns:w="w" xmlns:a="a" xmlns:wp="wp" xmlns:mc="mc" xmlns:v="v" xmlns:o="o" xmlns:w14="w14""#;

    fn scan(body: &str) -> BTreeMap<String, usize> {
        scan_part(&format!(r#"<w:document {NS}><w:body>{body}</w:body></w:document>"#))
    }

    fn drawing(uri: &str) -> String {
        format!(r#"<w:drawing><wp:inline><a:graphic><a:graphicData uri="{uri}"/></a:graphic></wp:inline></w:drawing>"#)
    }

    #[test]
    fn drawn_graphics_are_not_reported() {
        let pic = drawing("http://schemas.openxmlformats.org/drawingml/2006/picture");
        assert!(scan(&format!("<w:p><w:r>{pic}</w:r></w:p>")).is_empty());
    }

    #[test]
    fn chart_smartart_and_unknown_graphics_are_reported() {
        let body = format!(
            "<w:p><w:r>{}</w:r><w:r>{}</w:r><w:r>{}</w:r><w:r>{}</w:r></w:p>",
            drawing("http://schemas.openxmlformats.org/drawingml/2006/chart"),
            drawing("http://schemas.openxmlformats.org/drawingml/2006/chart"),
            drawing("http://schemas.openxmlformats.org/drawingml/2006/diagram"),
            drawing("http://schemas.microsoft.com/office/drawing/2017/model3d"),
        );
        let found = scan(&body);
        assert_eq!(found.get("chart"), Some(&2));
        assert_eq!(found.get("smartart"), Some(&1));
        assert_eq!(found.get("graphic:http://schemas.microsoft.com/office/drawing/2017/model3d"), Some(&1));
    }

    #[test]
    fn alt_chunk_and_sub_doc_are_reported() {
        let found = scan(r#"<w:altChunk r:id="x"/><w:p><w:subDoc r:id="y"/></w:p>"#);
        assert_eq!(found.get("altChunk"), Some(&1));
        assert_eq!(found.get("subDoc"), Some(&1));
    }

    #[test]
    fn run_level_vml_fallback_covers_the_choice() {
        let body = r#"<w:p><w:r><mc:AlternateContent><mc:Choice Requires="w14"><w14:contentPart r:id="i"/></mc:Choice><mc:Fallback><w:pict><v:shape><v:imagedata r:id="p"/></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r></w:p>"#;
        assert!(scan(body).is_empty());
    }

    #[test]
    fn drawingml_fallback_does_not_cover_the_choice() {
        let pic = drawing("http://schemas.openxmlformats.org/drawingml/2006/picture");
        let body = format!(
            r#"<w:p><w:r><mc:AlternateContent><mc:Choice Requires="w14"><w14:contentPart r:id="i"/></mc:Choice><mc:Fallback>{pic}</mc:Fallback></mc:AlternateContent></w:r></w:p>"#
        );
        assert_eq!(scan(&body).get("ink"), Some(&1));
    }

    #[test]
    fn paragraph_level_alternate_content_ignores_the_fallback() {
        let chart = drawing("http://schemas.openxmlformats.org/drawingml/2006/chart");
        let body = format!(
            r#"<w:p><mc:AlternateContent><mc:Choice Requires="c">{chart}</mc:Choice><mc:Fallback><w:pict/></mc:Fallback></mc:AlternateContent></w:p>"#
        );
        assert_eq!(scan(&body).get("chart"), Some(&1));
    }

    #[test]
    fn unsupported_content_inside_a_fallback_is_not_a_hole() {
        let chart = drawing("http://schemas.openxmlformats.org/drawingml/2006/chart");
        let shape = drawing("http://schemas.microsoft.com/office/word/2010/wordprocessingShape");
        let body = format!(
            r#"<w:p><w:r><mc:AlternateContent><mc:Choice Requires="wps">{shape}</mc:Choice><mc:Fallback>{chart}</mc:Fallback></mc:AlternateContent></w:r></w:p>"#
        );
        assert!(scan(&body).is_empty());
    }

    #[test]
    fn ole_object_needs_a_preview() {
        let with = r#"<w:p><w:r><w:object><v:shape><v:imagedata r:id="p"/></v:shape><o:OLEObject ProgID="Excel.Sheet.12"/></w:object></w:r></w:p>"#;
        let without = r#"<w:p><w:r><w:object><o:OLEObject ProgID="Package"/></w:object></w:r></w:p>"#;
        assert!(scan(with).is_empty());
        assert_eq!(scan(without).get("ole-without-preview"), Some(&1));
    }

    #[test]
    fn story_parts() {
        for name in ["word/document.xml", "word/header1.xml", "word/footer12.xml", "word/footnotes.xml", "word/endnotes.xml", "word/header.xml"] {
            assert!(is_story_part(name), "{name}");
        }
        for name in ["word/styles.xml", "word/comments.xml", "word/headerx.xml", "word/charts/chart1.xml"] {
            assert!(!is_story_part(name), "{name}");
        }
    }
}
