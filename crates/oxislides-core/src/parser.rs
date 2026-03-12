use quick_xml::events::Event;
use quick_xml::reader::Reader;
use thiserror::Error;

use oxi_common::archive::OoxmlArchive;
use oxi_common::relationships::parse_relationships;
use oxi_common::xml_utils::{emu_to_pt, get_attr, local_name};

use crate::ir::{
    Presentation, Shape, ShapeContent, Slide, SlideAlignment, SlideParagraph, SlideRun,
};

#[derive(Error, Debug)]
pub enum PptxError {
    #[error("Archive error: {0}")]
    Archive(#[from] oxi_common::OxiError),

    #[error("XML error: {0}")]
    Xml(#[from] quick_xml::Error),

    #[error("Invalid data: {0}")]
    InvalidData(String),
}

/// Information about a slide from presentation.xml
struct SlideInfo {
    r_id: String,
}

/// Parse presentation.xml to get slide relationship IDs (in order).
fn parse_presentation_slides(xml: &str) -> Result<(Vec<SlideInfo>, f32, f32), PptxError> {
    let mut reader = Reader::from_str(xml);
    let mut slides = Vec::new();
    // Default slide size: 10 inches x 7.5 inches in EMU
    let mut width_emu: f32 = 9144000.0;
    let mut height_emu: f32 = 6858000.0;

    loop {
        match reader.read_event()? {
            Event::Start(e) | Event::Empty(e) => {
                let name = local_name(e.name().as_ref());
                match name.as_str() {
                    "sldIdLst" => {} // container for sldId entries
                    "sldId" => {
                        // r:id attribute (namespaced, so try raw "r:id" first)
                        let r_id = {
                            let mut found = None;
                            for attr in e.attributes().flatten() {
                                let key =
                                    std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                                if key == "r:id" {
                                    found = Some(
                                        String::from_utf8_lossy(&attr.value).to_string(),
                                    );
                                    break;
                                }
                            }
                            found.unwrap_or_default()
                        };
                        if !r_id.is_empty() {
                            slides.push(SlideInfo { r_id });
                        }
                    }
                    "sldSz" => {
                        if let Some(cx) = get_attr(&e, "cx") {
                            if let Ok(v) = cx.parse::<f32>() {
                                width_emu = v;
                            }
                        }
                        if let Some(cy) = get_attr(&e, "cy") {
                            if let Ok(v) = cy.parse::<f32>() {
                                height_emu = v;
                            }
                        }
                    }
                    _ => {}
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok((slides, emu_to_pt(width_emu), emu_to_pt(height_emu)))
}

/// Parse a single slide XML into shapes.
fn parse_slide(
    xml: &str,
    slide_index: usize,
    archive: &mut OoxmlArchive,
    slide_rels_path: &str,
) -> Result<Slide, PptxError> {
    // Parse slide relationships for image resolution
    let rels = if let Ok(Some(rels_xml)) = archive.try_read_part(slide_rels_path) {
        parse_relationships(&rels_xml).unwrap_or_default()
    } else {
        Default::default()
    };

    let mut reader = Reader::from_str(xml);
    let mut shapes = Vec::new();
    let mut _depth = 0u32;
    let mut in_sp_tree = false;

    // Shape state
    let mut in_shape = false;
    let mut shape_x: f32 = 0.0;
    let mut shape_y: f32 = 0.0;
    let mut shape_w: f32 = 0.0;
    let mut shape_h: f32 = 0.0;
    let mut shape_paragraphs: Vec<SlideParagraph> = Vec::new();
    let mut shape_is_image = false;
    let mut shape_image_r_id: Option<String> = None;

    // Paragraph state
    let mut in_paragraph = false;
    let mut para_runs: Vec<SlideRun> = Vec::new();
    let mut para_alignment = SlideAlignment::default();

    // Run state
    let mut in_run = false;
    let mut run_text = String::new();
    let mut run_bold = false;
    let mut run_italic = false;
    let mut run_font_size: Option<f32> = None;
    let mut run_color: Option<String> = None;
    let mut run_font_family: Option<String> = None;

    let mut in_text = false;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let name = local_name(e.name().as_ref());
                _depth += 1;

                match name.as_str() {
                    "spTree" => {
                        in_sp_tree = true;
                    }
                    "sp" | "pic" if in_sp_tree => {
                        in_shape = true;
                        shape_x = 0.0;
                        shape_y = 0.0;
                        shape_w = 0.0;
                        shape_h = 0.0;
                        shape_paragraphs.clear();
                        shape_is_image = name == "pic";
                        shape_image_r_id = None;
                    }
                    "off" if in_shape => {
                        if let Some(x) = get_attr(&e, "x") {
                            shape_x = x.parse::<f32>().map(emu_to_pt).unwrap_or(0.0);
                        }
                        if let Some(y) = get_attr(&e, "y") {
                            shape_y = y.parse::<f32>().map(emu_to_pt).unwrap_or(0.0);
                        }
                    }
                    "ext" if in_shape => {
                        if let Some(cx) = get_attr(&e, "cx") {
                            shape_w = cx.parse::<f32>().map(emu_to_pt).unwrap_or(0.0);
                        }
                        if let Some(cy) = get_attr(&e, "cy") {
                            shape_h = cy.parse::<f32>().map(emu_to_pt).unwrap_or(0.0);
                        }
                    }
                    "blipFill" if in_shape => {
                        shape_is_image = true;
                    }
                    "blip" if in_shape && shape_is_image => {
                        // r:embed attribute for image reference
                        for attr in e.attributes().flatten() {
                            let key = std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                            if key == "r:embed" || key.ends_with(":embed") {
                                shape_image_r_id =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    "p" if in_shape => {
                        in_paragraph = true;
                        para_runs.clear();
                        para_alignment = SlideAlignment::default();
                    }
                    "pPr" if in_paragraph => {
                        if let Some(algn) = get_attr(&e, "algn") {
                            para_alignment = match algn.as_str() {
                                "ctr" => SlideAlignment::Center,
                                "r" => SlideAlignment::Right,
                                "just" => SlideAlignment::Justify,
                                _ => SlideAlignment::Left,
                            };
                        }
                    }
                    "r" if in_paragraph => {
                        in_run = true;
                        run_text.clear();
                        run_bold = false;
                        run_italic = false;
                        run_font_size = None;
                        run_color = None;
                        run_font_family = None;
                    }
                    "rPr" if in_run => {
                        if let Some(b) = get_attr(&e, "b") {
                            run_bold = b == "1" || b == "true";
                        }
                        if let Some(i) = get_attr(&e, "i") {
                            run_italic = i == "1" || i == "true";
                        }
                        if let Some(sz) = get_attr(&e, "sz") {
                            // Font size in hundredths of a point
                            if let Ok(v) = sz.parse::<f32>() {
                                run_font_size = Some(v / 100.0);
                            }
                        }
                    }
                    "solidFill" => {} // container
                    "srgbClr" if in_run => {
                        if let Some(val) = get_attr(&e, "val") {
                            run_color = Some(val);
                        }
                    }
                    "latin" | "ea" | "cs" if in_run => {
                        if let Some(typeface) = get_attr(&e, "typeface") {
                            if run_font_family.is_none() {
                                run_font_family = Some(typeface);
                            }
                        }
                    }
                    "t" if in_run => {
                        in_text = true;
                    }
                    _ => {}
                }
            }
            Event::Empty(e) => {
                let name = local_name(e.name().as_ref());
                match name.as_str() {
                    "off" if in_shape => {
                        if let Some(x) = get_attr(&e, "x") {
                            shape_x = x.parse::<f32>().map(emu_to_pt).unwrap_or(0.0);
                        }
                        if let Some(y) = get_attr(&e, "y") {
                            shape_y = y.parse::<f32>().map(emu_to_pt).unwrap_or(0.0);
                        }
                    }
                    "ext" if in_shape => {
                        if let Some(cx) = get_attr(&e, "cx") {
                            shape_w = cx.parse::<f32>().map(emu_to_pt).unwrap_or(0.0);
                        }
                        if let Some(cy) = get_attr(&e, "cy") {
                            shape_h = cy.parse::<f32>().map(emu_to_pt).unwrap_or(0.0);
                        }
                    }
                    "blip" if in_shape && shape_is_image => {
                        for attr in e.attributes().flatten() {
                            let key = std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                            if key == "r:embed" || key.ends_with(":embed") {
                                shape_image_r_id =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    "rPr" if in_run => {
                        if let Some(b) = get_attr(&e, "b") {
                            run_bold = b == "1" || b == "true";
                        }
                        if let Some(i) = get_attr(&e, "i") {
                            run_italic = i == "1" || i == "true";
                        }
                        if let Some(sz) = get_attr(&e, "sz") {
                            if let Ok(v) = sz.parse::<f32>() {
                                run_font_size = Some(v / 100.0);
                            }
                        }
                    }
                    "pPr" if in_paragraph => {
                        if let Some(algn) = get_attr(&e, "algn") {
                            para_alignment = match algn.as_str() {
                                "ctr" => SlideAlignment::Center,
                                "r" => SlideAlignment::Right,
                                "just" => SlideAlignment::Justify,
                                _ => SlideAlignment::Left,
                            };
                        }
                    }
                    "srgbClr" if in_run => {
                        if let Some(val) = get_attr(&e, "val") {
                            run_color = Some(val);
                        }
                    }
                    "latin" | "ea" | "cs" if in_run => {
                        if let Some(typeface) = get_attr(&e, "typeface") {
                            if run_font_family.is_none() {
                                run_font_family = Some(typeface);
                            }
                        }
                    }
                    _ => {}
                }
            }
            Event::End(e) => {
                let name = local_name(e.name().as_ref());
                _depth -= 1;

                match name.as_str() {
                    "spTree" => {
                        in_sp_tree = false;
                    }
                    "sp" | "pic" if in_shape => {
                        let content = if shape_is_image {
                            if let Some(ref r_id) = shape_image_r_id {
                                if let Some(rel) = rels.get(r_id) {
                                    // Load image data from archive
                                    let image_path =
                                        resolve_slide_relative_path(slide_rels_path, &rel.target);
                                    let data = archive
                                        .read_binary_part(&image_path)
                                        .unwrap_or_default();
                                    let ct = detect_content_type(&rel.target);
                                    ShapeContent::Image {
                                        data,
                                        content_type: ct,
                                    }
                                } else {
                                    ShapeContent::Placeholder
                                }
                            } else {
                                ShapeContent::Placeholder
                            }
                        } else if !shape_paragraphs.is_empty() {
                            ShapeContent::TextBox {
                                paragraphs: std::mem::take(&mut shape_paragraphs),
                            }
                        } else {
                            ShapeContent::Placeholder
                        };

                        shapes.push(Shape {
                            x: shape_x,
                            y: shape_y,
                            width: shape_w,
                            height: shape_h,
                            content,
                        });
                        in_shape = false;
                    }
                    "p" if in_paragraph => {
                        in_paragraph = false;
                        shape_paragraphs.push(SlideParagraph {
                            runs: std::mem::take(&mut para_runs),
                            alignment: para_alignment,
                        });
                    }
                    "r" if in_run => {
                        in_run = false;
                        if !run_text.is_empty() {
                            para_runs.push(SlideRun {
                                text: std::mem::take(&mut run_text),
                                font_size: run_font_size,
                                bold: run_bold,
                                italic: run_italic,
                                color: run_color.take(),
                                font_family: run_font_family.take(),
                            });
                        }
                    }
                    "t" => {
                        in_text = false;
                    }
                    _ => {}
                }
            }
            Event::Text(e) => {
                if in_text && in_run {
                    let text = e.unescape()?.to_string();
                    run_text.push_str(&text);
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(Slide {
        index: slide_index,
        shapes,
    })
}

/// Resolve a target path relative to the slide location.
/// e.g., slide rels at "ppt/slides/_rels/slide1.xml.rels", target "../media/image1.png"
///   -> "ppt/media/image1.png"
fn resolve_slide_relative_path(rels_path: &str, target: &str) -> String {
    if target.starts_with('/') {
        return target.trim_start_matches('/').to_string();
    }

    // Get the directory of the slide (parent of _rels)
    // rels_path: "ppt/slides/_rels/slide1.xml.rels"
    // slide dir: "ppt/slides/"
    let slide_dir = if let Some(pos) = rels_path.rfind("/_rels/") {
        &rels_path[..pos + 1] // "ppt/slides/"
    } else if let Some(pos) = rels_path.rfind('/') {
        &rels_path[..pos + 1]
    } else {
        ""
    };

    // Resolve "../" segments
    let mut base_parts: Vec<&str> = slide_dir
        .split('/')
        .filter(|s| !s.is_empty())
        .collect();
    for segment in target.split('/') {
        match segment {
            ".." => {
                base_parts.pop();
            }
            "." | "" => {}
            s => base_parts.push(s),
        }
    }

    base_parts.join("/")
}

/// Detect content type from file extension.
fn detect_content_type(path: &str) -> Option<String> {
    let lower = path.to_lowercase();
    if lower.ends_with(".png") {
        Some("image/png".to_string())
    } else if lower.ends_with(".jpg") || lower.ends_with(".jpeg") {
        Some("image/jpeg".to_string())
    } else if lower.ends_with(".gif") {
        Some("image/gif".to_string())
    } else if lower.ends_with(".bmp") {
        Some("image/bmp".to_string())
    } else if lower.ends_with(".svg") {
        Some("image/svg+xml".to_string())
    } else if lower.ends_with(".emf") {
        Some("image/x-emf".to_string())
    } else if lower.ends_with(".wmf") {
        Some("image/x-wmf".to_string())
    } else if lower.ends_with(".tiff") || lower.ends_with(".tif") {
        Some("image/tiff".to_string())
    } else {
        None
    }
}

/// Parse a .pptx file from raw bytes into a Presentation IR.
pub fn parse_pptx(data: &[u8]) -> Result<Presentation, PptxError> {
    let mut archive = OoxmlArchive::new(data)?;

    // 1. Parse presentation.xml for slide list and slide size
    let pres_xml = archive.read_part("ppt/presentation.xml")?;
    let (slide_infos, slide_width, slide_height) = parse_presentation_slides(&pres_xml)?;

    // 2. Parse presentation relationships
    let rels_xml = archive.read_part("ppt/_rels/presentation.xml.rels")?;
    let rels = parse_relationships(&rels_xml)?;

    // Build rId -> target path map
    let rid_to_path: std::collections::HashMap<String, String> = rels
        .into_iter()
        .map(|(id, rel)| (id, rel.target))
        .collect();

    // 3. Parse each slide
    let mut slides = Vec::new();
    for (i, info) in slide_infos.iter().enumerate() {
        let slide_target = match rid_to_path.get(&info.r_id) {
            Some(target) => {
                if target.starts_with('/') {
                    target.trim_start_matches('/').to_string()
                } else {
                    format!("ppt/{}", target)
                }
            }
            None => {
                log::warn!("No relationship found for slide rId={}, skipping", info.r_id);
                continue;
            }
        };

        // Slide rels path: e.g., "ppt/slides/_rels/slide1.xml.rels"
        let slide_rels_path = {
            if let Some(pos) = slide_target.rfind('/') {
                let dir = &slide_target[..pos + 1];
                let filename = &slide_target[pos + 1..];
                format!("{}/_rels/{}.rels", dir.trim_end_matches('/'), filename)
            } else {
                format!("_rels/{}.rels", slide_target)
            }
        };

        match archive.try_read_part(&slide_target)? {
            Some(slide_xml) => {
                let slide =
                    parse_slide(&slide_xml, i + 1, &mut archive, &slide_rels_path)?;
                slides.push(slide);
            }
            None => {
                log::warn!("Slide file '{}' not found in archive, skipping", slide_target);
            }
        }
    }

    Ok(Presentation {
        slides,
        slide_width,
        slide_height,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_presentation_slides() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId2"/>
    <p:sldId id="257" r:id="rId3"/>
  </p:sldIdLst>
  <p:sldSz cx="9144000" cy="6858000"/>
</p:presentation>"#;
        let (slides, w, h) = parse_presentation_slides(xml).unwrap();
        assert_eq!(slides.len(), 2);
        assert_eq!(slides[0].r_id, "rId2");
        assert_eq!(slides[1].r_id, "rId3");
        // 9144000 EMU / 12700 = 720pt
        assert!((w - 720.0).abs() < 0.1);
        // 6858000 EMU / 12700 = 540pt
        assert!((h - 540.0).abs() < 0.1);
    }

    #[test]
    fn test_resolve_slide_relative_path() {
        assert_eq!(
            resolve_slide_relative_path("ppt/slides/_rels/slide1.xml.rels", "../media/image1.png"),
            "ppt/media/image1.png"
        );
        assert_eq!(
            resolve_slide_relative_path("ppt/slides/_rels/slide1.xml.rels", "image1.png"),
            "ppt/slides/image1.png"
        );
        assert_eq!(
            resolve_slide_relative_path("ppt/slides/_rels/slide1.xml.rels", "/ppt/media/img.png"),
            "ppt/media/img.png"
        );
    }

    #[test]
    fn test_detect_content_type() {
        assert_eq!(detect_content_type("image1.png"), Some("image/png".to_string()));
        assert_eq!(detect_content_type("photo.JPEG"), Some("image/jpeg".to_string()));
        assert_eq!(detect_content_type("logo.svg"), Some("image/svg+xml".to_string()));
        assert_eq!(detect_content_type("unknown.xyz"), None);
    }
}
