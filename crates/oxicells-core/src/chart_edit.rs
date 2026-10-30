// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Preserve chart parts and patch the properties a macro actually changed.

use crate::{ir::Chart, parser::XlsxError};
use oxidocs_common::{
    archive::OoxmlArchive,
    xml_utils::{get_attr, local_name},
};
use quick_xml::{events::Event, Reader};

const C: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";

pub(crate) fn chart_part_names(archive: &mut OoxmlArchive) -> Result<Vec<String>, XlsxError> {
    let xml = archive.read_part("[Content_Types].xml")?;
    let mut reader = Reader::from_str(&xml);
    let mut parts = Vec::new();
    loop {
        match reader.read_event().map_err(XlsxError::Xml)? {
            Event::Start(element) | Event::Empty(element)
                if local_name(element.name().as_ref()) == "Override" =>
            {
                if get_attr(&element, "ContentType").as_deref()
                    == Some("application/vnd.openxmlformats-officedocument.drawingml.chart+xml")
                {
                    if let Some(part) = get_attr(&element, "PartName") {
                        parts.push(part.trim_start_matches('/').to_string());
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    parts.sort();
    parts.dedup();
    Ok(parts)
}

pub(crate) fn legend_position(chart: &Chart) -> Option<&str> {
    chart.legend.as_ref().map(|legend| legend.position.as_str())
}

fn escape(text: &str) -> String {
    quick_xml::escape::escape(text).into_owned()
}

/// Ranges of direct children of the chart element, excluding axes/extensions.
fn children_at(
    xml: &str,
    path: &[&str],
) -> Result<(Vec<(String, usize, usize)>, usize), XlsxError> {
    use quick_xml::{name::ResolveResult, reader::NsReader};
    let mut reader = NsReader::from_str(xml);
    let mut parents = Vec::new();
    let mut start = None;
    let mut ranges = Vec::new();
    let mut end = None;
    let matches_path = |parents: &[(String, bool)]| {
        parents.len() == path.len()
            && parents
                .iter()
                .zip(path)
                .all(|((name, native), expected)| *native && name == expected)
    };
    loop {
        let before = reader.buffer_position() as usize;
        let (namespace, event) = reader.read_resolved_event().map_err(XlsxError::Xml)?;
        let native = matches!(namespace,ResolveResult::Bound(ns) if matches!(ns.as_ref(),b"http://schemas.openxmlformats.org/drawingml/2006/chart"|b"http://purl.oclc.org/ooxml/drawingml/chart"));
        let after = reader.buffer_position() as usize;
        match event {
            Event::Start(element) => {
                let name = local_name(element.name().as_ref());
                if native && matches_path(&parents) {
                    start = Some((name.clone(), before));
                }
                parents.push((name, native));
            }
            Event::Empty(element) => {
                if native && matches_path(&parents) {
                    ranges.push((local_name(element.name().as_ref()), before, after));
                }
            }
            Event::End(_) => {
                if matches_path(&parents) {
                    end = Some(before);
                }
                if parents.len() == path.len() + 1 && matches_path(&parents[..path.len()]) {
                    if let Some((name, begin)) = start.take() {
                        ranges.push((name, begin, after));
                    }
                }
                parents.pop();
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok((
        ranges,
        end.ok_or_else(|| XlsxError::InvalidData("chart element is missing".into()))?,
    ))
}

fn replace_child(xml: &str, name: &str, replacement: &str) -> Result<String, XlsxError> {
    let (ranges, end) = children_at(xml, &["chartSpace", "chart"])?;
    if let Some((_, start, finish)) = ranges.iter().find(|(found, _, _)| found == name) {
        return Ok(format!(
            "{}{}{}",
            &xml[..*start],
            replacement,
            &xml[*finish..]
        ));
    }
    if replacement.is_empty() {
        return Ok(xml.to_string());
    }
    let before = match name {
        "title" => ranges.first().map(|(_, start, _)| *start),
        "autoTitleDeleted" => ranges
            .iter()
            .find(|(found, _, _)| found != "title")
            .map(|(_, start, _)| *start),
        "legend" => ranges
            .iter()
            .find(|(found, _, _)| {
                matches!(
                    found.as_str(),
                    "plotVisOnly" | "dispBlanksAs" | "showDLblsOverMax" | "extLst"
                )
            })
            .map(|(_, start, _)| *start),
        _ => None,
    }
    .unwrap_or(end);
    Ok(format!(
        "{}{}{}",
        &xml[..before],
        replacement,
        &xml[before..]
    ))
}

pub(crate) fn patch_chart(xml: &str, before: &Chart, after: &Chart) -> Result<String, XlsxError> {
    let mut reader = quick_xml::reader::NsReader::from_str(xml);
    let strict = loop {
        let (namespace, event) = reader.read_resolved_event().map_err(XlsxError::Xml)?;
        match event {
            Event::Start(element) | Event::Empty(element)
                if local_name(element.name().as_ref()) == "chartSpace" =>
            {
                break matches!(namespace,quick_xml::name::ResolveResult::Bound(ns) if ns.as_ref()==b"http://purl.oclc.org/ooxml/drawingml/chart");
            }
            Event::Eof => return Err(XlsxError::InvalidData("chart root is missing".into())),
            _ => {}
        }
    };
    let chart_ns = if strict {
        "http://purl.oclc.org/ooxml/drawingml/chart"
    } else {
        C
    };
    let drawing_ns = if strict {
        "http://purl.oclc.org/ooxml/drawingml/main"
    } else {
        A
    };
    let mut output = xml.to_string();
    if before.title != after.title {
        let title=after.title.as_ref().map(|text|format!(
            "<c:title xmlns:c=\"{chart_ns}\" xmlns:a=\"{drawing_ns}\"><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t xml:space=\"preserve\">{}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val=\"0\"/></c:title>",escape(text))).unwrap_or_default();
        if let Some(text) = &after.title {
            let tx=format!("<c:tx xmlns:c=\"{chart_ns}\" xmlns:a=\"{drawing_ns}\"><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t xml:space=\"preserve\">{}</a:t></a:r></a:p></c:rich></c:tx>",escape(text));
            if let Ok((children, _)) = children_at(&output, &["chartSpace", "chart", "title"]) {
                if let Some((_, start, end)) = children.iter().find(|(name, _, _)| name == "tx") {
                    output = format!("{}{}{}", &output[..*start], tx, &output[*end..]);
                } else {
                    output = replace_child(&output, "title", &title)?;
                }
            } else {
                output = replace_child(&output, "title", &title)?;
            }
        } else {
            output = replace_child(&output, "title", "")?;
        }
        // Excel uses this flag when a title is absent; keep it consistent.
        output = replace_child(
            &output,
            "autoTitleDeleted",
            &format!(
                "<c:autoTitleDeleted xmlns:c=\"{chart_ns}\" val=\"{}\"/>",
                if after.title.is_some() { 0 } else { 1 }
            ),
        )?;
    }
    if legend_position(before) != legend_position(after) {
        let legend=legend_position(after).map(|position|format!("<c:legend xmlns:c=\"{chart_ns}\"><c:legendPos val=\"{}\"/><c:overlay val=\"0\"/></c:legend>",escape(position))).unwrap_or_default();
        if let Some(position) = legend_position(after) {
            if let Ok((children, _)) = children_at(&output, &["chartSpace", "chart", "legend"]) {
                if let Some((_, start, end)) =
                    children.iter().find(|(name, _, _)| name == "legendPos")
                {
                    output = format!(
                        "{}<c:legendPos xmlns:c=\"{chart_ns}\" val=\"{}\"/>{}",
                        &output[..*start],
                        escape(position),
                        &output[*end..]
                    );
                } else {
                    output = replace_child(&output, "legend", &legend)?;
                }
            } else {
                output = replace_child(&output, "legend", &legend)?;
            }
        } else {
            output = replace_child(&output, "legend", "")?;
        }
    }
    Ok(output)
}
