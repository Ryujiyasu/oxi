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

/// Writes a newly created chart from portable chart data. Existing chart parts
/// continue through the differential patcher so their unmodelled XML survives.
pub(crate) fn new_chart_xml(chart: &Chart) -> Result<String, XlsxError> {
    use crate::ir::{ChartAxis, ChartSeries, Frame, Marker, ShapeLine};
    fn paint(line: &ShapeLine) -> String {
        let mut out = format!("<a:ln w=\"{}\">", line.width.max(0));
        let color = crate::editor::write_drawing_color(&line.color, line.color_theme,
            line.legacy_color_index, line.color_brightness, line.color_shade);
        if let Some(pattern) = &line.pattern {
            let back = crate::editor::write_drawing_color(line.back_color.as_deref().unwrap_or("FFFFFF"),
                line.back_color_theme, line.back_legacy_color_index, line.back_color_brightness, line.back_color_shade);
            out.push_str(&format!("<a:pattFill prst=\"{}\"><a:fgClr>{color}</a:fgClr><a:bgClr>{back}</a:bgClr></a:pattFill>", escape(pattern)));
        } else { out.push_str(&format!("<a:solidFill>{color}</a:solidFill>")); }
        if let Some(dash) = &line.dash { out.push_str(&format!("<a:prstDash val=\"{}\"/>", escape(dash))); }
        for (tag, kind, width, length) in [
            ("headEnd", &line.head_end, &line.head_end_width, &line.head_end_length),
            ("tailEnd", &line.tail_end, &line.tail_end_width, &line.tail_end_length),
        ] {
            if let Some(kind) = kind {
                out.push_str(&format!("<a:{tag} type=\"{}\"", escape(kind)));
                if let Some(width) = width { out.push_str(&format!(" w=\"{}\"", escape(width))); }
                if let Some(length) = length { out.push_str(&format!(" len=\"{}\"", escape(length))); }
                out.push_str("/>");
            }
        }
        out.push_str("</a:ln>");
        if let Some(cap) = &line.cap { out = out.replacen("<a:ln ", &format!("<a:ln cap=\"{}\" ", escape(cap)), 1); }
        out
    }
    fn marker(value: &Marker) -> String {
        let mut out = format!("<c:marker><c:symbol val=\"{}\"/><c:size val=\"{}\"/>", escape(&value.symbol), value.size);
        if value.fill.is_some() || value.line.is_some() {
            out.push_str("<c:spPr>");
            if let Some(color) = &value.fill { out.push_str(&format!("<a:solidFill><a:srgbClr val=\"{}\"/></a:solidFill>", escape(color))); }
            if let Some(color) = &value.line { out.push_str(&format!("<a:ln><a:solidFill><a:srgbClr val=\"{}\"/></a:solidFill></a:ln>", escape(color))); }
            out.push_str("</c:spPr>");
        }
        out.push_str("</c:marker>"); out
    }
    fn rich(text: &str) -> String {
        format!("<c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t xml:space=\"preserve\">{}</a:t></a:r></a:p></c:rich></c:tx>", escape(text))
    }
    fn text_properties(size: f32, face: Option<&str>) -> String {
        let mut font = String::new();
        if size.is_finite() && size > 0.0 { font.push_str(&format!(" sz=\"{}\"", (size * 100.0).round() as u32)); }
        let face = face.map(|face| format!("<a:latin typeface=\"{}\"/>", escape(face))).unwrap_or_default();
        format!("<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr{font}>{face}</a:defRPr></a:pPr><a:endParaRPr lang=\"en-US\"/></a:p></c:txPr>")
    }
    fn layout(frame: Option<&Frame>, inner: bool) -> String {
        match frame {
            None => "<c:layout/>".into(),
            Some(f) => format!("<c:layout><c:manualLayout>{}<c:xMode val=\"factor\"/><c:yMode val=\"factor\"/><c:wMode val=\"factor\"/><c:hMode val=\"factor\"/><c:x val=\"{}\"/><c:y val=\"{}\"/><c:w val=\"{}\"/><c:h val=\"{}\"/></c:manualLayout></c:layout>",
                if inner { "<c:layoutTarget val=\"inner\"/>" } else { "" }, f.x, f.y, f.w, f.h),
        }
    }
    fn numeric(values: &[Option<f64>]) -> String {
        let mut out = format!("<c:numLit><c:formatCode>General</c:formatCode><c:ptCount val=\"{}\"/>", values.len());
        for (i, value) in values.iter().enumerate() {
            if let Some(value) = value { out.push_str(&format!("<c:pt idx=\"{i}\"><c:v>{value}</c:v></c:pt>")); }
        }
        out.push_str("</c:numLit>"); out
    }
    fn series(series: &ChartSeries, index: usize, categories: &[String], scatter: bool) -> String {
        let mut out = format!("<c:ser><c:idx val=\"{index}\"/><c:order val=\"{index}\"/><c:tx><c:v>{}</c:v></c:tx>", escape(&series.name));
        if let Some(line) = &series.line { out.push_str(&format!("<c:spPr>{}</c:spPr>", paint(line))); }
        if let Some(value) = &series.marker { out.push_str(&marker(value)); }
        for point in &series.points {
            out.push_str(&format!("<c:dPt><c:idx val=\"{}\"/>", point.index));
            if let Some(value) = &point.marker { out.push_str(&marker(value)); }
            if let Some(line) = &point.line { out.push_str(&format!("<c:spPr>{}</c:spPr>", paint(line))); }
            out.push_str("</c:dPt>");
        }
        if !series.labels.is_empty() || series.label_pos.is_some() {
            out.push_str("<c:dLbls>");
            for label in &series.labels {
                out.push_str(&format!("<c:dLbl><c:idx val=\"{}\"/>", label.index));
                if let Some((x, y)) = label.offset {
                    out.push_str(&format!("<c:layout><c:manualLayout><c:xMode val=\"factor\"/><c:yMode val=\"factor\"/><c:x val=\"{x}\"/><c:y val=\"{y}\"/></c:manualLayout></c:layout>"));
                }
                if let Some(text) = &label.text { out.push_str(&rich(text)); }
                if let Some(code) = &label.number_format { out.push_str(&format!("<c:numFmt formatCode=\"{}\" sourceLinked=\"0\"/>", escape(code))); }
                out.push_str(&text_properties(label.size, label.face.as_deref()));
                if let Some(position) = &label.position { out.push_str(&format!("<c:dLblPos val=\"{}\"/>", escape(position))); }
                out.push_str("<c:showVal val=\"1\"/></c:dLbl>");
            }
            out.push_str(&text_properties(series.label_size, series.label_face.as_deref()));
            if let Some(position) = &series.label_pos { out.push_str(&format!("<c:dLblPos val=\"{}\"/>", escape(position))); }
            out.push_str("<c:showVal val=\"0\"/></c:dLbls>");
        }
        if scatter {
            let xs: Vec<_> = (0..series.values.len()).map(|i| Some(categories.get(i).and_then(|v| v.parse::<f64>().ok()).unwrap_or((i+1) as f64))).collect();
            out.push_str(&format!("<c:xVal>{}</c:xVal><c:yVal>{}</c:yVal>", numeric(&xs), numeric(&series.values)));
        } else {
            out.push_str(&format!("<c:cat><c:strLit><c:ptCount val=\"{}\"/>", categories.len()));
            for (i, value) in categories.iter().enumerate() { out.push_str(&format!("<c:pt idx=\"{i}\"><c:v>{}</c:v></c:pt>", escape(value))); }
            out.push_str(&format!("</c:strLit></c:cat><c:val>{}</c:val>", numeric(&series.values)));
        }
        out.push_str("</c:ser>"); out
    }
    fn axis(axis: Option<&ChartAxis>, tag: &str, id: u32, cross: u32, position: &str) -> String {
        let fallback = ChartAxis::default(); let axis = axis.unwrap_or(&fallback);
        let mut out = format!("<c:{tag}><c:axId val=\"{id}\"/><c:scaling><c:orientation val=\"minMax\"/>");
        if let Some(max) = axis.max { out.push_str(&format!("<c:max val=\"{max}\"/>")); }
        if let Some(min) = axis.min { out.push_str(&format!("<c:min val=\"{min}\"/>")); }
        out.push_str(&format!("</c:scaling><c:delete val=\"{}\"/><c:axPos val=\"{}\"/>", u8::from(axis.deleted), escape(if axis.position.is_empty() { position } else { &axis.position })));
        if let Some(line) = &axis.major_gridline { out.push_str(&format!("<c:majorGridlines><c:spPr>{}</c:spPr></c:majorGridlines>", paint(line))); }
        if let Some(code) = &axis.number_format { out.push_str(&format!("<c:numFmt formatCode=\"{}\" sourceLinked=\"0\"/>", escape(code))); }
        out.push_str(&format!("<c:majorTickMark val=\"{}\"/><c:tickLblPos val=\"{}\"/>", escape(if axis.major_tick.is_empty() { "none" } else { &axis.major_tick }), escape(if axis.tick_labels.is_empty() { "nextTo" } else { &axis.tick_labels })));
        if let Some(line) = &axis.line { out.push_str(&format!("<c:spPr>{}</c:spPr>", paint(line))); }
        out.push_str(&text_properties(axis.size, axis.face.as_deref()));
        out.push_str(&format!("<c:crossAx val=\"{cross}\"/><c:crosses val=\"autoZero\"/>"));
        if tag == "valAx" {
            if let Some(between) = &axis.cross_between { out.push_str(&format!("<c:crossBetween val=\"{}\"/>", escape(between))); }
            if let Some(unit) = axis.major_unit { out.push_str(&format!("<c:majorUnit val=\"{unit}\"/>")); }
        }
        out.push_str(&format!("</c:{tag}>")); out
    }
    let (plot, settings, axes) = match chart.kind.as_str() {
        "bar" => ("barChart", "<c:barDir val=\"col\"/><c:grouping val=\"clustered\"/><c:varyColors val=\"0\"/>", true),
        "line" => ("lineChart", "<c:grouping val=\"standard\"/><c:varyColors val=\"0\"/>", true),
        "area" => ("areaChart", "<c:grouping val=\"standard\"/><c:varyColors val=\"0\"/>", true),
        "scatter" => ("scatterChart", "<c:scatterStyle val=\"marker\"/><c:varyColors val=\"0\"/>", true),
        "pie" => ("pieChart", "<c:varyColors val=\"1\"/>", false),
        "doughnut" => ("doughnutChart", "<c:varyColors val=\"1\"/>", false),
        "radar" => ("radarChart", "<c:radarStyle val=\"standard\"/><c:varyColors val=\"0\"/>", true),
        kind => return Err(XlsxError::InvalidData(format!("new chart plot kind {kind:?} cannot be serialized"))),
    };
    if !chart.shapes.is_empty() {
        return Err(XlsxError::InvalidData("new chart annotations need their own drawing part".into()));
    }
    let finite_frame = |frame: Option<&Frame>| frame.is_none_or(|f| [f.x, f.y, f.w, f.h].iter().all(|v| v.is_finite()));
    if !finite_frame(chart.plot.as_ref()) || !finite_frame(chart.legend.as_ref().and_then(|l| l.frame.as_ref()))
        || chart.series.iter().any(|s| s.values.iter().flatten().any(|v| !v.is_finite())
            || s.labels.iter().filter_map(|l| l.offset).any(|(x,y)| !x.is_finite() || !y.is_finite()))
        || [&chart.value_axis, &chart.category_axis].into_iter().flatten()
            .any(|a| [a.min, a.max, a.major_unit].into_iter().flatten().any(|v| !v.is_finite())) {
        return Err(XlsxError::InvalidData("new chart contains a non-finite value".into()));
    }
    let mut out = format!("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><c:chartSpace xmlns:c=\"{C}\" xmlns:a=\"{A}\"><c:chart>");
    if let Some(title) = &chart.title { out.push_str(&format!("<c:title>{}<c:overlay val=\"0\"/></c:title>", rich(title))); }
    out.push_str(&format!("<c:autoTitleDeleted val=\"{}\"/><c:plotArea>{}<c:{plot}>{settings}", u8::from(chart.title.is_none()), layout(chart.plot.as_ref(), true)));
    for (index, value) in chart.series.iter().enumerate() { out.push_str(&series(value, index, &chart.categories, chart.kind == "scatter")); }
    if chart.kind == "doughnut" { out.push_str("<c:firstSliceAng val=\"0\"/><c:holeSize val=\"50\"/>"); }
    if axes { out.push_str("<c:axId val=\"1\"/><c:axId val=\"2\"/>"); }
    out.push_str(&format!("</c:{plot}>"));
    if axes {
        out.push_str(&axis(chart.category_axis.as_ref(), if chart.kind == "scatter" { "valAx" } else { "catAx" }, 1, 2, "b"));
        out.push_str(&axis(chart.value_axis.as_ref(), "valAx", 2, 1, "l"));
    }
    if let Some(color) = &chart.plot_fill { out.push_str(&format!("<c:spPr><a:solidFill><a:srgbClr val=\"{}\"/></a:solidFill></c:spPr>", escape(color))); }
    out.push_str("</c:plotArea>");
    if let Some(legend) = &chart.legend {
        out.push_str(&format!("<c:legend><c:legendPos val=\"{}\"/>{}<c:overlay val=\"0\"/>{}</c:legend>", escape(&legend.position), layout(legend.frame.as_ref(), false), text_properties(legend.size, legend.face.as_deref())));
    }
    out.push_str("<c:plotVisOnly val=\"1\"/><c:dispBlanksAs val=\"gap\"/></c:chart>");
    if let Some(color) = &chart.fill { out.push_str(&format!("<c:spPr><a:solidFill><a:srgbClr val=\"{}\"/></a:solidFill></c:spPr>", escape(color))); }
    out.push_str("</c:chartSpace>"); Ok(out)
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

#[cfg(test)]
mod new_chart_roundtrip_tests {
    #[test]
    fn new_plot_parts_keep_titles_legends_categories_and_gaps() {
        use crate::ir::{Chart, ChartSeries, Legend};
        for kind in ["bar", "line", "area", "pie", "doughnut", "scatter", "radar"] {
            let expected = Chart {
                kind: kind.into(), title: Some("Revenue & <total>\n2026".into()),
                legend: Some(Legend { position: "l".into(), size: 10.0, ..Default::default() }),
                categories: vec!["1".into(), "2".into(), "3".into()],
                series: vec![ChartSeries { name: "A & B".into(), values: vec![Some(1.25), None, Some(-3.5)], ..Default::default() }],
                ..Default::default()
            };
            let xml = super::new_chart_xml(&expected).unwrap();
            let actual = crate::chart::parse_chart_xml(&xml, &crate::parser::Theme::default()).unwrap();
            assert_eq!(actual.kind, expected.kind, "{kind}");
            assert_eq!(actual.title, expected.title, "{kind}");
            assert_eq!(super::legend_position(&actual), Some("l"), "{kind}");
            assert_eq!(actual.series.len(), 1, "{kind}");
            assert_eq!(actual.series[0].name, expected.series[0].name, "{kind}");
            assert_eq!(actual.series[0].values, expected.series[0].values, "{kind}");
            assert_eq!(actual.categories, expected.categories, "{kind}");
        }
    }
}
