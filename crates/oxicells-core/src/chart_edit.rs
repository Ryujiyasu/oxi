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


fn series_data_equal(a: &crate::ir::ChartSeries, b: &crate::ir::ChartSeries) -> bool {
    a.name == b.name && a.values == b.values
        && a.name_reference == b.name_reference && a.values_reference == b.values_reference
        && a.category_reference == b.category_reference && a.category_values == b.category_values
}

pub(crate) fn series_data_changed(a: &Chart, b: &Chart) -> bool {
    a.series.len() != b.series.len()
        || a.series.iter().zip(&b.series).any(|(a, b)| !series_data_equal(a, b))
        || (a.categories != b.categories && a.series.iter().zip(&b.series)
            .any(|(a, b)| a.category_values.is_none() || b.category_values.is_none()))
}

/// References and their caches are kept together; literal text is not inferred
/// to be numeric merely because it can be parsed as a number.
fn series_data_parts(series: &crate::ir::ChartSeries, shared: &[String], scatter: bool, ns: &str) -> (String, String, String) {
    use crate::ir::ChartCategoryValue;
    fn numeric(values: &[Option<f64>], reference: Option<&str>) -> String {
        let cache = if reference.is_some() { "numCache" } else { "numLit" };
        let mut out = String::new();
        if let Some(reference) = reference {
            out.push_str(&format!("<c:numRef><c:f>{}</c:f>", escape(reference.trim_start_matches('='))));
        }
        out.push_str(&format!("<c:{cache}><c:formatCode>General</c:formatCode><c:ptCount val=\"{}\"/>", values.len()));
        for (i, value) in values.iter().enumerate() {
            if let Some(value) = value { out.push_str(&format!("<c:pt idx=\"{i}\"><c:v>{value}</c:v></c:pt>")); }
        }
        out.push_str(&format!("</c:{cache}>"));
        if reference.is_some() { out.push_str("</c:numRef>"); }
        out
    }
    let name = if let Some(reference) = &series.name_reference {
        format!("<c:strRef><c:f>{}</c:f><c:strCache><c:ptCount val=\"1\"/><c:pt idx=\"0\"><c:v>{}</c:v></c:pt></c:strCache></c:strRef>",
            escape(reference.trim_start_matches('=')), escape(&series.name))
    } else { format!("<c:v>{}</c:v>", escape(&series.name)) };
    let name = format!("<c:tx xmlns:c=\"{ns}\">{name}</c:tx>");
    let legacy: Vec<_> = if scatter {
        (0..series.values.len()).map(|i| ChartCategoryValue::Number(
            shared.get(i).and_then(|v| v.parse::<f64>().ok()).unwrap_or((i + 1) as f64))).collect()
    } else { shared.iter().map(|v| ChartCategoryValue::Text(v.clone())).collect() };
    let categories = series.category_values.as_deref().unwrap_or(&legacy);
    let cats = if categories.is_empty() && series.category_reference.is_none() { String::new() } else {
        let is_numeric = categories.iter().all(|v| matches!(v, ChartCategoryValue::Number(_) | ChartCategoryValue::Blank));
        let body = if is_numeric {
            let values: Vec<_> = categories.iter().map(|v| match v {
                ChartCategoryValue::Number(n) => Some(*n), _ => None,
            }).collect();
            numeric(&values, series.category_reference.as_deref())
        } else {
            let cache = if series.category_reference.is_some() { "strCache" } else { "strLit" };
            let mut out = String::new();
            if let Some(reference) = &series.category_reference {
                out.push_str(&format!("<c:strRef><c:f>{}</c:f>", escape(reference.trim_start_matches('='))));
            }
            out.push_str(&format!("<c:{cache}><c:ptCount val=\"{}\"/>", categories.len()));
            for (i, value) in categories.iter().enumerate() {
                let text = match value {
                    ChartCategoryValue::Text(s) => s.clone(),
                    ChartCategoryValue::Number(n) => n.to_string(),
                    ChartCategoryValue::Blank => continue,
                };
                out.push_str(&format!("<c:pt idx=\"{i}\"><c:v>{}</c:v></c:pt>", escape(&text)));
            }
            out.push_str(&format!("</c:{cache}>"));
            if series.category_reference.is_some() { out.push_str("</c:strRef>"); }
            out
        };
        let tag = if scatter { "xVal" } else { "cat" };
        format!("<c:{tag} xmlns:c=\"{ns}\">{body}</c:{tag}>")
    };
    let tag = if scatter { "yVal" } else { "val" };
    let vals = format!("<c:{tag} xmlns:c=\"{ns}\">{}</c:{tag}>", numeric(&series.values, series.values_reference.as_deref()));
    (name, cats, vals)
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
    fn series(series: &ChartSeries, index: usize, categories: &[String], scatter: bool) -> String {
        let (name, cats, values) = series_data_parts(series, categories, scatter, C);
        let mut out = format!("<c:ser><c:idx val=\"{index}\"/><c:order val=\"{index}\"/>{name}");
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
        out.push_str(&cats);
        out.push_str(&values);
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


fn patch_series_data(xml: &str, before: &Chart, after: &Chart, ns: &str, drawing_ns: &str) -> Result<String, XlsxError> {
    // Keep the source namespace bindings while inspecting individual series.
    let mut reader = Reader::from_str(xml);
    let (open, close) = loop {
        let start = reader.buffer_position() as usize;
        match reader.read_event().map_err(XlsxError::Xml)? {
            Event::Start(element) if local_name(element.name().as_ref()) == "chartSpace" => {
                let finish = reader.buffer_position() as usize;
                let close = format!("</{}>", String::from_utf8_lossy(element.name().as_ref()));
                break (xml[start..finish].to_string(), close);
            }
            Event::Eof => return Err(XlsxError::InvalidData("chart root is missing".into())),
            _ => {}
        }
    };
    let (plots, _) = children_at(xml, &["chartSpace", "chart", "plotArea"])?;
    let plots: Vec<_> = plots.iter().filter(|(name, _, _)| name.ends_with("Chart")).collect();
    if plots.len() != 1 {
        return Err(XlsxError::InvalidData("series editing requires an unambiguous source plot".into()));
    }
    let plot = plots[0].0.as_str();
    let path = ["chartSpace", "chart", "plotArea", plot];
    let (children, end) = children_at(xml, &path)?;
    let existing: Vec<_> = children.iter().filter(|(name, _, _)| name == "ser").collect();
    let scatter = plot == "scatterChart";
    let replace = |fragment: &str, tag: &str, replacement: &str| -> Result<String, XlsxError> {
        let wrapped = format!("{open}{fragment}{close}");
        let (children, end) = children_at(&wrapped, &["chartSpace", "ser"])?;
        let (start, finish) = if let Some((_, start, finish)) = children.iter().find(|(name, _, _)| name == tag) {
            (*start, *finish)
        } else {
            if replacement.is_empty() { return Ok(fragment.to_string()); }
            let at = children.iter().find(|(name, _, _)| match tag {
                "tx" => !matches!(name.as_str(), "idx" | "order"),
                "cat" | "xVal" => matches!(name.as_str(), "val" | "yVal" | "smooth" | "extLst"),
                _ => matches!(name.as_str(), "smooth" | "extLst"),
            }).map_or(end, |(_, start, _)| *start);
            (at, at)
        };
        let revised = format!("{}{}{}", &wrapped[..start], replacement, &wrapped[finish..]);
        Ok(revised[open.len()..revised.len() - close.len()].to_string())
    };
    let mut edits = Vec::new();
    for (i, (_, start, finish)) in existing.iter().enumerate() {
        let replacement = if let Some(series) = after.series.get(i) {
            let (name, cats, values) = series_data_parts(series, &after.categories, scatter, ns);
            let previous = before.series.get(i);
            let name_changed = previous.is_none_or(|old| old.name != series.name || old.name_reference != series.name_reference);
            let cats_changed = previous.is_none_or(|old|
                old.category_reference != series.category_reference || old.category_values != series.category_values
                || ((old.category_values.is_none() || series.category_values.is_none()) && before.categories != after.categories));
            let values_changed = previous.is_none_or(|old| old.values != series.values || old.values_reference != series.values_reference);
            let mut fragment = xml[*start..*finish].to_string();
            if name_changed { fragment = replace(&fragment, "tx", &name)?; }
            if cats_changed { fragment = replace(&fragment, if scatter { "xVal" } else { "cat" }, &cats)?; }
            if values_changed { fragment = replace(&fragment, if scatter { "yVal" } else { "val" }, &values)?; }
            fragment
        } else { String::new() };
        edits.push((*start, *finish, replacement));
    }
    let mut appended = String::new();
    for (i, series) in after.series.iter().enumerate().skip(existing.len()) {
        // Reuse the new-series serializer for modeled formatting, then insert
        // only the series, retaining the original plot, axes and extensions.
        let kind = if scatter { "scatter" } else { "bar" };
        let generated = new_chart_xml(&Chart {
            kind: kind.into(), categories: after.categories.clone(),
            series: vec![series.clone()], ..Default::default()
        })?;
        let tag = if scatter { "scatterChart" } else { "barChart" };
        let (ranges, _) = children_at(&generated, &["chartSpace", "chart", "plotArea", tag])?;
        let (_, begin, finish) = ranges.iter().find(|(name, _, _)| name == "ser").unwrap();
        let fragment = generated[*begin..*finish]
            .replacen("<c:ser>", &format!("<c:ser xmlns:c=\"{ns}\" xmlns:a=\"{drawing_ns}\">"), 1)
            .replace(&format!("xmlns:c=\"{C}\""), &format!("xmlns:c=\"{ns}\""))
            .replacen("<c:idx val=\"0\"/>", &format!("<c:idx val=\"{i}\"/>"), 1)
            .replacen("<c:order val=\"0\"/>", &format!("<c:order val=\"{i}\"/>"), 1);
        appended.push_str(&fragment);
    }
    if !appended.is_empty() {
        let at = existing.last().map(|(_, _, end)| *end).unwrap_or_else(||
            children.iter().find(|(name, _, _)| matches!(name.as_str(), "dLbls" | "dropLines" | "hiLowLines" | "upDownBars" | "marker" | "smooth" | "axId" | "extLst"))
                .map_or(end, |(_, start, _)| *start));
        edits.push((at, at, appended));
    }
    edits.sort_by_key(|(start, finish, _)| (*start, *finish));
    let mut out = xml.to_string();
    for (start, finish, replacement) in edits.into_iter().rev() {
        out.replace_range(start..finish, &replacement);
    }
    Ok(out)
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
    if series_data_changed(before, after) {
        output = patch_series_data(&output, before, after, chart_ns, drawing_ns)?;
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

#[cfg(test)]
mod series_data_tests {
    use super::*;
    use crate::ir::{ChartCategoryValue as Category, ChartSeries};

    #[test]
    fn series_references_and_category_types_roundtrip_independently() {
        for kind in ["bar", "scatter"] {
            let first = ChartSeries {
                name: "A & B".into(), name_reference: Some("'Sales & Costs'!$B$1".into()),
                values_reference: Some("'Sales & Costs'!$B$2:$B$4".into()),
                category_reference: Some("'Sales & Costs'!$A$2:$A$4".into()),
                category_values: Some(vec![Category::Text("1".into()), Category::Text(" 2 ".into()), Category::Text("3".into())]),
                values: vec![Some(1.25), None, Some(3.75)], ..Default::default()
            };
            let second = ChartSeries {
                name: "Numeric".into(),
                category_values: Some(vec![Category::Number(10.0), Category::Number(20.0)]),
                values: vec![Some(4.0), Some(5.0)], ..Default::default()
            };
            let expected = Chart { kind: kind.into(), series: vec![first, second],
                categories: vec!["legacy fallback must not replace per-series data".into()], ..Default::default() };
            let xml = new_chart_xml(&expected).unwrap();
            let actual = crate::chart::parse_chart_xml(&xml, &crate::parser::Theme::default()).unwrap();
            for (a, b) in actual.series.iter().zip(&expected.series) {
                assert!(series_data_equal(a, b), "{kind}: {a:?} != {b:?}");
            }
            assert_eq!(actual.series.len(), 2);
            assert!(xml.contains("<c:strRef><c:f>&apos;Sales &amp; Costs&apos;!$A$2:$A$4</c:f>")
                || xml.contains("<c:strRef><c:f>'Sales &amp; Costs'!$A$2:$A$4</c:f>"));
        }
    }

    #[test]
    fn editing_data_preserves_source_formatting_extensions_and_other_series() {
        for strict in [false, true] {
            let initial = Chart { kind: "bar".into(), series: vec![
                ChartSeries { name: "First".into(), values: vec![Some(1.0)], ..Default::default() },
                ChartSeries { name: "Second".into(), values: vec![Some(2.0)], ..Default::default() },
            ], ..Default::default() };
            let mut xml = new_chart_xml(&initial).unwrap();
            let extension = "<c:extLst><c:ext uri=\"retained\"><x:keep xmlns:x=\"urn:retained\">unchanged</x:keep></c:ext></c:extLst>";
            let foreign = "<x:val xmlns:x=\"urn:foreign\">foreign value</x:val>";
            xml = xml.replacen("</c:ser>", &format!("{foreign}{extension}</c:ser>"), 1);
            if strict { xml = xml.replace(C, "http://purl.oclc.org/ooxml/drawingml/chart")
                .replace(A, "http://purl.oclc.org/ooxml/drawingml/main"); }
            let before = crate::chart::parse_chart_xml(&xml, &crate::parser::Theme::default()).unwrap();
            let mut after = before.clone();
            after.series[0].values = vec![Some(7.0)];
            after.series[0].values_reference = Some("Sheet1!$B$2".into());
            let patched = patch_chart(&xml, &before, &after).unwrap();
            assert!(patched.contains(extension) && patched.contains(foreign));
            let ranges = |text: &str| children_at(text, &["chartSpace", "chart", "plotArea", "barChart"]).unwrap().0
                .into_iter().filter(|(name, _, _)| name == "ser").map(|(_, a, b)| text[a..b].to_string()).collect::<Vec<_>>();
            assert_eq!(ranges(&xml)[1], ranges(&patched)[1]);
            let actual = crate::chart::parse_chart_xml(&patched, &crate::parser::Theme::default()).unwrap();
            assert!(series_data_equal(&actual.series[0], &after.series[0]));
        }
    }
}
