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
    fn series(series: &ChartSeries, index: usize, categories: &[String], scatter: bool, options: &crate::ir::ChartPlotOptions, kind: &str) -> String {
        let (name, cats, values) = series_data_parts(series, categories, scatter, C);
        let mut out = format!("<c:ser><c:idx val=\"{index}\"/><c:order val=\"{index}\"/>{name}");
        if scatter && options.connected == Some(false) {
            out.push_str("<c:spPr><a:ln><a:noFill/></a:ln></c:spPr>");
        } else if let Some(line) = &series.line { out.push_str(&format!("<c:spPr>{}</c:spPr>", paint(line))); }
        if let Some(value) = &series.marker { out.push_str(&marker(value)); }
        else if matches!(kind, "line" | "scatter" | "radar") && !(kind == "line" && options.three_dimensional) {
            out.push_str(&format!("<c:marker><c:symbol val=\"{}\"/></c:marker>",
                if options.markers.unwrap_or(true) { "auto" } else { "none" }));
        }
        if matches!(kind, "pie" | "doughnut") {
            if let Some(amount) = series.explosion.or(options.explosion) {
                out.push_str(&format!("<c:explosion val=\"{amount}\"/>"));
            }
        }
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
        if matches!(kind, "line" | "scatter") { out.push_str(&format!("<c:smooth val=\"{}\"/>", u8::from(options.smooth))); }
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
    use crate::ir::ChartStacking;
    let options = &chart.plot_options;
    let grouping = match options.stacking {
        ChartStacking::Stacked => "stacked", ChartStacking::Percent => "percentStacked",
        ChartStacking::None => if chart.kind == "bar" && !options.depth_series { "clustered" } else { "standard" },
    };
    let (plot, settings, axes) = match chart.kind.as_str() {
        "bar" => (if options.three_dimensional { "bar3DChart" } else { "barChart" }, format!("<c:barDir val=\"{}\"/><c:grouping val=\"{grouping}\"/><c:varyColors val=\"0\"/>",
            if options.horizontal { "bar" } else { "col" }), true),
        "line" => (if options.three_dimensional { "line3DChart" } else { "lineChart" }, format!("<c:grouping val=\"{grouping}\"/><c:varyColors val=\"0\"/>"), true),
        "area" => (if options.three_dimensional { "area3DChart" } else { "areaChart" }, format!("<c:grouping val=\"{grouping}\"/><c:varyColors val=\"0\"/>"), true),
        "scatter" => ("scatterChart", format!("<c:scatterStyle val=\"{}\"/><c:varyColors val=\"0\"/>",
            if options.smooth { "smoothMarker" } else { "lineMarker" }), true),
        "pie" => (if options.three_dimensional { "pie3DChart" } else { "pieChart" }, "<c:varyColors val=\"1\"/>".into(), false),
        "doughnut" => ("doughnutChart", "<c:varyColors val=\"1\"/>".into(), false),
        "radar" => ("radarChart", format!("<c:radarStyle val=\"{}\"/><c:varyColors val=\"0\"/>",
            if options.filled { "filled" } else { "marker" }), true),
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
    out.push_str(&format!("<c:autoTitleDeleted val=\"{}\"/>", u8::from(chart.title.is_none())));
    if options.three_dimensional {
        out.push_str(&format!("<c:view3D><c:rotX val=\"15\"/><c:rotY val=\"{}\"/><c:rAngAx val=\"0\"/></c:view3D>",
            if chart.kind == "pie" { 0 } else { 20 }));
    }
    out.push_str(&format!("<c:plotArea>{}<c:{plot}>{settings}", layout(chart.plot.as_ref(), true)));
    for (index, value) in chart.series.iter().enumerate() { out.push_str(&series(value, index, &chart.categories, chart.kind == "scatter", options, &chart.kind)); }
    if chart.kind == "line" && !options.three_dimensional {
        let markers = chart.series.is_empty() || options.markers.unwrap_or(true);
        out.push_str(&format!("<c:marker val=\"{}\"/><c:smooth val=\"{}\"/>", u8::from(markers), u8::from(options.smooth)));
    }
    if chart.kind == "bar" && options.three_dimensional {
        use crate::ir::ChartBarShape;
        let shape = match options.bar_shape { ChartBarShape::Box => "box", ChartBarShape::Cylinder => "cylinder",
            ChartBarShape::Cone => "cone", ChartBarShape::Pyramid => "pyramid" };
        out.push_str(&format!("<c:shape val=\"{shape}\"/>"));
    }
    if chart.kind == "bar" && !options.three_dimensional && options.stacking != ChartStacking::None { out.push_str("<c:overlap val=\"100\"/>"); }
    if chart.kind == "doughnut" { out.push_str("<c:firstSliceAng val=\"0\"/><c:holeSize val=\"50\"/>"); }
    let depth_axis = options.three_dimensional && (options.depth_series || chart.kind == "line"
        || (chart.kind == "area" && options.stacking == ChartStacking::None));
    if axes {
        out.push_str("<c:axId val=\"1\"/><c:axId val=\"2\"/>");
        if depth_axis { out.push_str("<c:axId val=\"3\"/>"); }
    }
    out.push_str(&format!("</c:{plot}>"));
    if axes {
        out.push_str(&axis(chart.category_axis.as_ref(), if chart.kind == "scatter" { "valAx" } else { "catAx" }, 1, 2, "b"));
        out.push_str(&axis(chart.value_axis.as_ref(), "valAx", 2, 1, "l"));
        if depth_axis { out.push_str(&axis(None, "serAx", 3, 2, "b")); }
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
        "view3D" => ranges.iter().find(|(found, _, _)| matches!(found.as_str(), "floor" | "sideWall" | "backWall" | "plotArea" | "legend"))
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



fn series_fragment_envelope(xml: &str, at: usize, ns: &str) -> Result<(String, String), XlsxError> {
    use std::collections::BTreeMap;
    let mut reader = Reader::from_str(xml);
    let mut bindings: BTreeMap<String, String> = BTreeMap::new();
    let mut scopes = Vec::new();
    loop {
        let position = reader.buffer_position() as usize;
        let event = reader.read_event().map_err(XlsxError::Xml)?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                let previous = bindings.clone();
                for attribute in element.attributes().flatten() {
                    let name = String::from_utf8_lossy(attribute.key.as_ref());
                    let prefix = if name == "xmlns" { Some("") } else { name.strip_prefix("xmlns:") };
                    if let Some(prefix) = prefix {
                        let value = attribute.unescape_value().map(|s| s.into_owned())
                            .unwrap_or_else(|_| String::from_utf8_lossy(&attribute.value).into_owned());
                        bindings.insert(prefix.to_string(), value);
                    }
                }
                if position == at {
                    let mut prefix = "oxiChartPatch".to_string();
                    while bindings.contains_key(&prefix) { prefix.push('x'); }
                    let mut open = format!("<{prefix}:chartSpace xmlns:{prefix}=\"{ns}\"");
                    for (name, value) in bindings {
                        if name.is_empty() { open.push_str(&format!(" xmlns=\"{}\"", escape(&value))); }
                        else { open.push_str(&format!(" xmlns:{name}=\"{}\"", escape(&value))); }
                    }
                    open.push('>');
                    return Ok((open, format!("</{prefix}:chartSpace>")));
                }
                if matches!(event, Event::Start(_)) { scopes.push(previous); } else { bindings = previous; }
            }
            Event::End(_) => { bindings = scopes.pop().unwrap_or_default(); }
            Event::Eof => return Err(XlsxError::InvalidData("series namespace context is missing".into())),
            _ => {}
        }
    }
}

fn patch_series_data(xml: &str, before: &Chart, after: &Chart, ns: &str, drawing_ns: &str) -> Result<String, XlsxError> {
    // Keep the source namespace bindings while inspecting individual series.
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
    let replace = |fragment: &str, tag: &str, replacement: &str, open: &str, close: &str| -> Result<String, XlsxError> {
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
        let (open, close) = series_fragment_envelope(xml, *start, ns)?;
        let replacement = if let Some(series) = after.series.get(i) {
            let (name, cats, values) = series_data_parts(series, &after.categories, scatter, ns);
            let previous = before.series.get(i);
            let name_changed = previous.is_none_or(|old| old.name != series.name || old.name_reference != series.name_reference);
            let cats_changed = previous.is_none_or(|old|
                old.category_reference != series.category_reference || old.category_values != series.category_values
                || ((old.category_values.is_none() || series.category_values.is_none()) && before.categories != after.categories));
            let values_changed = previous.is_none_or(|old| old.values != series.values || old.values_reference != series.values_reference);
            let mut fragment = xml[*start..*finish].to_string();
            if name_changed { fragment = replace(&fragment, "tx", &name, &open, &close)?; }
            if cats_changed { fragment = replace(&fragment, if scatter { "xVal" } else { "cat" }, &cats, &open, &close)?; }
            if values_changed { fragment = replace(&fragment, if scatter { "yVal" } else { "val" }, &values, &open, &close)?; }
            fragment
        } else { String::new() };
        edits.push((*start, *finish, replacement));
    }
    let mut appended = String::new();
    for (i, series) in after.series.iter().enumerate().skip(existing.len()) {
        // Reuse the new-series serializer for modeled formatting, then insert
        // only the series, retaining the original plot, axes and extensions.
        let kind = if scatter { "scatter" } else if matches!(before.kind.as_str(), "pie" | "doughnut") { &before.kind } else { "bar" };
        let generated = new_chart_xml(&Chart {
            kind: kind.into(), categories: after.categories.clone(),
            plot_options: after.plot_options.clone(), series: vec![series.clone()], ..Default::default()
        })?;
        let tag = if scatter { "scatterChart" } else if kind == "pie" { if after.plot_options.three_dimensional { "pie3DChart" } else { "pieChart" } }
            else if kind == "doughnut" { "doughnutChart" } else if after.plot_options.three_dimensional { "bar3DChart" } else { "barChart" };
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


pub(crate) fn plot_options_changed(before: &Chart, after: &Chart) -> bool {
    before.kind != after.kind || before.plot_options != after.plot_options
        || before.series.iter().zip(&after.series).any(|(a, b)| a.explosion != b.explosion)
}

fn patch_plot_options(xml: &str, before: &Chart, after: &Chart, ns: &str, drawing_ns: &str) -> Result<String, XlsxError> {
    let area = ["chartSpace", "chart", "plotArea"];
    let (old_children, _) = children_at(xml, &area)?;
    let old_plots: Vec<_> = old_children.iter().filter(|(name, _, _)| name.ends_with("Chart")).collect();
    if old_plots.len() != 1 {
        return Err(XlsxError::InvalidData("plot editing requires an unambiguous source plot".into()));
    }
    let (old_tag, old_start, old_end) = old_plots[0];
    let old_path = ["chartSpace", "chart", "plotArea", old_tag.as_str()];
    let (old_settings, _) = children_at(xml, &old_path)?;
    let mut modeled = after.clone();
    modeled.shapes.clear(); // Annotations belong to the retained source chart.
    let generated = new_chart_xml(&modeled)?
        .replace(&format!("xmlns:c=\"{C}\""), &format!("xmlns:c=\"{ns}\""))
        .replace(&format!("xmlns:a=\"{A}\""), &format!("xmlns:a=\"{drawing_ns}\""));
    let (new_children, _) = children_at(&generated, &area)?;
    let (new_tag, new_start, new_end) = new_children.iter().find(|(name, _, _)| name.ends_with("Chart"))
        .ok_or_else(|| XlsxError::InvalidData("new plot is missing".into()))?;
    let mut plot = generated[*new_start..*new_end].to_string();
    let new_path = ["chartSpace", "chart", "plotArea", new_tag.as_str()];
    let (new_settings, _) = children_at(&generated, &new_path)?;
    let old_series: Vec<_> = old_settings.iter().filter(|(name, _, _)| name == "ser").collect();
    let new_series: Vec<_> = new_settings.iter().filter(|(name, _, _)| name == "ser").collect();
    let (old_series_fields, _) = children_at(xml, &["chartSpace", "chart", "plotArea", old_tag.as_str(), "ser"])
        .unwrap_or_default();
    for (i, (_, begin, end)) in new_series.iter().enumerate().rev() {
        if let Some((_, old_begin, old_end)) = old_series.get(i) {
            let additions: String = old_series_fields.iter().filter(|(name, a, b)|
                name == "extLst" && a >= old_begin && b <= old_end).map(|(_, a, b)| &xml[*a..*b]).collect();
            if !additions.is_empty() {
                let at = end - new_start - "</c:ser>".len();
                plot.insert_str(at, &additions);
            }
        }
        let _ = begin;
    }
    let same_kind = before.kind == after.kind;
    let child_position = |fragment: &str, successors: &[&str]| -> Result<usize, XlsxError> {
        let prefix = format!("<c:chartSpace xmlns:c=\"{ns}\" xmlns:a=\"{drawing_ns}\"><c:chart><c:plotArea>");
        let wrapped = format!("{prefix}{fragment}</c:plotArea></c:chart></c:chartSpace>");
        let (fields, end) = children_at(&wrapped, &["chartSpace", "chart", "plotArea", new_tag])?;
        Ok(fields.iter().filter(|(name, _, _)| successors.contains(&name.as_str()))
            .map(|(_, a, _)| *a).min().unwrap_or(end) - prefix.len())
    };
    for (name, begin, end) in &old_settings {
        let retain = name == "extLst" || name == "dLbls" || (old_tag == new_tag
            && !matches!(name.as_str(), "ser" | "axId" | "barDir" | "grouping" | "scatterStyle" | "radarStyle" | "marker" | "smooth" | "overlap" | "varyColors" | "shape"));
        if !retain { continue; }
        // These source properties precede the plot's axes; extension lists are last.
        let successors: &[&str] = match name.as_str() {
            "dLbls" => &["dropLines", "hiLowLines", "upDownBars", "gapWidth", "gapDepth", "overlap", "shape", "serLines", "marker", "smooth", "firstSliceAng", "holeSize", "axId", "extLst"],
            "gapWidth" => &["gapDepth", "overlap", "serLines", "shape", "axId", "extLst"],
            "gapDepth" => &["shape", "axId", "extLst"],
            "dropLines" => &["hiLowLines", "upDownBars", "marker", "smooth", "axId", "extLst"],
            "hiLowLines" => &["upDownBars", "marker", "smooth", "axId", "extLst"],
            "upDownBars" => &["marker", "smooth", "axId", "extLst"],
            "extLst" => &[],
            _ => &["axId", "extLst"],
        };
        let prefix = format!("<c:chartSpace xmlns:c=\"{ns}\" xmlns:a=\"{drawing_ns}\"><c:chart><c:plotArea>");
        let wrapped = format!("{prefix}{plot}</c:plotArea></c:chart></c:chartSpace>");
        let (fields, _) = children_at(&wrapped, &["chartSpace", "chart", "plotArea", new_tag])?;
        if let Some((_, a, b)) = fields.iter().find(|(tag, _, _)| tag == name) {
            // A retained setting replaces a generated default rather than
            // introducing duplicate single-valued OOXML children.
            plot.replace_range(a - prefix.len()..b - prefix.len(), &xml[*begin..*end]);
        } else {
            let at = child_position(&plot, successors)?;
            plot.insert_str(at, &xml[*begin..*end]);
        }
    }
    let axis_id = |fragment: &str| -> Result<String, XlsxError> {
        let mut reader = Reader::from_str(fragment);
        match reader.read_event().map_err(XlsxError::Xml)? {
            Event::Start(e) | Event::Empty(e) => get_attr(&e, "val")
                .ok_or_else(|| XlsxError::InvalidData("axis identity is missing".into())),
            _ => Err(XlsxError::InvalidData("axis identity is missing".into())),
        }
    };
    let old_ids: Vec<_> = old_settings.iter().filter(|(name, _, _)| name == "axId")
        .map(|(_, a, b)| axis_id(&xml[*a..*b])).collect::<Result<_, _>>()?;
    let new_ids = new_settings.iter().filter(|(name, _, _)| name == "axId").count();
    let keep_axes = same_kind && old_ids.len() == new_ids;
    if keep_axes {
        for i in 0..new_ids { plot = plot.replace(&format!("<c:axId val=\"{}\"/>", i + 1), &format!("<c:axId val=\"oxi-axis-{i}\"/>")); }
        for (i, id) in old_ids.iter().enumerate() { plot = plot.replace(&format!("val=\"oxi-axis-{i}\""), &format!("val=\"{id}\"")); }
    }
    plot = plot.replacen(&format!("<c:{new_tag}>"), &format!("<c:{new_tag} xmlns:c=\"{ns}\" xmlns:a=\"{drawing_ns}\">"), 1);
    let mut replacement = plot;
    if !keep_axes {
        for (tag, a, b) in &new_children {
            if matches!(tag.as_str(), "catAx" | "valAx" | "dateAx" | "serAx") {
                replacement.push_str(&generated[*a..*b].replacen(&format!("<c:{tag}>"),
                    &format!("<c:{tag} xmlns:c=\"{ns}\" xmlns:a=\"{drawing_ns}\">"), 1));
            }
        }
    }
    let mut edits = vec![(*old_start, *old_end, replacement)];
    if !keep_axes {
        for (tag, a, b) in &old_children {
            if matches!(tag.as_str(), "catAx" | "valAx" | "dateAx" | "serAx") { edits.push((*a, *b, String::new())); }
        }
    }
    edits.sort_by_key(|(a, b, _)| (*a, *b));
    let mut output = xml.to_string();
    for (a, b, replacement) in edits.into_iter().rev() { output.replace_range(a..b, &replacement); }
    if before.plot_options.three_dimensional != after.plot_options.three_dimensional {
        if after.plot_options.three_dimensional {
            let (fields, _) = children_at(&generated, &["chartSpace", "chart"])?;
            if let Some((_, a, b)) = fields.iter().find(|(name, _, _)| name == "view3D") {
                let view = generated[*a..*b].replacen("<c:view3D>", &format!("<c:view3D xmlns:c=\"{ns}\">"), 1);
                output = replace_child(&output, "view3D", &view)?;
            }
        } else {
            for tag in ["view3D", "floor", "sideWall", "backWall"] { output = replace_child(&output, tag, "")?; }
        }
    }
    if keep_axes && before.plot_options.horizontal != after.plot_options.horizontal && after.kind == "bar" {
        for tag in ["catAx", "dateAx", "valAx"] {
            if let Ok((fields, _)) = children_at(&output, &["chartSpace", "chart", "plotArea", tag]) {
                if let Some((_, a, b)) = fields.iter().find(|(name, _, _)| name == "axPos") {
                    let position = if tag == "valAx" { if after.plot_options.horizontal { "b" } else { "l" } }
                        else if after.plot_options.horizontal { "l" } else { "b" };
                    output.replace_range(*a..*b, &format!("<c:axPos xmlns:c=\"{ns}\" val=\"{position}\"/>"));
                }
            }
        }
    }
    Ok(output)
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
    if plot_options_changed(before, after) { output = patch_plot_options(&output, before, after, chart_ns, drawing_ns)?; }
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
    if series_data_changed(before, after) && !plot_options_changed(before, after) {
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

#[cfg(test)]
mod plot_options_tests {
    use super::*;
    use crate::ir::{ChartAxis, ChartPlotOptions, ChartSeries, ChartStacking};

    #[test]
    fn native_marker_line_and_grouping_choices_survive_serialization() {
        for (kind, options) in [
            ("bar", ChartPlotOptions { horizontal: true, stacking: ChartStacking::Percent, ..Default::default() }),
            ("line", ChartPlotOptions { stacking: ChartStacking::Stacked, markers: Some(false), ..Default::default() }),
            ("line", ChartPlotOptions { markers: Some(true), ..Default::default() }),
            ("scatter", ChartPlotOptions { markers: Some(true), connected: Some(false), ..Default::default() }),
            ("scatter", ChartPlotOptions { markers: Some(false), connected: Some(true), smooth: true, ..Default::default() }),
            ("radar", ChartPlotOptions { markers: Some(false), ..Default::default() }),
            ("radar", ChartPlotOptions { markers: Some(false), filled: true, ..Default::default() }),
        ] {
            let chart = Chart { kind: kind.into(), plot_options: options.clone(),
                series: vec![ChartSeries { name: "Revenue".into(), values: vec![Some(1.0)], ..Default::default() }],
                ..Default::default() };
            let xml = new_chart_xml(&chart).unwrap();
            let restored = crate::chart::parse_chart_xml(&xml, &crate::parser::Theme::default()).unwrap();
            assert_eq!(restored.plot_options, options, "{kind}");
            if kind == "line" && options.markers == Some(false) { assert!(xml.contains("<c:symbol val=\"none\"/>")); }
            if kind == "scatter" && options.connected == Some(false) { assert!(xml.contains("<a:ln><a:noFill/></a:ln>")); }
        }
    }

    #[test]
    fn changing_plot_settings_keeps_source_extensions_axes_and_label_order() {
        for strict in [false, true] {
            let chart = Chart { kind: "bar".into(), series: vec![ChartSeries {
                name: "Revenue".into(), values: vec![Some(1.0)], ..Default::default()
            }], value_axis: Some(ChartAxis { number_format: Some("$0.00".into()), ..Default::default() }), ..Default::default() };
            let mut xml = new_chart_xml(&chart).unwrap();
            let labels = "<c:dLbls><c:showVal val=\"1\"/></c:dLbls>";
            let extension = "<c:extLst><c:ext uri=\"kept\"><x:keep xmlns:x=\"urn:retained\">retained</x:keep></c:ext></c:extLst>";
            xml = xml.replacen("</c:barChart>", &format!("{labels}<c:gapWidth val=\"83\"/>{extension}</c:barChart>"), 1)
                .replacen("</c:ser>", &format!("{extension}</c:ser>"), 1)
                .replacen("</c:valAx>", &format!("{extension}</c:valAx>"), 1);
            if strict { xml = xml.replace(C, "http://purl.oclc.org/ooxml/drawingml/chart").replace(A, "http://purl.oclc.org/ooxml/drawingml/main"); }
            let before = crate::chart::parse_chart_xml(&xml, &crate::parser::Theme::default()).unwrap();
            for kind in ["bar", "line", "scatter", "doughnut"] {
                let mut after = before.clone();after.kind = kind.into();
                after.plot_options = ChartPlotOptions { horizontal: kind == "bar", stacking: if kind == "bar" { ChartStacking::Stacked } else { ChartStacking::None },
                    markers: if matches!(kind, "line" | "scatter") { Some(false) } else { None },
                    connected: if kind == "scatter" { Some(false) } else { None }, ..Default::default() };
                let patched = patch_chart(&xml, &before, &after).unwrap();
                let tag = match kind { "bar" => "barChart", "line" => "lineChart", "scatter" => "scatterChart", _ => "doughnutChart" };
                let (fields, _) = children_at(&patched, &["chartSpace", "chart", "plotArea", tag]).unwrap();
                let names: Vec<_> = fields.iter().map(|(n, _, _)| n.as_str()).collect();
                assert!(names.iter().position(|n| *n == "dLbls").unwrap() > names.iter().position(|n| *n == "ser").unwrap());
                if let Some(hole) = names.iter().position(|n| *n == "holeSize") {
                    assert!(names.iter().position(|n| *n == "dLbls").unwrap() < hole);
                }
                assert!(patched.contains(extension));
                assert!(patched.contains("<x:keep xmlns:x=\"urn:retained\">retained</x:keep>"));
                if kind == "bar" { assert!(patched.contains("<c:gapWidth val=\"83\"/>")); assert!(patched.contains("$0.00")); }
            }
        }
    }
}

#[cfg(test)]
mod scoped_namespace_tests {
    use super::*;
    #[test]
    fn data_and_plot_edits_inherit_namespace_bindings_from_plot_ancestors() {
        let chart = Chart { kind: "bar".into(), series: vec![crate::ir::ChartSeries {
            name: "Revenue".into(), values: vec![Some(1.0)], ..Default::default()
        }], ..Default::default() };
        let mut xml = new_chart_xml(&chart).unwrap().replace("<c:", "<q:").replace("</c:", "</q:");
        xml = xml.replacen(&format!("xmlns:c=\"{C}\""), &format!("xmlns:q=\"{C}\" xmlns:c=\"urn:root-prefix\""), 1)
            .replacen("<q:plotArea>", &format!("<q:plotArea xmlns:c=\"{C}\">"), 1);
        let a = xml.find("<q:ser>").unwrap();let b = xml.find("</q:ser>").unwrap() + "</q:ser>".len();
        let series = xml[a..b].replace("<q:", "<c:").replace("</q:", "</c:");
        xml.replace_range(a..b, &series);
        let before = crate::chart::parse_chart_xml(&xml, &crate::parser::Theme::default()).unwrap();
        for change_kind in [false, true] {
            let mut after = before.clone();after.series[0].values = vec![Some(7.0)];
            if change_kind { after.kind = "line".into();after.plot_options.markers=Some(false); }
            let patched = patch_chart(&xml, &before, &after).unwrap();
            let parsed = crate::chart::parse_chart_xml(&patched, &crate::parser::Theme::default()).unwrap();
            assert_eq!(parsed.kind, after.kind);
            assert!(series_data_equal(&parsed.series[0], &after.series[0]));
        }
    }
}

#[cfg(test)]
mod spatial_plot_tests {
    use super::*;
    use crate::ir::{ChartBarShape, ChartPlotOptions, ChartSeries, ChartStacking};

    #[test]
    fn spatial_plot_choices_and_independent_slice_separation_roundtrip() {
        for (kind, options) in [
            ("bar", ChartPlotOptions { three_dimensional: true, ..Default::default() }),
            ("bar", ChartPlotOptions { three_dimensional: true, depth_series: true, bar_shape: ChartBarShape::Cylinder, ..Default::default() }),
            ("bar", ChartPlotOptions { three_dimensional: true, horizontal: true, stacking: ChartStacking::Percent, bar_shape: ChartBarShape::Pyramid, ..Default::default() }),
            ("line", ChartPlotOptions { three_dimensional: true, ..Default::default() }),
            ("area", ChartPlotOptions { three_dimensional: true, stacking: ChartStacking::Stacked, ..Default::default() }),
            ("pie", ChartPlotOptions { three_dimensional: true, explosion: Some(25), ..Default::default() }),
            ("doughnut", ChartPlotOptions { explosion: Some(25), ..Default::default() }),
        ] {
            let chart = Chart { kind: kind.into(), plot_options: options.clone(),
                series: vec![ChartSeries { name: "Revenue".into(), values: vec![Some(1.0)], explosion: options.explosion, ..Default::default() }],
                ..Default::default() };
            let xml = new_chart_xml(&chart).unwrap();
            let restored = crate::chart::parse_chart_xml(&xml, &crate::parser::Theme::default()).unwrap();
            assert_eq!(restored.kind, chart.kind);
            assert_eq!(restored.plot_options, options, "{kind}");
            assert_eq!(restored.series[0].explosion, chart.series[0].explosion);
            if kind == "bar" { assert!(!xml.contains("<c:overlap")); }
            if kind == "line" { assert!(!xml.contains("<c:marker")); }
        }
        let chart = Chart { kind: "pie".into(), series: vec![
            ChartSeries { explosion: Some(25), values: vec![Some(1.0)], ..Default::default() },
            ChartSeries { explosion: Some(0), values: vec![Some(2.0)], ..Default::default() },
        ], ..Default::default() };
        let xml = new_chart_xml(&chart).unwrap();
        let restored = crate::chart::parse_chart_xml(&xml, &crate::parser::Theme::default()).unwrap();
        assert_eq!(restored.series.iter().map(|s| s.explosion).collect::<Vec<_>>(), vec![Some(25), Some(0)]);
    }

    #[test]
    fn changing_dimensions_inserts_and_removes_view_and_depth_axes() {
        for strict in [false, true] {
            let initial = Chart { kind: "bar".into(), series: vec![ChartSeries {
                name: "Revenue".into(), values: vec![Some(1.0)], ..Default::default()
            }], ..Default::default() };
            let mut xml = new_chart_xml(&initial).unwrap();
            if strict { xml = xml.replace(C, "http://purl.oclc.org/ooxml/drawingml/chart")
                .replace(A, "http://purl.oclc.org/ooxml/drawingml/main"); }
            let mut before = crate::chart::parse_chart_xml(&xml, &crate::parser::Theme::default()).unwrap();
            for (three_dimensional, depth_series) in [(true, false), (true, true), (false, false)] {
                let mut after = before.clone();
                after.plot_options.three_dimensional = three_dimensional;
                after.plot_options.depth_series = depth_series;
                after.series[0].values = vec![Some(7.0)];
                xml = patch_chart(&xml, &before, &after).unwrap();
                assert_eq!(xml.contains(":view3D"), three_dimensional);
                assert_eq!(xml.contains(":serAx"), depth_series);
                let (fields, _) = children_at(&xml, &["chartSpace", "chart"]).unwrap();
                if three_dimensional {
                    let view = fields.iter().position(|(tag, _, _)| tag == "view3D").unwrap();
                    let plot = fields.iter().position(|(tag, _, _)| tag == "plotArea").unwrap();
                    assert!(view < plot);
                }
                before = crate::chart::parse_chart_xml(&xml, &crate::parser::Theme::default()).unwrap();
                assert_eq!(before.plot_options.three_dimensional, three_dimensional);
                assert_eq!(before.plot_options.depth_series, depth_series);
                assert_eq!(before.series[0].values, vec![Some(7.0)]);
            }
        }
    }
}

#[cfg(test)]
mod retained_spatial_properties_tests {
    use super::*;
    use crate::ir::{ChartBarShape, ChartPlotOptions, ChartSeries};
    #[test]
    fn subtype_edits_keep_retained_property_order_and_single_values() {
        let chart = Chart { kind: "bar".into(), plot_options: ChartPlotOptions { three_dimensional: true, ..Default::default() },
            series: vec![ChartSeries { values: vec![Some(1.0)], ..Default::default() }], ..Default::default() };
        let xml = new_chart_xml(&chart).unwrap().replace("<c:shape", "<c:gapWidth val=\"200\"/><c:gapDepth val=\"100\"/><c:shape")
            .replace("<c:rotX val=\"15\"/>", "<c:rotX val=\"39\"/>");
        let before = crate::chart::parse_chart_xml(&xml, &crate::parser::Theme::default()).unwrap();
        let mut after = before.clone();after.plot_options.bar_shape = ChartBarShape::Cylinder;
        let changed = patch_chart(&xml, &before, &after).unwrap();
        let (fields, _) = children_at(&changed, &["chartSpace", "chart", "plotArea", "bar3DChart"]).unwrap();
        let position = |name| fields.iter().position(|(tag, _, _)| tag == name).unwrap();
        assert!(position("gapWidth") < position("gapDepth"));
        assert!(position("gapDepth") < position("shape"));
        assert!(changed.contains("<c:rotX val=\"39\"/>"));
        assert_eq!(fields.iter().filter(|(tag, _, _)| tag == "shape").count(), 1);

        let chart = Chart { kind: "doughnut".into(), series: chart.series.clone(), ..Default::default() };
        let xml = new_chart_xml(&chart).unwrap().replace("<c:holeSize val=\"50\"/>", "<c:holeSize val=\"75\"/>")
            .replace("<c:firstSliceAng val=\"0\"/>", "<c:firstSliceAng val=\"35\"/>");
        let before = crate::chart::parse_chart_xml(&xml, &crate::parser::Theme::default()).unwrap();
        let mut after = before.clone();after.plot_options.explosion = Some(25);
        let changed = patch_chart(&xml, &before, &after).unwrap();
        let (fields, _) = children_at(&changed, &["chartSpace", "chart", "plotArea", "doughnutChart"]).unwrap();
        for (tag, expected) in [("holeSize", "75"), ("firstSliceAng", "35")] {
            let values: Vec<_> = fields.iter().filter(|(name, _, _)| name == tag)
                .map(|(_, a, b)| &changed[*a..*b]).collect();
            assert_eq!(values.len(), 1, "{tag}");
            assert!(values[0].contains(&format!("val=\"{expected}\"")));
        }
    }
}
