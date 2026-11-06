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
    a.name == b.name && a.explicit_empty_name == b.explicit_empty_name && a.values == b.values
        && a.name_reference == b.name_reference && a.values_reference == b.values_reference
        && a.category_reference == b.category_reference && a.category_values == b.category_values
        && a.text_category_gaps == b.text_category_gaps
        && a.category_values_filtered == b.category_values_filtered
        && a.bubble_sizes == b.bubble_sizes && a.bubble_sizes_reference == b.bubble_sizes_reference
        && a.shaded_bubbles == b.shaded_bubbles
}

pub(crate) fn series_data_changed(a: &Chart, b: &Chart) -> bool {
    a.series.len() != b.series.len()
        || a.series.iter().zip(&b.series).any(|(a, b)| !series_data_equal(a, b))
        || (a.categories != b.categories && a.series.iter().zip(&b.series)
            .any(|(a, b)| a.category_values.is_none() || b.category_values.is_none()))
}

fn numeric_series_data(values: &[Option<f64>], reference: Option<&str>) -> String {
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


fn bubble_series_data(series: &crate::ir::ChartSeries, ns: &str) -> String {
    if series.bubble_sizes.is_none() && series.bubble_sizes_reference.is_none() { return String::new(); }
    format!("<c:bubbleSize xmlns:c=\"{ns}\">{}</c:bubbleSize>",
        numeric_series_data(series.bubble_sizes.as_deref().unwrap_or_default(), series.bubble_sizes_reference.as_deref()))
}

/// References and their caches are kept together; literal text is not inferred
/// to be numeric merely because it can be parsed as a number.
fn series_data_parts(series: &crate::ir::ChartSeries, shared: &[String], scatter: bool, ns: &str) -> (String, String, String) {
    use crate::ir::ChartCategoryValue;
    let name = if let Some(reference) = &series.name_reference {
        format!("<c:strRef><c:f>{}</c:f><c:strCache><c:ptCount val=\"1\"/><c:pt idx=\"0\"><c:v>{}</c:v></c:pt></c:strCache></c:strRef>",
            escape(reference.trim_start_matches('=')), escape(&series.name))
    } else { format!("<c:v>{}</c:v>", escape(&series.name)) };
    let name = if series.name.is_empty() && !series.explicit_empty_name && series.name_reference.is_none() {
        String::new()
    } else { format!("<c:tx xmlns:c=\"{ns}\">{name}</c:tx>") };
    let legacy: Vec<_> = if scatter {
        (0..series.values.len()).map(|i| ChartCategoryValue::Number(
            shared.get(i).and_then(|v| v.parse::<f64>().ok()).unwrap_or((i + 1) as f64))).collect()
    } else { shared.iter().map(|v| ChartCategoryValue::Text(v.clone())).collect() };
    let categories = series.category_values.as_deref().unwrap_or(&legacy);
    let cats = if categories.is_empty() && series.category_reference.is_none() && !series.text_category_gaps { String::new() } else {
        let is_numeric = !series.text_category_gaps && categories.iter().all(|v| matches!(v, ChartCategoryValue::Number(_) | ChartCategoryValue::Blank));
        let body = if is_numeric {
            let values: Vec<_> = categories.iter().map(|v| match v {
                ChartCategoryValue::Number(n) => Some(*n), _ => None,
            }).collect();
            numeric_series_data(&values, series.category_reference.as_deref())
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
    let vals = format!("<c:{tag} xmlns:c=\"{ns}\">{}</c:{tag}>", numeric_series_data(&series.values, series.values_reference.as_deref()));
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
        if matches!(kind, "pie" | "doughnut" | "ofPie") {
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
        if !series.category_values_filtered || scatter { out.push_str(&cats); }
        out.push_str(&values);
        if kind == "bubble" {
            out.push_str(&bubble_series_data(series, C));
            out.push_str(&format!("<c:bubble3D val=\"{}\"/>", u8::from(series.shaded_bubbles)));
        }
        if matches!(kind, "line" | "scatter") { out.push_str(&format!("<c:smooth val=\"{}\"/>", u8::from(options.smooth))); }
        if series.category_values_filtered && !scatter && !cats.is_empty() {
            out.push_str(&format!("<c:extLst><c:ext uri=\"{{02D57815-91ED-43cb-92C2-25804820EDAC}}\">{}</c:ext></c:extLst>",
                filtered_category_title(&cats, C)));
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
        "bubble" => ("bubbleChart", "<c:varyColors val=\"0\"/>".into(), true),
        "pie" => (if options.three_dimensional { "pie3DChart" } else { "pieChart" }, "<c:varyColors val=\"1\"/>".into(), false),
        "doughnut" => ("doughnutChart", "<c:varyColors val=\"1\"/>".into(), false),
        "ofPie" => ("ofPieChart", format!("<c:ofPieType val=\"{}\"/><c:varyColors val=\"1\"/>",
            if options.of_pie.as_ref().is_some_and(|p| p.secondary == crate::ir::ChartSecondaryPlot::Bar) { "bar" } else { "pie" }), false),
        "radar" => ("radarChart", format!("<c:radarStyle val=\"{}\"/><c:varyColors val=\"0\"/>",
            if options.filled { "filled" } else { "marker" }), true),
        kind => return Err(XlsxError::InvalidData(format!("new chart plot kind {kind:?} cannot be serialized"))),
    };
    if !chart.shapes.is_empty() {
        return Err(XlsxError::InvalidData("new chart annotations need their own drawing part".into()));
    }
    let finite_frame = |frame: Option<&Frame>| frame.is_none_or(|f| [f.x, f.y, f.w, f.h].iter().all(|v| v.is_finite()));
    if options.of_pie.as_ref().and_then(|p| p.split_position).is_some_and(|p| !p.is_finite())
        || !finite_frame(chart.plot.as_ref()) || !finite_frame(chart.legend.as_ref().and_then(|l| l.frame.as_ref()))
        || chart.series.iter().any(|s| s.values.iter().flatten().any(|v| !v.is_finite())
            || s.bubble_sizes.iter().flatten().flatten().any(|v| !v.is_finite())
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
    for (index, value) in chart.series.iter().enumerate() { out.push_str(&series(value, index, &chart.categories, matches!(chart.kind.as_str(), "scatter" | "bubble"), options, &chart.kind)); }
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
    if chart.kind == "ofPie" {
        use crate::ir::ChartPieSplitMode;
        if let Some(pie) = &options.of_pie {
            let split = match pie.split_mode { ChartPieSplitMode::Automatic => None,
                ChartPieSplitMode::Position => Some("pos"), ChartPieSplitMode::Value => Some("val"),
                ChartPieSplitMode::Percent => Some("percent"), ChartPieSplitMode::Custom => Some("cust") };
            if let Some(split) = split { out.push_str(&format!("<c:splitType val=\"{split}\"/>")); }
            if let Some(position) = pie.split_position { out.push_str(&format!("<c:splitPos val=\"{position}\"/>")); }
            if pie.split_mode == ChartPieSplitMode::Custom {
                out.push_str("<c:custSplit>");
                for index in &pie.custom_points { out.push_str(&format!("<c:secondPiePt val=\"{index}\"/>")); }
                out.push_str("</c:custSplit>");
            }
            if let Some(size) = pie.secondary_size { out.push_str(&format!("<c:secondPieSize val=\"{size}\"/>")); }
        }
    }
    let depth_axis = options.three_dimensional && (options.depth_series || chart.kind == "line"
        || (chart.kind == "area" && options.stacking == ChartStacking::None));
    if axes {
        out.push_str("<c:axId val=\"1\"/><c:axId val=\"2\"/>");
        if depth_axis { out.push_str("<c:axId val=\"3\"/>"); }
    }
    out.push_str(&format!("</c:{plot}>"));
    if axes {
        out.push_str(&axis(chart.category_axis.as_ref(), if matches!(chart.kind.as_str(), "scatter" | "bubble") { "valAx" } else { "catAx" }, 1, 2, "b"));
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



fn ordered_source_series<'a>(xml: &str, fields: &'a [(String, usize, usize)]) -> Vec<&'a (String, usize, usize)> {
    let orders = crate::chart::series_positions(xml, "order");
    let mut series: Vec<_> = fields.iter().filter(|(name, _, _)| name == "ser").enumerate().collect();
    series.sort_by_key(|(index, _)| orders.get(*index).copied().flatten().unwrap_or(*index));
    series.into_iter().map(|(_, fragment)| fragment).collect()
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

/// Category expressions retained for formulas but excluded from the plotted axis.
fn filtered_category_title(categories: &str, ns: &str) -> String {
    if categories.is_empty() { return String::new(); }
    let body = &categories[categories.find('>').unwrap() + 1..categories.rfind("</").unwrap()];
    format!("<c15:filteredCategoryTitle xmlns:c15=\"http://schemas.microsoft.com/office/drawing/2012/chart\" xmlns:c=\"{ns}\"><c15:cat>{body}</c15:cat></c15:filteredCategoryTitle>")
}

fn patch_filtered_category_title(fragment: &str, categories: &str, open: &str, close: &str, ns: &str) -> Result<String, XlsxError> {
    use quick_xml::{name::ResolveResult, reader::NsReader};
    let wrapped = format!("{open}{fragment}{close}");
    let mut reader = NsReader::from_str(&wrapped);
    let mut parents: Vec<(String, bool)> = Vec::new();
    let mut start = None;
    let mut range = None;
    loop {
        let before = reader.buffer_position() as usize;
        let (namespace, event) = reader.read_resolved_event().map_err(XlsxError::Xml)?;
        let native = matches!(&namespace, ResolveResult::Bound(ns) if matches!(ns.as_ref(),
            b"http://schemas.openxmlformats.org/drawingml/2006/chart" | b"http://purl.oclc.org/ooxml/drawingml/chart"));
        let extension = matches!(&namespace, ResolveResult::Bound(ns) if ns.as_ref() == b"http://schemas.microsoft.com/office/drawing/2012/chart");
        let at_extension = parents.len() == 4 && parents.iter().all(|(_, native)| *native)
            && parents.iter().map(|(name, _)| name.as_str()).eq(["chartSpace", "ser", "extLst", "ext"]);
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                let name = local_name(element.name().as_ref());
                if extension && name == "filteredCategoryTitle" && at_extension {
                    if matches!(event, Event::Empty(_)) {
                        range = Some((before, reader.buffer_position() as usize));
                    } else { start = Some((before, parents.len() + 1)); }
                }
                if matches!(event, Event::Start(_)) { parents.push((name, native)); }
            }
            Event::End(_) => {
                if let Some((begin, depth)) = start {
                    if depth == parents.len() { range = Some((begin, reader.buffer_position() as usize)); start = None; }
                }
                parents.pop();
            }
            Event::Eof => break,
            _ => {}
        }
    }
    let title = filtered_category_title(categories, ns);
    let revised = if let Some((begin, end)) = range {
        format!("{}{}{}", &wrapped[..begin], title, &wrapped[end..])
    } else if title.is_empty() { wrapped.clone() }
    else {
        let extension = format!("<c:ext xmlns:c=\"{ns}\" uri=\"{{02D57815-91ED-43cb-92C2-25804820EDAC}}\">{title}</c:ext>");
        if let Ok((_, end)) = children_at(&wrapped, &["chartSpace", "ser", "extLst"]) {
            format!("{}{}{}", &wrapped[..end], extension, &wrapped[end..])
        } else {
            let (_, end) = children_at(&wrapped, &["chartSpace", "ser"])?;
            format!("{}<c:extLst xmlns:c=\"{ns}\">{extension}</c:extLst>{}", &wrapped[..end], &wrapped[end..])
        }
    };
    Ok(revised[open.len()..revised.len() - close.len()].to_string())
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
    let existing = ordered_source_series(xml, &children);
    let scatter = matches!(plot, "scatterChart" | "bubbleChart");
    let replace = |fragment: &str, tag: &str, replacement: &str, open: &str, close: &str| -> Result<String, XlsxError> {
        let wrapped = format!("{open}{fragment}{close}");
        let (children, end) = children_at(&wrapped, &["chartSpace", "ser"])?;
        let (start, finish) = if let Some((_, start, finish)) = children.iter().find(|(name, _, _)| name == tag) {
            (*start, *finish)
        } else {
            if replacement.is_empty() { return Ok(fragment.to_string()); }
            let at = children.iter().find(|(name, _, _)| match tag {
                "tx" => !matches!(name.as_str(), "idx" | "order"),
                "cat" | "xVal" => matches!(name.as_str(), "val" | "yVal" | "bubbleSize" | "bubble3D" | "smooth" | "extLst"),
                "val" | "yVal" => matches!(name.as_str(), "bubbleSize" | "bubble3D" | "smooth" | "extLst"),
                "bubbleSize" => matches!(name.as_str(), "bubble3D" | "extLst"),
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
            let name_changed = previous.is_none_or(|old| old.name != series.name || old.explicit_empty_name != series.explicit_empty_name || old.name_reference != series.name_reference);
            let cats_changed = previous.is_none_or(|old|
                old.category_reference != series.category_reference || old.category_values != series.category_values
                || old.text_category_gaps != series.text_category_gaps
                || old.category_values_filtered != series.category_values_filtered
                || ((old.category_values.is_none() || series.category_values.is_none()) && before.categories != after.categories));
            let values_changed = previous.is_none_or(|old| old.values != series.values || old.values_reference != series.values_reference);
            let mut fragment = xml[*start..*finish].to_string();
            if name_changed { fragment = replace(&fragment, "tx", &name, &open, &close)?; }
            if cats_changed {
                if series.category_values_filtered && !scatter {
                    fragment = patch_filtered_category_title(&fragment, &cats, &open, &close, ns)?;
                    fragment = replace(&fragment, "cat", "", &open, &close)?;
                } else {
                    fragment = replace(&fragment, if scatter { "xVal" } else { "cat" }, &cats, &open, &close)?;
                    if previous.is_some_and(|old| old.category_values_filtered) {
                        fragment = patch_filtered_category_title(&fragment, "", &open, &close, ns)?;
                    }
                }
            }
            if values_changed { fragment = replace(&fragment, if scatter { "yVal" } else { "val" }, &values, &open, &close)?; }
            if plot == "bubbleChart" {
                if previous.is_none_or(|old| old.bubble_sizes != series.bubble_sizes || old.bubble_sizes_reference != series.bubble_sizes_reference) {
                    fragment = replace(&fragment, "bubbleSize", &bubble_series_data(series, ns), &open, &close)?;
                }
                if previous.is_none_or(|old| old.shaded_bubbles != series.shaded_bubbles) {
                    let shading = format!("<c:bubble3D xmlns:c=\"{ns}\" val=\"{}\"/>", u8::from(series.shaded_bubbles));
                    fragment = replace(&fragment, "bubble3D", &shading, &open, &close)?;
                }
            }
            fragment
        } else { String::new() };
        edits.push((*start, *finish, replacement));
    }
    let mut appended = String::new();
    let next_position = |field| match crate::chart::series_positions(xml, field).into_iter().enumerate()
        .map(|(index, value)| value.unwrap_or(index)).max() {
            Some(value) => value.checked_add(1), None => Some(0),
        };
    let next_order = next_position("order");
    let next_id = next_position("idx");
    for (i, series) in after.series.iter().enumerate().skip(existing.len()) {
        let next = |position: Option<usize>| position.and_then(|value| value.checked_add(i - existing.len()))
            .ok_or_else(|| XlsxError::InvalidData("series identifiers cannot be extended".into()));
        let new_order = next(next_order)?;
        let new_id = next(next_id)?;
        // Reuse the new-series serializer for modeled formatting, then insert
        // only the series, retaining the original plot, axes and extensions.
        let kind = if plot == "bubbleChart" { "bubble" } else if scatter { "scatter" } else if matches!(before.kind.as_str(), "pie" | "doughnut" | "ofPie") { &before.kind } else { "bar" };
        let generated = new_chart_xml(&Chart {
            kind: kind.into(), categories: after.categories.clone(),
            plot_options: after.plot_options.clone(), series: vec![series.clone()], ..Default::default()
        })?;
        let tag = if kind == "bubble" { "bubbleChart" } else if scatter { "scatterChart" } else if kind == "pie" { if after.plot_options.three_dimensional { "pie3DChart" } else { "pieChart" } }
            else if kind == "doughnut" { "doughnutChart" } else if kind == "ofPie" { "ofPieChart" } else if after.plot_options.three_dimensional { "bar3DChart" } else { "barChart" };
        let (ranges, _) = children_at(&generated, &["chartSpace", "chart", "plotArea", tag])?;
        let (_, begin, finish) = ranges.iter().find(|(name, _, _)| name == "ser").unwrap();
        let fragment = generated[*begin..*finish]
            .replacen("<c:ser>", &format!("<c:ser xmlns:c=\"{ns}\" xmlns:a=\"{drawing_ns}\">"), 1)
            .replace(&format!("xmlns:c=\"{C}\""), &format!("xmlns:c=\"{ns}\""))
            .replacen("<c:idx val=\"0\"/>", &format!("<c:idx val=\"{new_id}\"/>"), 1)
            .replacen("<c:order val=\"0\"/>", &format!("<c:order val=\"{new_order}\"/>"), 1);
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
    let old_series = ordered_source_series(xml, &old_settings);
    let new_series: Vec<_> = new_settings.iter().filter(|(name, _, _)| name == "ser").collect();
    let (new_series_fields, _) = children_at(&generated, &["chartSpace", "chart", "plotArea", new_tag, "ser"])
        .unwrap_or_default();
    for (i, (_, begin, end)) in new_series.iter().enumerate().rev() {
        if let Some((_, old_begin, old_end)) = old_series.get(i) {
            let (open, close) = series_fragment_envelope(xml, *old_begin, ns)?;
            let fragment = &xml[*old_begin..*old_end];
            let (fields, _) = children_at(&format!("{open}{fragment}{close}"), &["chartSpace", "ser"])?;
            let mut additions: String = fields.iter().filter(|(name, _, _)| name == "extLst")
                .map(|(_, a, b)| &fragment[a - open.len()..b - open.len()]).collect();
            if !additions.is_empty() {
                let series = &after.series[i];
                let scatter = matches!(after.kind.as_str(), "scatter" | "bubble");
                let (_, categories, _) = series_data_parts(series, &after.categories, scatter, ns);
                let categories = if series.category_values_filtered && !scatter { categories.as_str() } else { "" };
                let revised = patch_filtered_category_title(fragment, categories, &open, &close, ns)?;
                let wrapped = format!("{open}{revised}{close}");
                let (fields, _) = children_at(&wrapped, &["chartSpace", "ser"])?;
                additions = fields.iter().filter(|(name, _, _)| name == "extLst")
                    .map(|(_, a, b)| &wrapped[*a..*b]).collect();
                if let Some((_, a, b)) = new_series_fields.iter().find(|(name, a, b)|
                    name == "extLst" && a >= begin && b <= end) {
                    plot.replace_range(a - new_start..b - new_start, &additions);
                } else {
                    let at = end - new_start - "</c:ser>".len();
                    plot.insert_str(at, &additions);
                }
            }
        }
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
            && !matches!(name.as_str(), "ser" | "axId" | "barDir" | "grouping" | "scatterStyle" | "radarStyle" | "marker" | "smooth" | "overlap" | "varyColors" | "shape" | "ofPieType" | "splitType" | "splitPos" | "custSplit" | "secondPieSize"));
        if !retain { continue; }
        // These source properties precede the plot's axes; extension lists are last.
        let successors: &[&str] = match name.as_str() {
            "dLbls" => &["dropLines", "hiLowLines", "upDownBars", "gapWidth", "gapDepth", "overlap", "shape", "serLines", "marker", "smooth", "firstSliceAng", "holeSize", "splitType", "splitPos", "custSplit", "secondPieSize", "axId", "extLst"],
            "gapWidth" => &["gapDepth", "overlap", "splitType", "splitPos", "custSplit", "secondPieSize", "serLines", "shape", "axId", "extLst"],
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
mod filtered_category_source_tests {
    use super::*;
    use crate::{chart::parse_chart_xml, ir::ChartCategoryValue, parser::Theme};

    #[test]
    fn mixed_plot_series_keep_each_plots_numeric_sources() {
        let categories = "<c:xVal><c:numLit><c:ptCount val=\"1\"/><c:pt idx=\"0\"><c:v>33</c:v></c:pt></c:numLit></c:xVal>";
        let line = "<c:lineChart><c:ser><c:val><c:numLit><c:pt idx=\"0\"><c:v>11</c:v></c:pt></c:numLit></c:val></c:ser></c:lineChart>";
        let scatter = format!("<c:scatterChart><c:ser>{categories}<c:yVal><c:numLit><c:pt idx=\"0\"><c:v>22</c:v></c:pt></c:numLit></c:yVal></c:ser></c:scatterChart>");
        for reversed in [false, true] {
            let plots = if reversed { format!("{scatter}{line}") } else { format!("{line}{scatter}") };
            let xml = format!("<c:chartSpace xmlns:c=\"{C}\"><c:chart><c:plotArea>{plots}</c:plotArea></c:chart></c:chartSpace>");
            let chart = parse_chart_xml(&xml, &Theme::default()).unwrap();
            let (line, scatter) = if reversed { (&chart.series[1], &chart.series[0]) }
                else { (&chart.series[0], &chart.series[1]) };
            assert_eq!(line.values, [Some(11.0)]);
            assert_eq!(scatter.values, [Some(22.0)]);
            assert_eq!(scatter.category_values, Some(vec![ChartCategoryValue::Number(33.0)]));
        }
    }

    #[test]
    fn filtered_category_sources_survive_creation_and_edits_without_becoming_axis_labels() {
        let controls: serde_json::Value = serde_json::from_str(include_str!(
            "../../../tests/fixtures/vba/chart_import_contracts.json")).unwrap();
        let mut checked = 0;
        for case in controls["category_cases"].as_array().unwrap() {
            if !case["id"].as_str().unwrap().starts_with("filtered:")
                && !case["id"].as_str().unwrap().starts_with("native_saved:") { continue; }
            let source = case["xml"].as_str().unwrap();
            let chart = parse_chart_xml(source, &Theme::default()).unwrap();
            if !chart.series.iter().any(|series| series.category_values_filtered) { continue; }
            checked += 1;
            for strict in [false, true] {
                let xml = if strict {
                    source.replace(C, "http://purl.oclc.org/ooxml/drawingml/chart")
                        .replace(A, "http://purl.oclc.org/ooxml/drawingml/main")
                } else { source.to_string() };
                let before = parse_chart_xml(&xml, &Theme::default()).unwrap();
                assert!(before.categories.is_empty());
                let created = parse_chart_xml(&new_chart_xml(&before).unwrap(), &Theme::default()).unwrap();
                for (expected, actual) in before.series.iter().zip(&created.series) {
                    assert_eq!(actual.category_values, expected.category_values, "{}", case["id"]);
                    assert!(actual.category_values_filtered);
                }
                let mut after = before.clone();
                after.series[0].values[0] = Some(91.0);
                after.series[0].category_values = Some(vec![ChartCategoryValue::Text("Edited".into())]);
                after.series[0].text_category_gaps = false;
                let patched = patch_chart(&xml, &before, &after).unwrap();
                let read = parse_chart_xml(&patched, &Theme::default()).unwrap();
                assert_eq!(read.series[0].category_values, after.series[0].category_values);
                assert!(read.series[0].category_values_filtered);
                assert_eq!(read.series[0].values[0], Some(91.0));
                assert!(read.categories.is_empty());
                assert_eq!(read.series.get(1), before.series.get(1));
                assert_eq!(patched.matches("uniqueId").count(), xml.matches("uniqueId").count());
                let mut options = after.clone();
                options.plot_options.markers = Some(!after.plot_options.markers.unwrap_or(true));
                let changed_plot = patch_chart(&patched, &after, &options).unwrap();
                assert_eq!(changed_plot.matches(":extLst>").count(), patched.matches(":extLst>").count());
                let read_plot = parse_chart_xml(&changed_plot, &Theme::default()).unwrap();
                assert!(read_plot.categories.is_empty());
                assert!(read_plot.series[0].category_values_filtered);
                assert_eq!(read_plot.series[0].category_values, options.series[0].category_values);
                assert_eq!(read_plot.series[0].values, options.series[0].values);
                let mut explicit = after.clone();
                explicit.series[0].category_values_filtered = false;
                let reactivated = patch_chart(&patched, &after, &explicit).unwrap();
                let read = parse_chart_xml(&reactivated, &Theme::default()).unwrap();
                assert!(!read.series[0].category_values_filtered);
                assert_eq!(read.series[0].category_values, explicit.series[0].category_values);
                assert_eq!(read.series.get(1), before.series.get(1));
            }
        }
        assert_eq!(checked, 10);
    }
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


#[cfg(test)]
mod bubble_data_tests {
    use super::*;
    use crate::ir::ChartSeries;

    fn source() -> String {
        format!(r#"<c:chartSpace xmlns:c="{C}" xmlns:a="{A}"><c:chart><c:plotArea><c:bubbleChart>
<c:varyColors val="0"/><c:ser><c:idx val="0"/><c:order val="0"/>
<c:tx><c:v>Revenue</c:v></c:tx>
<c:xVal><c:numLit><c:ptCount val="3"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt></c:numLit></c:xVal>
<c:yVal><c:numLit><c:ptCount val="3"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt><c:pt idx="2"><c:v>30</c:v></c:pt></c:numLit></c:yVal>
<c:bubbleSize><c:numRef><c:f>'Sales &amp; Costs'!$C$1:$C$4</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="4"/><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>4</c:v></c:pt></c:numCache></c:numRef></c:bubbleSize>
<c:bubble3D val="1"/><c:extLst><c:ext uri="keep"><x:keep xmlns:x="urn:extra">unchanged</x:keep></c:ext></c:extLst></c:ser>
<c:ser><c:idx val="1"/><c:order val="1"/><c:tx><c:v>Other</c:v></c:tx><c:yVal><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>99</c:v></c:pt></c:numLit></c:yVal></c:ser>
<c:bubbleScale val="125"/><c:showNegBubbles val="1"/><c:axId val="1"/><c:axId val="2"/></c:bubbleChart></c:plotArea></c:chart></c:chartSpace>"#)
    }

    #[test]
    fn parsing_keeps_bubble_sizes_separate_from_values_and_preserves_sparse_length() {
        for strict in [false, true] {
            let mut xml = source();
            if strict { xml = xml.replace(C, "http://purl.oclc.org/ooxml/drawingml/chart")
                .replace(A, "http://purl.oclc.org/ooxml/drawingml/main"); }
            let chart = crate::chart::parse_chart_xml(&xml, &crate::parser::Theme::default()).unwrap();
            assert_eq!(chart.kind, "bubble");
            assert_eq!(chart.series[0].values, vec![Some(10.0), Some(20.0), Some(30.0)]);
            assert_eq!(chart.series[1].values, vec![Some(99.0)]);
            assert_eq!(chart.series[0].bubble_sizes, Some(vec![Some(2.0), None, Some(4.0), None]));
            assert_eq!(chart.series[0].bubble_sizes_reference.as_deref(), Some("'Sales & Costs'!$C$1:$C$4"));
            assert!(chart.series[0].shaded_bubbles);
            assert!(chart.series[1].bubble_sizes.is_none());
            assert!(!chart.plot_options.three_dimensional);
        }
    }

    #[test]
    fn serializing_bubble_sizes_uses_two_numeric_axes_without_a_spatial_view() {
        for shaded in [false, true] {
            for reference in [None, Some("Sheet1!$C$1:$C$3")] {
                let series = ChartSeries {
                    name:"Revenue".into(), values:vec![Some(10.0), Some(20.0), Some(30.0)],
                    category_values:Some((1..=3).map(|n| crate::ir::ChartCategoryValue::Number(n as f64)).collect()),
                    bubble_sizes:Some(vec![Some(2.0), None, Some(4.0)]),
                    bubble_sizes_reference:reference.map(str::to_string), shaded_bubbles:shaded,
                    ..Default::default()
                };
                let chart = Chart { kind:"bubble".into(), series:vec![series], ..Default::default() };
                let xml = new_chart_xml(&chart).unwrap();
                assert!(xml.contains("<c:bubbleChart>") && xml.contains("<c:bubbleSize"));
                assert_eq!(xml.matches("<c:valAx>").count(), 2);
                assert!(!xml.contains("<c:catAx>") && !xml.contains("<c:view3D>"));
                let restored = crate::chart::parse_chart_xml(&xml, &crate::parser::Theme::default()).unwrap();
                assert!(series_data_equal(&chart.series[0], &restored.series[0]));
                assert_eq!(restored.series[0].shaded_bubbles, shaded);
            }
        }
    }

    #[test]
    fn editing_bubble_sizes_keeps_other_series_and_unmodeled_xml() {
        for strict in [false, true] {
            let mut xml = source();
            if strict { xml = xml.replace(C, "http://purl.oclc.org/ooxml/drawingml/chart")
                .replace(A, "http://purl.oclc.org/ooxml/drawingml/main"); }
            let before = crate::chart::parse_chart_xml(&xml, &crate::parser::Theme::default()).unwrap();
            let mut after = before.clone();
            after.series[0].bubble_sizes = Some(vec![Some(5.0), Some(6.0), Some(7.0)]);
            after.series[0].bubble_sizes_reference = None;
            assert!(series_data_changed(&before, &after));
            let patched = patch_chart(&xml, &before, &after).unwrap();
            assert!(patched.contains("<x:keep xmlns:x=\"urn:extra\">unchanged</x:keep>"));
            assert!(patched.contains("<c:bubbleScale val=\"125\"/>"));
            let ranges = |text: &str| children_at(text, &["chartSpace", "chart", "plotArea", "bubbleChart"]).unwrap().0
                .into_iter().filter(|(name, _, _)| name == "ser").map(|(_, a, b)| text[a..b].to_string()).collect::<Vec<_>>();
            assert_eq!(ranges(&xml)[1], ranges(&patched)[1]);
            let actual = crate::chart::parse_chart_xml(&patched, &crate::parser::Theme::default()).unwrap();
            assert!(series_data_equal(&after.series[0], &actual.series[0]));
            assert_eq!(actual.series[0].values, before.series[0].values);
        }
    }
}


#[cfg(test)]
mod ofpie_data_tests {
    use super::*;
    use crate::ir::{ChartCategoryValue, ChartSeries};

    #[test]
    fn new_ofpie_parts_keep_underlying_caches_without_virtual_aggregate_points() {
        let expected = Chart { kind: "ofPie".into(),
            series: vec![ChartSeries { name: "Revenue".into(),
                category_values: Some((1..=3).map(|i| ChartCategoryValue::Text(i.to_string())).collect()),
                values: vec![Some(1.25), Some(2.5), Some(3.75)], ..Default::default() }],
            ..Default::default() };
        let xml = new_chart_xml(&expected).unwrap();
        assert!(xml.contains("<c:ofPieChart>") && xml.contains("<c:ofPieType val=\"pie\"/>"));
        assert!(!xml.contains("<c:valAx>") && !xml.contains("<c:catAx>") && !xml.contains("<c:view3D>"));
        let restored = crate::chart::parse_chart_xml(&xml, &crate::parser::Theme::default()).unwrap();
        assert_eq!(restored.kind, "ofPie");
        assert_eq!(restored.series[0].values, expected.series[0].values);
        assert_eq!(restored.series[0].category_values, expected.series[0].category_values);
        assert_eq!(restored.series[0].values.len(), 3, "the aggregate is an API view, not source data");
    }

    #[test]
    fn actual_excel_ofpie_parts_keep_split_state_and_source_data() {
        use crate::ir::{ChartOfPieOptions, ChartPieSplitMode as Mode, ChartSecondaryPlot as Secondary};
        // Self-authored controls saved by Excel, including rejected assignments.
        let controls = [
            (r###"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c16r2="http://schemas.microsoft.com/office/drawing/2015/06/chart"><c:date1904 val="0"/><c:lang val="ja-JP"/><c:roundedCorners val="0"/><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="c14" xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart"><c14:style val="102"/></mc:Choice><mc:Fallback><c:style val="2"/></mc:Fallback></mc:AlternateContent><c:chart><c:autoTitleDeleted val="0"/><c:plotArea><c:layout/><c:ofPieChart><c:ofPieType val="pie"/><c:varyColors val="1"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Revenue</c:v></c:tx><c:cat><c:strLit><c:ptCount val="3"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="3"/><c:pt idx="0"><c:v>1.25</c:v></c:pt><c:pt idx="1"><c:v>2.5</c:v></c:pt><c:pt idx="2"><c:v>3.75</c:v></c:pt></c:numLit></c:val><c:extLst><c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}" xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart"><c16:uniqueId val="{00000000-E4E1-4C27-AF5B-754229312CEE}"/></c:ext></c:extLst></c:ser><c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/><c:showLeaderLines val="1"/></c:dLbls><c:gapWidth val="150"/><c:secondPieSize val="75"/><c:serLines/></c:ofPieChart></c:plotArea><c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/><c:showDLblsOverMax val="0"/><c:extLst><c:ext uri="{56B9EC1D-385E-4148-901F-78D8002777C0}" xmlns:c16r3="http://schemas.microsoft.com/office/drawing/2017/03/chart"><c16r3:dataDisplayOptions16><c16r3:dispNaAsBlank val="1"/></c16r3:dataDisplayOptions16></c:ext></c:extLst></c:chart></c:chartSpace>"###, ChartOfPieOptions { secondary: Secondary::Pie, split_mode: Mode::Automatic, split_position: None, custom_points: vec![], secondary_size: Some(75) }, vec![Some(1.25),Some(2.5),Some(3.75)]),
            (r###"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c16r2="http://schemas.microsoft.com/office/drawing/2015/06/chart"><c:date1904 val="0"/><c:lang val="ja-JP"/><c:roundedCorners val="0"/><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="c14" xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart"><c14:style val="102"/></mc:Choice><mc:Fallback><c:style val="2"/></mc:Fallback></mc:AlternateContent><c:chart><c:autoTitleDeleted val="0"/><c:plotArea><c:layout/><c:ofPieChart><c:ofPieType val="pie"/><c:varyColors val="1"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Revenue</c:v></c:tx><c:cat><c:strLit><c:ptCount val="5"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt><c:pt idx="3"><c:v>4</c:v></c:pt><c:pt idx="4"><c:v>5</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="5"/><c:pt idx="0"><c:v>-5</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>-4</c:v></c:pt><c:pt idx="3"><c:v>2</c:v></c:pt><c:pt idx="4"><c:v>-3</c:v></c:pt></c:numLit></c:val><c:extLst><c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}" xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart"><c16:uniqueId val="{00000000-836C-44F0-9F5A-1A4F6F691754}"/></c:ext></c:extLst></c:ser><c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/><c:showLeaderLines val="1"/></c:dLbls><c:gapWidth val="150"/><c:secondPieSize val="75"/><c:serLines/></c:ofPieChart></c:plotArea><c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/><c:showDLblsOverMax val="0"/><c:extLst><c:ext uri="{56B9EC1D-385E-4148-901F-78D8002777C0}" xmlns:c16r3="http://schemas.microsoft.com/office/drawing/2017/03/chart"><c16r3:dataDisplayOptions16><c16r3:dispNaAsBlank val="1"/></c16r3:dataDisplayOptions16></c:ext></c:extLst></c:chart></c:chartSpace>"###, ChartOfPieOptions { secondary: Secondary::Pie, split_mode: Mode::Automatic, split_position: None, custom_points: vec![], secondary_size: Some(75) }, vec![Some(-5.0),Some(1.0),Some(-4.0),Some(2.0),Some(-3.0)]),
            (r###"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c16r2="http://schemas.microsoft.com/office/drawing/2015/06/chart"><c:date1904 val="0"/><c:lang val="ja-JP"/><c:roundedCorners val="0"/><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="c14" xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart"><c14:style val="102"/></mc:Choice><mc:Fallback><c:style val="2"/></mc:Fallback></mc:AlternateContent><c:chart><c:autoTitleDeleted val="0"/><c:plotArea><c:layout/><c:ofPieChart><c:ofPieType val="pie"/><c:varyColors val="1"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Revenue</c:v></c:tx><c:cat><c:strLit><c:ptCount val="8"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt><c:pt idx="3"><c:v>4</c:v></c:pt><c:pt idx="4"><c:v>5</c:v></c:pt><c:pt idx="5"><c:v>6</c:v></c:pt><c:pt idx="6"><c:v>7</c:v></c:pt><c:pt idx="7"><c:v>8</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="8"/><c:pt idx="0"><c:v>8</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>7</c:v></c:pt><c:pt idx="3"><c:v>2</c:v></c:pt><c:pt idx="4"><c:v>6</c:v></c:pt><c:pt idx="5"><c:v>3</c:v></c:pt><c:pt idx="6"><c:v>5</c:v></c:pt><c:pt idx="7"><c:v>4</c:v></c:pt></c:numLit></c:val><c:extLst><c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}" xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart"><c16:uniqueId val="{00000000-7EB6-406E-A12F-A5A45495861B}"/></c:ext></c:extLst></c:ser><c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/><c:showLeaderLines val="1"/></c:dLbls><c:gapWidth val="150"/><c:splitType val="pos"/><c:splitPos val="3"/><c:secondPieSize val="75"/><c:serLines/></c:ofPieChart></c:plotArea><c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/><c:showDLblsOverMax val="0"/><c:extLst><c:ext uri="{56B9EC1D-385E-4148-901F-78D8002777C0}" xmlns:c16r3="http://schemas.microsoft.com/office/drawing/2017/03/chart"><c16r3:dataDisplayOptions16><c16r3:dispNaAsBlank val="1"/></c16r3:dataDisplayOptions16></c:ext></c:extLst></c:chart></c:chartSpace>"###, ChartOfPieOptions { secondary: Secondary::Pie, split_mode: Mode::Position, split_position: Some(3.0), custom_points: vec![], secondary_size: Some(75) }, vec![Some(8.0),Some(1.0),Some(7.0),Some(2.0),Some(6.0),Some(3.0),Some(5.0),Some(4.0)]),
            (r###"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c16r2="http://schemas.microsoft.com/office/drawing/2015/06/chart"><c:date1904 val="0"/><c:lang val="ja-JP"/><c:roundedCorners val="0"/><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="c14" xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart"><c14:style val="102"/></mc:Choice><mc:Fallback><c:style val="2"/></mc:Fallback></mc:AlternateContent><c:chart><c:autoTitleDeleted val="0"/><c:plotArea><c:layout/><c:ofPieChart><c:ofPieType val="pie"/><c:varyColors val="1"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Revenue</c:v></c:tx><c:cat><c:strLit><c:ptCount val="8"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt><c:pt idx="3"><c:v>4</c:v></c:pt><c:pt idx="4"><c:v>5</c:v></c:pt><c:pt idx="5"><c:v>6</c:v></c:pt><c:pt idx="6"><c:v>7</c:v></c:pt><c:pt idx="7"><c:v>8</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="8"/><c:pt idx="0"><c:v>8</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>7</c:v></c:pt><c:pt idx="3"><c:v>2</c:v></c:pt><c:pt idx="4"><c:v>6</c:v></c:pt><c:pt idx="5"><c:v>3</c:v></c:pt><c:pt idx="6"><c:v>5</c:v></c:pt><c:pt idx="7"><c:v>4</c:v></c:pt></c:numLit></c:val><c:extLst><c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}" xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart"><c16:uniqueId val="{00000000-9B06-41AF-98DD-586F32EEFCF0}"/></c:ext></c:extLst></c:ser><c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/><c:showLeaderLines val="1"/></c:dLbls><c:gapWidth val="150"/><c:splitType val="val"/><c:splitPos val="3"/><c:secondPieSize val="75"/><c:serLines/></c:ofPieChart></c:plotArea><c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/><c:showDLblsOverMax val="0"/><c:extLst><c:ext uri="{56B9EC1D-385E-4148-901F-78D8002777C0}" xmlns:c16r3="http://schemas.microsoft.com/office/drawing/2017/03/chart"><c16r3:dataDisplayOptions16><c16r3:dispNaAsBlank val="1"/></c16r3:dataDisplayOptions16></c:ext></c:extLst></c:chart></c:chartSpace>"###, ChartOfPieOptions { secondary: Secondary::Pie, split_mode: Mode::Value, split_position: Some(3.0), custom_points: vec![], secondary_size: Some(75) }, vec![Some(8.0),Some(1.0),Some(7.0),Some(2.0),Some(6.0),Some(3.0),Some(5.0),Some(4.0)]),
            (r###"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c16r2="http://schemas.microsoft.com/office/drawing/2015/06/chart"><c:date1904 val="0"/><c:lang val="ja-JP"/><c:roundedCorners val="0"/><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="c14" xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart"><c14:style val="102"/></mc:Choice><mc:Fallback><c:style val="2"/></mc:Fallback></mc:AlternateContent><c:chart><c:autoTitleDeleted val="0"/><c:plotArea><c:layout/><c:ofPieChart><c:ofPieType val="pie"/><c:varyColors val="1"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Revenue</c:v></c:tx><c:cat><c:strLit><c:ptCount val="8"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt><c:pt idx="3"><c:v>4</c:v></c:pt><c:pt idx="4"><c:v>5</c:v></c:pt><c:pt idx="5"><c:v>6</c:v></c:pt><c:pt idx="6"><c:v>7</c:v></c:pt><c:pt idx="7"><c:v>8</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="8"/><c:pt idx="0"><c:v>8</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>7</c:v></c:pt><c:pt idx="3"><c:v>2</c:v></c:pt><c:pt idx="4"><c:v>6</c:v></c:pt><c:pt idx="5"><c:v>3</c:v></c:pt><c:pt idx="6"><c:v>5</c:v></c:pt><c:pt idx="7"><c:v>4</c:v></c:pt></c:numLit></c:val><c:extLst><c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}" xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart"><c16:uniqueId val="{00000000-8864-4797-ADC3-485930F484A9}"/></c:ext></c:extLst></c:ser><c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/><c:showLeaderLines val="1"/></c:dLbls><c:gapWidth val="150"/><c:splitType val="percent"/><c:splitPos val="20"/><c:secondPieSize val="75"/><c:serLines/></c:ofPieChart></c:plotArea><c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/><c:showDLblsOverMax val="0"/><c:extLst><c:ext uri="{56B9EC1D-385E-4148-901F-78D8002777C0}" xmlns:c16r3="http://schemas.microsoft.com/office/drawing/2017/03/chart"><c16r3:dataDisplayOptions16><c16r3:dispNaAsBlank val="1"/></c16r3:dataDisplayOptions16></c:ext></c:extLst></c:chart></c:chartSpace>"###, ChartOfPieOptions { secondary: Secondary::Pie, split_mode: Mode::Percent, split_position: Some(20.0), custom_points: vec![], secondary_size: Some(75) }, vec![Some(8.0),Some(1.0),Some(7.0),Some(2.0),Some(6.0),Some(3.0),Some(5.0),Some(4.0)]),
            (r###"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c16r2="http://schemas.microsoft.com/office/drawing/2015/06/chart"><c:date1904 val="0"/><c:lang val="ja-JP"/><c:roundedCorners val="0"/><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="c14" xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart"><c14:style val="102"/></mc:Choice><mc:Fallback><c:style val="2"/></mc:Fallback></mc:AlternateContent><c:chart><c:autoTitleDeleted val="0"/><c:plotArea><c:layout/><c:ofPieChart><c:ofPieType val="pie"/><c:varyColors val="1"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Revenue</c:v></c:tx><c:cat><c:strLit><c:ptCount val="8"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt><c:pt idx="3"><c:v>4</c:v></c:pt><c:pt idx="4"><c:v>5</c:v></c:pt><c:pt idx="5"><c:v>6</c:v></c:pt><c:pt idx="6"><c:v>7</c:v></c:pt><c:pt idx="7"><c:v>8</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="8"/><c:pt idx="0"><c:v>8</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>7</c:v></c:pt><c:pt idx="3"><c:v>2</c:v></c:pt><c:pt idx="4"><c:v>6</c:v></c:pt><c:pt idx="5"><c:v>3</c:v></c:pt><c:pt idx="6"><c:v>5</c:v></c:pt><c:pt idx="7"><c:v>4</c:v></c:pt></c:numLit></c:val><c:extLst><c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}" xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart"><c16:uniqueId val="{00000000-2A24-46E2-A7BF-489CC4E8BF8C}"/></c:ext></c:extLst></c:ser><c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/><c:showLeaderLines val="1"/></c:dLbls><c:gapWidth val="150"/><c:splitType val="cust"/><c:custSplit><c:secondPiePt val="5"/><c:secondPiePt val="6"/><c:secondPiePt val="7"/></c:custSplit><c:secondPieSize val="75"/><c:serLines/></c:ofPieChart></c:plotArea><c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/><c:showDLblsOverMax val="0"/><c:extLst><c:ext uri="{56B9EC1D-385E-4148-901F-78D8002777C0}" xmlns:c16r3="http://schemas.microsoft.com/office/drawing/2017/03/chart"><c16r3:dataDisplayOptions16><c16r3:dispNaAsBlank val="1"/></c16r3:dataDisplayOptions16></c:ext></c:extLst></c:chart></c:chartSpace>"###, ChartOfPieOptions { secondary: Secondary::Pie, split_mode: Mode::Custom, split_position: None, custom_points: vec![5,6,7], secondary_size: Some(75) }, vec![Some(8.0),Some(1.0),Some(7.0),Some(2.0),Some(6.0),Some(3.0),Some(5.0),Some(4.0)]),
            (r###"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c16r2="http://schemas.microsoft.com/office/drawing/2015/06/chart"><c:date1904 val="0"/><c:lang val="ja-JP"/><c:roundedCorners val="0"/><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="c14" xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart"><c14:style val="102"/></mc:Choice><mc:Fallback><c:style val="2"/></mc:Fallback></mc:AlternateContent><c:chart><c:autoTitleDeleted val="0"/><c:plotArea><c:layout/><c:ofPieChart><c:ofPieType val="pie"/><c:varyColors val="1"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Revenue</c:v></c:tx><c:cat><c:strLit><c:ptCount val="5"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt><c:pt idx="3"><c:v>4</c:v></c:pt><c:pt idx="4"><c:v>5</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="5"/><c:pt idx="0"><c:v>-5</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>-4</c:v></c:pt><c:pt idx="3"><c:v>2</c:v></c:pt><c:pt idx="4"><c:v>-3</c:v></c:pt></c:numLit></c:val><c:extLst><c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}" xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart"><c16:uniqueId val="{00000000-B667-4715-8E4E-0C6E40A011F5}"/></c:ext></c:extLst></c:ser><c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/><c:showLeaderLines val="1"/></c:dLbls><c:gapWidth val="150"/><c:secondPieSize val="75"/><c:serLines/></c:ofPieChart></c:plotArea><c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/><c:showDLblsOverMax val="0"/><c:extLst><c:ext uri="{56B9EC1D-385E-4148-901F-78D8002777C0}" xmlns:c16r3="http://schemas.microsoft.com/office/drawing/2017/03/chart"><c16r3:dataDisplayOptions16><c16r3:dispNaAsBlank val="1"/></c16r3:dataDisplayOptions16></c:ext></c:extLst></c:chart></c:chartSpace>"###, ChartOfPieOptions { secondary: Secondary::Pie, split_mode: Mode::Automatic, split_position: None, custom_points: vec![], secondary_size: Some(75) }, vec![Some(-5.0),Some(1.0),Some(-4.0),Some(2.0),Some(-3.0)]),
            (r###"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c16r2="http://schemas.microsoft.com/office/drawing/2015/06/chart"><c:date1904 val="0"/><c:lang val="ja-JP"/><c:roundedCorners val="0"/><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="c14" xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart"><c14:style val="102"/></mc:Choice><mc:Fallback><c:style val="2"/></mc:Fallback></mc:AlternateContent><c:chart><c:autoTitleDeleted val="0"/><c:plotArea><c:layout/><c:ofPieChart><c:ofPieType val="pie"/><c:varyColors val="1"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Revenue</c:v></c:tx><c:cat><c:strLit><c:ptCount val="5"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt><c:pt idx="3"><c:v>4</c:v></c:pt><c:pt idx="4"><c:v>5</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="5"/><c:pt idx="0"><c:v>-5</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>-4</c:v></c:pt><c:pt idx="3"><c:v>2</c:v></c:pt><c:pt idx="4"><c:v>-3</c:v></c:pt></c:numLit></c:val><c:extLst><c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}" xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart"><c16:uniqueId val="{00000000-B606-49AE-81C3-95BEDF6C9A40}"/></c:ext></c:extLst></c:ser><c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/><c:showLeaderLines val="1"/></c:dLbls><c:gapWidth val="150"/><c:splitType val="pos"/><c:splitPos val="2"/><c:secondPieSize val="75"/><c:serLines/></c:ofPieChart></c:plotArea><c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/><c:showDLblsOverMax val="0"/><c:extLst><c:ext uri="{56B9EC1D-385E-4148-901F-78D8002777C0}" xmlns:c16r3="http://schemas.microsoft.com/office/drawing/2017/03/chart"><c16r3:dataDisplayOptions16><c16r3:dispNaAsBlank val="1"/></c16r3:dataDisplayOptions16></c:ext></c:extLst></c:chart></c:chartSpace>"###, ChartOfPieOptions { secondary: Secondary::Pie, split_mode: Mode::Position, split_position: Some(2.0), custom_points: vec![], secondary_size: Some(75) }, vec![Some(-5.0),Some(1.0),Some(-4.0),Some(2.0),Some(-3.0)]),
            (r###"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c16r2="http://schemas.microsoft.com/office/drawing/2015/06/chart"><c:date1904 val="0"/><c:lang val="ja-JP"/><c:roundedCorners val="0"/><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="c14" xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart"><c14:style val="102"/></mc:Choice><mc:Fallback><c:style val="2"/></mc:Fallback></mc:AlternateContent><c:chart><c:autoTitleDeleted val="0"/><c:plotArea><c:layout/><c:ofPieChart><c:ofPieType val="bar"/><c:varyColors val="1"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Revenue</c:v></c:tx><c:cat><c:strLit><c:ptCount val="5"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt><c:pt idx="3"><c:v>4</c:v></c:pt><c:pt idx="4"><c:v>5</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="5"/><c:pt idx="0"><c:v>0</c:v></c:pt><c:pt idx="1"><c:v>0</c:v></c:pt><c:pt idx="2"><c:v>0</c:v></c:pt><c:pt idx="3"><c:v>0</c:v></c:pt><c:pt idx="4"><c:v>0</c:v></c:pt></c:numLit></c:val><c:extLst><c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}" xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart"><c16:uniqueId val="{00000000-C164-4AB4-B729-D3B9424D4BFB}"/></c:ext></c:extLst></c:ser><c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/><c:showLeaderLines val="1"/></c:dLbls><c:gapWidth val="150"/><c:secondPieSize val="75"/><c:serLines/></c:ofPieChart></c:plotArea><c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/><c:showDLblsOverMax val="0"/><c:extLst><c:ext uri="{56B9EC1D-385E-4148-901F-78D8002777C0}" xmlns:c16r3="http://schemas.microsoft.com/office/drawing/2017/03/chart"><c16r3:dataDisplayOptions16><c16r3:dispNaAsBlank val="1"/></c16r3:dataDisplayOptions16></c:ext></c:extLst></c:chart><c:printSettings><c:headerFooter/><c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/><c:pageSetup/></c:printSettings></c:chartSpace>"###, ChartOfPieOptions { secondary: Secondary::Bar, split_mode: Mode::Automatic, split_position: None, custom_points: vec![], secondary_size: Some(75) }, vec![Some(0.0),Some(0.0),Some(0.0),Some(0.0),Some(0.0)]),
            (r###"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c16r2="http://schemas.microsoft.com/office/drawing/2015/06/chart"><c:date1904 val="0"/><c:lang val="ja-JP"/><c:roundedCorners val="0"/><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="c14" xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart"><c14:style val="102"/></mc:Choice><mc:Fallback><c:style val="2"/></mc:Fallback></mc:AlternateContent><c:chart><c:autoTitleDeleted val="0"/><c:plotArea><c:layout/><c:ofPieChart><c:ofPieType val="pie"/><c:varyColors val="1"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Revenue</c:v></c:tx><c:cat><c:strLit><c:ptCount val="5"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt><c:pt idx="3"><c:v>4</c:v></c:pt><c:pt idx="4"><c:v>5</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="5"/><c:pt idx="0"><c:v>-5</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>-4</c:v></c:pt><c:pt idx="3"><c:v>2</c:v></c:pt><c:pt idx="4"><c:v>-3</c:v></c:pt></c:numLit></c:val><c:extLst><c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}" xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart"><c16:uniqueId val="{00000000-B33C-452E-82DB-9A3845A780DE}"/></c:ext></c:extLst></c:ser><c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/><c:showLeaderLines val="1"/></c:dLbls><c:gapWidth val="150"/><c:splitType val="val"/><c:splitPos val="1"/><c:secondPieSize val="75"/><c:serLines/></c:ofPieChart></c:plotArea><c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/><c:showDLblsOverMax val="0"/><c:extLst><c:ext uri="{56B9EC1D-385E-4148-901F-78D8002777C0}" xmlns:c16r3="http://schemas.microsoft.com/office/drawing/2017/03/chart"><c16r3:dataDisplayOptions16><c16r3:dispNaAsBlank val="1"/></c16r3:dataDisplayOptions16></c:ext></c:extLst></c:chart></c:chartSpace>"###, ChartOfPieOptions { secondary: Secondary::Pie, split_mode: Mode::Value, split_position: Some(1.0), custom_points: vec![], secondary_size: Some(75) }, vec![Some(-5.0),Some(1.0),Some(-4.0),Some(2.0),Some(-3.0)]),
            (r###"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c16r2="http://schemas.microsoft.com/office/drawing/2015/06/chart"><c:date1904 val="0"/><c:lang val="ja-JP"/><c:roundedCorners val="0"/><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="c14" xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart"><c14:style val="102"/></mc:Choice><mc:Fallback><c:style val="2"/></mc:Fallback></mc:AlternateContent><c:chart><c:autoTitleDeleted val="0"/><c:plotArea><c:layout/><c:ofPieChart><c:ofPieType val="pie"/><c:varyColors val="1"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Revenue</c:v></c:tx><c:cat><c:strLit><c:ptCount val="5"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt><c:pt idx="3"><c:v>4</c:v></c:pt><c:pt idx="4"><c:v>5</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="5"/><c:pt idx="0"><c:v>-5</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>-4</c:v></c:pt><c:pt idx="3"><c:v>2</c:v></c:pt><c:pt idx="4"><c:v>-3</c:v></c:pt></c:numLit></c:val><c:extLst><c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}" xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart"><c16:uniqueId val="{00000000-D752-4FD7-9E8E-3C15831A5F42}"/></c:ext></c:extLst></c:ser><c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/><c:showLeaderLines val="1"/></c:dLbls><c:gapWidth val="150"/><c:splitType val="percent"/><c:splitPos val="10"/><c:secondPieSize val="75"/><c:serLines/></c:ofPieChart></c:plotArea><c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/><c:showDLblsOverMax val="0"/><c:extLst><c:ext uri="{56B9EC1D-385E-4148-901F-78D8002777C0}" xmlns:c16r3="http://schemas.microsoft.com/office/drawing/2017/03/chart"><c16r3:dataDisplayOptions16><c16r3:dispNaAsBlank val="1"/></c16r3:dataDisplayOptions16></c:ext></c:extLst></c:chart></c:chartSpace>"###, ChartOfPieOptions { secondary: Secondary::Pie, split_mode: Mode::Percent, split_position: Some(10.0), custom_points: vec![], secondary_size: Some(75) }, vec![Some(-5.0),Some(1.0),Some(-4.0),Some(2.0),Some(-3.0)]),
            (r###"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c16r2="http://schemas.microsoft.com/office/drawing/2015/06/chart"><c:date1904 val="0"/><c:lang val="ja-JP"/><c:roundedCorners val="0"/><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="c14" xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart"><c14:style val="102"/></mc:Choice><mc:Fallback><c:style val="2"/></mc:Fallback></mc:AlternateContent><c:chart><c:autoTitleDeleted val="0"/><c:plotArea><c:layout/><c:ofPieChart><c:ofPieType val="bar"/><c:varyColors val="1"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Revenue</c:v></c:tx><c:cat><c:strLit><c:ptCount val="5"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt><c:pt idx="3"><c:v>4</c:v></c:pt><c:pt idx="4"><c:v>5</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="5"/><c:pt idx="0"><c:v>-5</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>-4</c:v></c:pt><c:pt idx="3"><c:v>2</c:v></c:pt><c:pt idx="4"><c:v>-3</c:v></c:pt></c:numLit></c:val><c:extLst><c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}" xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart"><c16:uniqueId val="{00000000-05B2-4F21-8D11-20BA42B54F20}"/></c:ext></c:extLst></c:ser><c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/><c:showLeaderLines val="1"/></c:dLbls><c:gapWidth val="150"/><c:splitType val="val"/><c:splitPos val="1"/><c:secondPieSize val="75"/><c:serLines/></c:ofPieChart></c:plotArea><c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/><c:showDLblsOverMax val="0"/><c:extLst><c:ext uri="{56B9EC1D-385E-4148-901F-78D8002777C0}" xmlns:c16r3="http://schemas.microsoft.com/office/drawing/2017/03/chart"><c16r3:dataDisplayOptions16><c16r3:dispNaAsBlank val="1"/></c16r3:dataDisplayOptions16></c:ext></c:extLst></c:chart><c:printSettings><c:headerFooter/><c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/><c:pageSetup/></c:printSettings></c:chartSpace>"###, ChartOfPieOptions { secondary: Secondary::Bar, split_mode: Mode::Value, split_position: Some(1.0), custom_points: vec![], secondary_size: Some(75) }, vec![Some(-5.0),Some(1.0),Some(-4.0),Some(2.0),Some(-3.0)]),
            (r###"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c16r2="http://schemas.microsoft.com/office/drawing/2015/06/chart"><c:date1904 val="0"/><c:lang val="ja-JP"/><c:roundedCorners val="0"/><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="c14" xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart"><c14:style val="102"/></mc:Choice><mc:Fallback><c:style val="2"/></mc:Fallback></mc:AlternateContent><c:chart><c:autoTitleDeleted val="0"/><c:plotArea><c:layout/><c:ofPieChart><c:ofPieType val="bar"/><c:varyColors val="1"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Revenue</c:v></c:tx><c:cat><c:strLit><c:ptCount val="5"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt><c:pt idx="3"><c:v>4</c:v></c:pt><c:pt idx="4"><c:v>5</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="5"/><c:pt idx="0"><c:v>-5</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>-4</c:v></c:pt><c:pt idx="3"><c:v>2</c:v></c:pt><c:pt idx="4"><c:v>-3</c:v></c:pt></c:numLit></c:val><c:extLst><c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}" xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart"><c16:uniqueId val="{00000000-A226-4B49-AD88-C8539D994904}"/></c:ext></c:extLst></c:ser><c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/><c:showLeaderLines val="1"/></c:dLbls><c:gapWidth val="150"/><c:splitType val="percent"/><c:splitPos val="10"/><c:secondPieSize val="75"/><c:serLines/></c:ofPieChart></c:plotArea><c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/><c:showDLblsOverMax val="0"/><c:extLst><c:ext uri="{56B9EC1D-385E-4148-901F-78D8002777C0}" xmlns:c16r3="http://schemas.microsoft.com/office/drawing/2017/03/chart"><c16r3:dataDisplayOptions16><c16r3:dispNaAsBlank val="1"/></c16r3:dataDisplayOptions16></c:ext></c:extLst></c:chart><c:printSettings><c:headerFooter/><c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/><c:pageSetup/></c:printSettings></c:chartSpace>"###, ChartOfPieOptions { secondary: Secondary::Bar, split_mode: Mode::Percent, split_position: Some(10.0), custom_points: vec![], secondary_size: Some(75) }, vec![Some(-5.0),Some(1.0),Some(-4.0),Some(2.0),Some(-3.0)]),
        ];
        for (xml, options, values) in controls {
            let chart = crate::chart::parse_chart_xml(xml, &crate::parser::Theme::default()).unwrap();
            assert_eq!(chart.kind, "ofPie");
            assert_eq!(chart.plot_options.of_pie, Some(options));
            assert_eq!(chart.series[0].values, values);
            let saved = new_chart_xml(&chart).unwrap();
            let reopened = crate::chart::parse_chart_xml(&saved, &crate::parser::Theme::default()).unwrap();
            assert_eq!(reopened.plot_options.of_pie, chart.plot_options.of_pie);
            assert_eq!(reopened.series[0].values, chart.series[0].values);
        }
    }

    fn split_chart(options: crate::ir::ChartOfPieOptions) -> Chart {
        Chart { kind: "ofPie".into(),
            plot_options: crate::ir::ChartPlotOptions { of_pie: Some(options), ..Default::default() },
            series: vec![ChartSeries { name: "Revenue".into(),
                category_values: Some((1..=8).map(|i| ChartCategoryValue::Text(i.to_string())).collect()),
                values: [8.0,1.0,7.0,2.0,6.0,3.0,5.0,4.0].into_iter().map(Some).collect(),
                ..Default::default() }], ..Default::default() }
    }

    #[test]
    fn split_modes_and_secondary_plots_survive_both_chart_namespaces() {
        use crate::ir::{ChartOfPieOptions, ChartPieSplitMode as Mode, ChartSecondaryPlot as Secondary};
        for strict in [false, true] {
            for secondary in [Secondary::Pie, Secondary::Bar] {
                for (split_mode, split_position, custom_points) in [
                    (Mode::Automatic, None, vec![]), (Mode::Position, Some(3.0), vec![]),
                    (Mode::Value, Some(1.0), vec![]), (Mode::Percent, Some(10.0), vec![]),
                    (Mode::Custom, None, vec![5,6,7]),
                ] {
                    let options = ChartOfPieOptions { secondary, split_mode, split_position, custom_points, secondary_size: Some(75) };
                    let chart = split_chart(options.clone());
                    let mut xml = new_chart_xml(&chart).unwrap();
                    xml = xml.replacen("</c:ofPieChart>", "<x:splitPos xmlns:x=\"urn:foreign\" val=\"999\"/></c:ofPieChart>", 1);
                    if strict { xml = xml.replace(C, "http://purl.oclc.org/ooxml/drawingml/chart")
                        .replace(A, "http://purl.oclc.org/ooxml/drawingml/main"); }
                    let read = crate::chart::parse_chart_xml(&xml, &crate::parser::Theme::default()).unwrap();
                    assert_eq!(read.plot_options.of_pie.as_ref(), Some(&options));
                    assert_eq!(read.series[0].values, chart.series[0].values);
                    assert_eq!(read.series[0].values.len(), 8);
                }
            }
        }
    }

    #[test]
    fn split_edits_and_series_additions_preserve_extensions_and_schema_order() {
        use crate::ir::{ChartOfPieOptions, ChartPieSplitMode as Mode, ChartSecondaryPlot as Secondary};
        for strict in [false, true] {
            let chart = split_chart(ChartOfPieOptions { split_mode: Mode::Position, split_position: Some(3.0),
                secondary_size: Some(75), ..Default::default() });
            let extension = "<c:extLst><c:ext uri=\"keep\"><x:keep xmlns:x=\"urn:keep\">retained</x:keep></c:ext></c:extLst>";
            let labels = "<c:dLbls><c:showVal val=\"1\"/></c:dLbls>";
            let mut xml = new_chart_xml(&chart).unwrap()
                .replacen("</c:ser>", &format!("{extension}</c:ser>"), 1)
                .replacen("<c:splitType", &format!("{labels}<c:gapWidth val=\"83\"/><c:splitType"), 1)
                .replacen("</c:ofPieChart>", &format!("{extension}</c:ofPieChart>"), 1);
            if strict { xml = xml.replace(C, "http://purl.oclc.org/ooxml/drawingml/chart")
                .replace(A, "http://purl.oclc.org/ooxml/drawingml/main"); }
            let before = crate::chart::parse_chart_xml(&xml, &crate::parser::Theme::default()).unwrap();
            let mut after = before.clone();
            after.plot_options.of_pie = Some(ChartOfPieOptions { secondary: Secondary::Bar, split_mode: Mode::Custom,
                custom_points: vec![5,6,7], secondary_size: Some(75), ..Default::default() });
            let patched = patch_chart(&xml, &before, &after).unwrap();
            assert_eq!(patched.matches("<c:splitType").count(), 1);
            assert!(!patched.contains("<c:splitPos"));
            assert_eq!(patched.matches("uri=\"keep\"").count(), 2);
            assert!(patched.contains(labels) && patched.contains("<c:gapWidth val=\"83\"/>"));
            let (fields, _) = children_at(&patched, &["chartSpace","chart","plotArea","ofPieChart"]).unwrap();
            let order: Vec<_> = fields.iter().map(|(n,_,_)| n.as_str()).collect();
            let position = |tag| order.iter().position(|n| *n == tag).unwrap();
            assert!(position("ser") < position("dLbls") && position("dLbls") < position("gapWidth"));
            assert!(position("gapWidth") < position("splitType") && position("splitType") < position("custSplit"));
            assert!(position("custSplit") < position("secondPieSize") && position("secondPieSize") < position("extLst"));
            let read = crate::chart::parse_chart_xml(&patched, &crate::parser::Theme::default()).unwrap();
            assert_eq!(read.plot_options.of_pie, after.plot_options.of_pie);
            assert_eq!(read.series[0].values.len(), 8);
            let mut added = before.clone();
            let mut next = before.series[0].clone();next.name = "Other".into();added.series.push(next);
            let appended = patch_chart(&xml, &before, &added).unwrap();
            assert!(appended.contains("<c:ofPieChart>") && appended.contains(extension));
            let read = crate::chart::parse_chart_xml(&appended, &crate::parser::Theme::default()).unwrap();
            assert_eq!(read.series.len(), 2);
            assert_eq!(read.series[1].name, "Other");
            assert_eq!(read.plot_options.of_pie, before.plot_options.of_pie);
        }
    }


    #[test]
    fn ordered_series_edits_and_appends_preserve_source_identity() {
        let controls: serde_json::Value = serde_json::from_str(include_str!(
            "../../../tests/fixtures/vba/chart_import_contracts.json")).unwrap();
        let mut checked = 0;
        for case in controls["category_cases"].as_array().unwrap().iter()
            .filter(|case| case["id"].as_str().unwrap().ends_with(":two_unique_order_swapped")) {
            let xml = case["xml"].as_str().unwrap();
            for sparse in [false, true] {
                let xml = if sparse {
                    xml.replace("<c:order val=\"1\"/>", "<c:order val=\"20\"/>")
                        .replace("<c:order val=\"0\"/>", "<c:order val=\"10\"/>")
                        .replace("<c:idx val=\"0\"/>", "<c:idx val=\"40\"/>")
                        .replace("<c:idx val=\"1\"/>", "<c:idx val=\"7\"/>")
                } else { xml.to_string() };
                let before = crate::chart::parse_chart_xml(&xml, &crate::parser::Theme::default()).unwrap();
                assert_eq!(before.series.iter().map(|s| s.name.as_str()).collect::<Vec<_>>(), ["Cost", "Revenue"]);
                let mut after = before.clone();
                after.series[0].values[0] = Some(91.0);
                let patched = patch_chart(&xml, &before, &after).unwrap();
                let read = crate::chart::parse_chart_xml(&patched, &crate::parser::Theme::default()).unwrap();
                assert_eq!(read.series[0].name, "Cost");
                assert_eq!(read.series[0].values[0], Some(91.0));
                assert_eq!(read.series[1], before.series[1], "{}", case["id"]);
                assert_eq!(read.categories, before.categories);
                let mut later = after.series[0].clone(); later.name = "Later".into();
                after.series.push(later);
                let appended = patch_chart(&xml, &before, &after).unwrap();
                let read = crate::chart::parse_chart_xml(&appended, &crate::parser::Theme::default()).unwrap();
                assert_eq!(read.series.iter().map(|s| s.name.as_str()).collect::<Vec<_>>(), ["Cost", "Revenue", "Later"]);
                assert_eq!(read.series[0].values[0], Some(91.0));
                assert_eq!(read.series[1], before.series[1]);
                let mut ids = crate::chart::series_positions(&appended, "idx");
                ids.sort(); ids.dedup(); assert_eq!(ids.len(), 3);
                assert_eq!(read.category_reference, before.category_reference);
                if !sparse {
                    let exhausted = xml.replace("<c:idx val=\"0\"/>",
                        &format!("<c:idx val=\"{}\"/>", usize::MAX));
                    assert!(patch_chart(&exhausted, &before, &after).is_err());
                }
                checked += 1;
            }
        }
        assert_eq!(checked, 20);
    }

    #[test]
    fn a_nonfinite_split_position_cannot_be_written() {
        use crate::ir::{ChartOfPieOptions, ChartPieSplitMode};
        let chart = split_chart(ChartOfPieOptions { split_mode: ChartPieSplitMode::Value,
            split_position: Some(f64::NAN), ..Default::default() });
        assert!(new_chart_xml(&chart).is_err());
    }
}
