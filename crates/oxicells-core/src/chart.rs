// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Reading a chart part.
//!
//! A chart hangs from the sheet's drawing but keeps its own part, and that
//! part carries a cache of every number it plots. So a chart can be drawn
//! from the file alone: nothing here reads a cell or waits on a formula.

use oxidocs_common::xml_utils::{get_attr, local_name};
use quick_xml::events::Event;

use crate::ir::{
    Chart, ChartAxis, ChartCategoryValue, ChartPoint, ChartSeries, DataLabel, Frame, Legend, Marker, ShapeLine,
};
use crate::parser::{scheme_colour, shaded, Theme};

/// Whose `<c:spPr>`, `<c:txPr>` or `<c:marker>` is being read.
#[derive(Clone, Copy, PartialEq)]
enum Whose {
    Space,
    Plot,
    Series,
    Point,
    Axis,
    Gridline,
    Legend,
    /// The block that dresses every label of a series.
    Labels,
    Label,
}

/// Where the numbers being read are going.
#[derive(Clone, Copy, PartialEq)]
enum Reading {
    Nothing,
    Values,
    BubbleSizes,
    Categories,
    SeriesName,
}

#[derive(Default)]
struct Painted {
    fill: Option<String>,
    line: Option<ShapeLine>,
    size: Option<f32>,
    face: Option<String>,
    no_fill: bool,
}

/// Reads a chart part into the picture it describes.
///
/// Returns `None` for a chart of a kind nothing draws yet, so a bar chart
/// leaves the grid as it was rather than half a picture.
pub(crate) fn parse_chart_xml(xml: &str, theme: &Theme) -> Option<Chart> {
    let mut reader = quick_xml::reader::NsReader::from_str(xml);
    let mut buf = Vec::new();

    let mut chart = Chart::default();
    let (title, automatic_title, explicit_default) = title_state(xml).unwrap_or_default();
    chart.title_from_data = title.is_none() && (automatic_title || explicit_default);
    chart.title_default_is_explicit = explicit_default;
    chart.title = title;
    let mut series: Vec<ChartSeries> = Vec::new();
    let mut current = ChartSeries::default();
    let mut point = ChartPoint::default();
    let mut label = DataLabel::default();
    let mut axis = ChartAxis::default();
    let mut legend = Legend::default();

    // Whose paint is being read, innermost last.
    let mut whose: Vec<Whose> = vec![Whose::Space];
    let mut paint = Painted::default();
    let mut in_line = false;
    let mut in_marker = false;
    let mut marker = Marker::default();

    let mut reading = Reading::Nothing;
    let mut in_cache = false;
    let mut at: Option<usize> = None;
    let mut values: Vec<(usize, f64)> = Vec::new();
    let mut bubble_sizes: Vec<(usize, f64)> = Vec::new();
    let mut bubble_size_count = 0;
    let mut categories: Vec<(usize, String)> = Vec::new();
    let mut typed_categories: Vec<(usize, ChartCategoryValue)> = Vec::new();
    let mut numeric_categories = false;
    let mut category_count = 0;
    let mut category_cache_seen = false;
    let mut active_categories_seen = false;
    let mut parents: Vec<(String, bool, bool)> = Vec::new();
    let mut in_formula = false;
    let mut text = String::new();
    let mut in_value = false;
    let mut in_run = false;
    let mut said = String::new();

    // A manual layout is stated as four numbers with no name of their own
    // beyond x, y, w and h, and both the plot and the legend state one.
    let mut layout = Frame::default();
    let mut in_layout = false;
    let mut layout_inner = true;
    let mut label_offset: Option<(f64, f64)> = None;
    let mut foreign_depth = 0usize;

    loop {
        let event = reader.read_resolved_event_into(&mut buf);
        let native = matches!(&event, Ok((quick_xml::name::ResolveResult::Bound(ns), _))
            if matches!(ns.as_ref(), b"http://schemas.openxmlformats.org/drawingml/2006/chart"
                | b"http://purl.oclc.org/ooxml/drawingml/chart"));
        let extension = matches!(&event, Ok((quick_xml::name::ResolveResult::Bound(ns), _))
            if ns.as_ref() == b"http://schemas.microsoft.com/office/drawing/2012/chart");
        let drawing = matches!(&event, Ok((quick_xml::name::ResolveResult::Bound(ns), _))
            if matches!(ns.as_ref(), b"http://schemas.openxmlformats.org/drawingml/2006/main"
                | b"http://purl.oclc.org/ooxml/drawingml/main"));
        // A foreign spPr/solidFill must neither reset our paint accumulator
        // nor provide the chart's RGB. Its entire subtree is opaque, even
        // when a descendant happens to reuse a DrawingML namespace.
        let foreign = match &event {
            Ok((_, Event::Start(_))) => {
                if foreign_depth > 0 || !(native || drawing || extension) {
                    foreign_depth += 1; true
                } else { false }
            }
            Ok((_, Event::Empty(_))) => foreign_depth > 0 || !(native || drawing || extension),
            Ok((_, Event::End(_))) if foreign_depth > 0 => { foreign_depth -= 1; true }
            Ok((_, Event::Text(_) | Event::CData(_))) => foreign_depth > 0,
            _ => false,
        };
        let direct_series = parents.last().is_some_and(|(name, native, _)| *native && name == "ser");
        let xy = parents.iter().rev().find(|(name, native, _)| *native && name.ends_with("Chart"))
            .is_some_and(|(name, _, _)| matches!(name.as_str(), "scatterChart" | "bubbleChart"));
        let filtered_title = parents.len() >= 4
            && parents[parents.len() - 1].0 == "filteredCategoryTitle" && parents[parents.len() - 1].2
            && parents[parents.len() - 2].0 == "ext" && parents[parents.len() - 2].1
            && parents[parents.len() - 3].0 == "extLst" && parents[parents.len() - 3].1
            && parents[parents.len() - 4].0 == "ser" && parents[parents.len() - 4].1;
        match &event {
            Ok((_, Event::Start(element))) =>
                parents.push((local_name(element.name().as_ref()), native, extension)),
            Ok((_, Event::End(_))) => { parents.pop(); }
            _ => {}
        }
        if foreign { buf.clear(); continue; }
        match event.map(|(_, event)| event) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                let name = local_name(e.name().as_ref());
                let value = get_attr(e, "val");
                match name.as_str() {
                    "lineChart" | "barChart" | "pieChart" | "areaChart" | "scatterChart"
                    | "line3DChart" | "bar3DChart" | "pie3DChart" | "area3DChart"
                    | "doughnutChart" | "radarChart" | "bubbleChart" | "ofPieChart"
                    | "surfaceChart" | "stockChart" => {
                        if chart.kind.is_empty() {
                            chart.kind = name.trim_end_matches("Chart").trim_end_matches("3D").to_string();
                        }
                    }
                    "plotArea" => whose.push(Whose::Plot),
                    "ser" => {
                        whose.push(Whose::Series);
                        current = ChartSeries::default();
                        values.clear();
                        bubble_sizes.clear();
                        bubble_size_count = 0;
                        category_count = 0;
                        category_cache_seen = false;
                        active_categories_seen = false;
                        numeric_categories = false;
                    }
                    "explosion" if whose.last() == Some(&Whose::Series) => {
                        current.explosion = value.as_deref().and_then(|v| v.parse().ok());
                    }
                    "dPt" => {
                        whose.push(Whose::Point);
                        point = ChartPoint::default();
                    }
                    "dLbls" => whose.push(Whose::Labels),
                    "dLbl" => {
                        whose.push(Whose::Label);
                        label = DataLabel::default();
                        label_offset = None;
                    }
                    "dLblPos" => match whose.last() {
                        Some(Whose::Label) => label.position = value,
                        Some(Whose::Labels) => current.label_pos = value,
                        _ => {}
                    },
                    "catAx" | "valAx" | "dateAx" => {
                        whose.push(Whose::Axis);
                        axis = ChartAxis::default();
                        axis.size = 10.0;
                    }
                    "majorGridlines" => whose.push(Whose::Gridline),
                    "legend" => {
                        whose.push(Whose::Legend);
                        legend = Legend::default();
                        legend.size = 10.0;
                    }
                    "idx" if whose.last() == Some(&Whose::Point) => {
                        point.index = number(&value).unwrap_or(0.0) as u32;
                    }
                    "idx" if whose.last() == Some(&Whose::Label) => {
                        label.index = number(&value).unwrap_or(0.0) as u32;
                    }
                    "marker" => {
                        in_marker = true;
                        marker = Marker { symbol: String::new(), size: 7, ..Marker::default() };
                    }
                    "symbol" if in_marker => {
                        marker.symbol = value.unwrap_or_default();
                    }
                    "size" if in_marker => {
                        marker.size = number(&value).unwrap_or(7.0) as u32;
                    }
                    // `<c:rich>` dresses text the way `<c:txPr>` does — a
                    // data label that states its own words carries its size
                    // in there and has no `txPr` at all. Both of the labels
                    // `e12c…zuhyo` draws at ten points where Excel draws
                    // fourteen are of that shape, and the `<c:spPr>` that
                    // follows was throwing the size away before anything
                    // could read it.
                    "spPr" | "txPr" | "rich" => paint = Painted::default(),
                    "ln" => {
                        in_line = true;
                        paint.line = Some(ShapeLine {
                            color: "000000".into(),
                            color_theme: None,
                            legacy_color_index: None,
                            color_brightness: None, color_shade: None,
                            width: get_attr(e, "w")
                                .and_then(|w| w.parse().ok())
                                .unwrap_or(9525),
                            dash: None,
                            cap: get_attr(e, "cap"),
                            head_end: None,
                            tail_end: None,
                            head_end_length: None,
                            head_end_width: None,
                            tail_end_length: None,
                            tail_end_width: None,
                            pattern: None,
                            back_color: None,
                            back_color_theme: None,
                            back_legacy_color_index: None,
                            back_color_brightness: None, back_color_shade: None,
                        });
                    }
                    "noFill" => {
                        if in_line {
                            paint.line = None;
                        } else {
                            paint.no_fill = true;
                        }
                    }
                    "prstDash" => {
                        if let Some(line) = &mut paint.line {
                            line.dash = value.filter(|kind| kind != "solid");
                        }
                    }
                    "srgbClr" | "schemeClr" | "sysClr" => {
                        let hex = match name.as_str() {
                            "srgbClr" => value.clone(),
                            "sysClr" => get_attr(e, "lastClr").or_else(|| value.clone()),
                            _ => value.as_deref().and_then(|v| scheme_colour(v, theme)),
                        };
                        if let Some(hex) = hex {
                            // A chart states a shade the same way a shape
                            // does, as a child of the colour.
                            let hex = shaded(&hex, &[]);
                            if in_marker {
                                if in_line {
                                    marker.line = Some(hex);
                                } else {
                                    marker.fill = Some(hex);
                                }
                            } else if in_line {
                                if let Some(line) = &mut paint.line {
                                    line.color = hex;
                                }
                            } else {
                                paint.fill = Some(hex);
                            }
                        }
                    }
                    "defRPr" => {
                        paint.size = get_attr(e, "sz")
                            .and_then(|sz| sz.parse::<f32>().ok())
                            .map(|sz| sz / 100.0);
                    }
                    "latin" | "ea" => {
                        if paint.face.is_none() {
                            paint.face = get_attr(e, "typeface");
                        }
                    }
                    "layout" => {
                        in_layout = true;
                        layout = Frame::default();
                        layout_inner = true;
                    }
                    "layoutTarget" if in_layout => {
                        layout_inner = value.as_deref() == Some("inner");
                    }
                    "x" if in_layout => layout.x = number(&value).unwrap_or(0.0),
                    "y" if in_layout => layout.y = number(&value).unwrap_or(0.0),
                    "w" if in_layout => layout.w = number(&value).unwrap_or(0.0),
                    "h" if in_layout => layout.h = number(&value).unwrap_or(0.0),
                    "max" => axis.max = number(&value),
                    "min" => axis.min = number(&value),
                    "majorUnit" => axis.major_unit = number(&value),
                    "axPos" => axis.position = value.unwrap_or_default(),
                    "majorTickMark" => axis.major_tick = value.unwrap_or_default(),
                    "tickLblPos" => axis.tick_labels = value.unwrap_or_default(),
                    "crossBetween" => axis.cross_between = value,
                    "delete" if whose.last() == Some(&Whose::Axis) => {
                        axis.deleted = value.as_deref() == Some("1");
                    }
                    "numFmt" => {
                        let code = get_attr(e, "formatCode");
                        match whose.last() {
                            Some(Whose::Axis) => axis.number_format = code,
                            Some(Whose::Label) => label.number_format = code,
                            _ => {}
                        }
                    }
                    "legendPos" => legend.position = value.unwrap_or_default(),
                    "val" | "yVal" => {
                        reading = if native && direct_series && ((name == "yVal") == xy) {
                            Reading::Values
                        } else { Reading::Nothing };
                    },
                    "bubbleSize" => {
                        reading = Reading::BubbleSizes;
                        current.bubble_sizes = Some(Vec::new());
                    }
                    "bubble3D" if whose.last() == Some(&Whose::Series) => {
                        current.shaded_bubbles = !matches!(value.as_deref(), Some("0" | "false"));
                    }
                    "ptCount" if reading == Reading::BubbleSizes => {
                        bubble_size_count = value.as_deref().and_then(|v| v.parse::<usize>().ok()).unwrap_or(0);
                    }
                    "ptCount" if reading == Reading::Categories => {
                        category_count = value.as_deref().and_then(|v| v.parse::<usize>().ok()).unwrap_or(0);
                    }
                    "cat" | "xVal" => {
                        let active = native && direct_series && ((name == "xVal") == xy);
                        let filtered = extension && name == "cat" && filtered_title && !active_categories_seen;
                        if active || filtered {
                            categories.clear();
                            typed_categories.clear();
                            category_count = 0;
                            category_cache_seen = false;
                            numeric_categories = false;
                            current.category_reference = None;
                            current.category_values_filtered = filtered;
                            active_categories_seen |= active;
                            reading = Reading::Categories;
                        } else { reading = Reading::Nothing; }
                    },
                    "tx" => {
                        reading = Reading::SeriesName;
                        if whose.last() == Some(&Whose::Series) {
                            current.explicit_empty_name = true;
                        }
                    },
                    "numCache" | "strCache" | "strLit" | "numLit" => {
                        in_cache = true;
                        if reading == Reading::Categories {
                            numeric_categories = matches!(name.as_str(), "numCache" | "numLit");
                            category_cache_seen = true;
                        }
                    }
                    "f" if whose.contains(&Whose::Series) => {
                        in_formula = true;
                        text.clear();
                    }
                    "pt" => at = get_attr(e, "idx").and_then(|i| i.parse().ok()),
                    "v" => {
                        in_value = true;
                        text.clear();
                    }
                    "r" => in_run = true,
                    "t" if in_run => {
                        in_value = true;
                        text.clear();
                    }
                    _ => {}
                }
            }
            Ok(Event::Text(e)) => {
                if in_value || in_formula {
                    text.push_str(&e.unescape().unwrap_or_default());
                }
            }
            Ok(Event::End(ref e)) => {
                let name = local_name(e.name().as_ref());
                match name.as_str() {
                    "f" if in_formula => {
                        in_formula = false;
                        match reading {
                            Reading::SeriesName => current.name_reference = Some(text.clone()),
                            Reading::Values => current.values_reference = Some(text.clone()),
                            Reading::BubbleSizes => current.bubble_sizes_reference = Some(text.clone()),
                            Reading::Categories => current.category_reference = Some(text.clone()),
                            Reading::Nothing => {}
                        }
                        text.clear();
                    }
                    "v" | "t" => {
                        in_value = false;
                        if in_run {
                            said.push_str(&text);
                        } else if in_cache {
                            match reading {
                                Reading::Values => {
                                    if let (Some(index), Ok(number)) =
                                        (at, text.trim().parse::<f64>())
                                    {
                                        values.push((index, number));
                                    }
                                }
                                Reading::BubbleSizes => {
                                    if let (Some(index), Ok(number)) = (at, text.trim().parse::<f64>()) {
                                        bubble_sizes.push((index, number));
                                    }
                                }
                                Reading::Categories => {
                                    if let Some(index) = at {
                                        categories.push((index, text.clone()));
                                        let value = if numeric_categories {
                                            text.trim().parse::<f64>().ok().map(ChartCategoryValue::Number)
                                                .unwrap_or(ChartCategoryValue::Blank)
                                        } else { ChartCategoryValue::Text(text.clone()) };
                                        typed_categories.push((index, value));
                                    }
                                }
                                Reading::SeriesName => current.name = text.clone(),
                                Reading::Nothing => {}
                            }
                        } else if reading == Reading::SeriesName && current.name.is_empty() {
                            // `<c:tx><c:v>` names the series outright.
                            current.name = text.clone();
                        }
                        text.clear();
                    }
                    "r" => in_run = false,
                    "numCache" | "strCache" | "strLit" | "numLit" => in_cache = false,
                    "tx" => {
                        if whose.last() == Some(&Whose::Series) {
                            current.explicit_empty_name = current.name.is_empty() && current.name_reference.is_none();
                        }
                        reading = Reading::Nothing;
                    },
                    "val" | "yVal" | "bubbleSize" | "cat" | "xVal" => reading = Reading::Nothing,
                    "ln" => in_line = false,
                    "marker" => {
                        in_marker = false;
                        if !marker.symbol.is_empty() {
                            match whose.last() {
                                Some(Whose::Point) => point.marker = Some(marker.clone()),
                                _ => current.marker = Some(marker.clone()),
                            }
                        }
                    }
                    "spPr" | "txPr" | "rich" => match whose.last() {
                        Some(Whose::Series) => {
                            if current.line.is_none() {
                                current.line = paint.line.take();
                            }
                        }
                        Some(Whose::Point) => {
                            if point.line.is_none() {
                                point.line = paint.line.take();
                            }
                        }
                        Some(Whose::Axis) => {
                            if axis.line.is_none() {
                                axis.line = paint.line.take();
                            }
                            if let Some(size) = paint.size {
                                axis.size = size;
                            }
                            if axis.face.is_none() {
                                axis.face = paint.face.take();
                            }
                        }
                        Some(Whose::Gridline) => {
                            if axis.major_gridline.is_none() {
                                axis.major_gridline = paint.line.take();
                            }
                        }
                        Some(Whose::Legend) => {
                            if let Some(size) = paint.size {
                                legend.size = size;
                            }
                            if legend.face.is_none() {
                                legend.face = paint.face.take();
                            }
                        }
                        Some(Whose::Label) => {
                            if let Some(size) = paint.size {
                                label.size = size;
                            }
                            if label.face.is_none() {
                                label.face = paint.face.take();
                            }
                        }
                        Some(Whose::Labels) => {
                            if let Some(size) = paint.size {
                                current.label_size = size;
                            }
                            if current.label_face.is_none() {
                                current.label_face = paint.face.take();
                            }
                        }
                        Some(Whose::Plot) => {
                            if name == "spPr" && native && !paint.no_fill
                                && parents.last().is_some_and(|(parent, chart, _)| *chart && parent == "plotArea") {
                                chart.plot_fill = paint.fill.take();
                            }
                        }
                        Some(Whose::Space) | None => {
                            // Title/text properties also contain solidFill.
                            // Only the chart space's own spPr is its background.
                            if name == "spPr" && native && !paint.no_fill
                                && parents.last().is_some_and(|(parent, chart, _)| *chart && parent == "chartSpace") {
                                chart.fill = paint.fill.take();
                            }
                        }
                    },
                    "layout" => {
                        in_layout = false;
                        let stated = layout.w > 0.0 || layout.h > 0.0;
                        match whose.last() {
                            Some(Whose::Plot) if stated && layout_inner => {
                                chart.plot = Some(layout)
                            }
                            Some(Whose::Legend) if stated => legend.frame = Some(layout),
                            Some(Whose::Label) if stated || layout.x != 0.0 || layout.y != 0.0 => {
                                label_offset = Some((layout.x, layout.y));
                            }
                            _ => {}
                        }
                    }
                    "dPt" => {
                        whose.pop();
                        if point.marker.is_some() || point.line.is_some() {
                            current.points.push(point.clone());
                        }
                    }
                    "dLbl" => {
                        whose.pop();
                        label.offset = label_offset;
                        if !said.is_empty() {
                            label.text = Some(said.clone());
                        }
                        said.clear();
                        current.labels.push(label.clone());
                    }
                    "ser" => {
                        whose.pop();
                        let last = values.iter().map(|(at, _)| *at).max();
                        current.values = match last {
                            Some(last) => {
                                let mut held = vec![None; last + 1];
                                for (at, number) in &values {
                                    held[*at] = Some(*number);
                                }
                                held
                            }
                            None => Vec::new(),
                        };
                        if current.bubble_sizes.is_some() {
                            let len = bubble_size_count.max(bubble_sizes.iter().map(|(i, _)| i + 1).max().unwrap_or(0));
                            let mut held = vec![None; len];
                            for (i, number) in &bubble_sizes { held[*i] = Some(*number); }
                            current.bubble_sizes = Some(held);
                        }
                        if !current.category_values_filtered && chart.categories.is_empty() && !categories.is_empty() {
                            let last =
                                categories.iter().map(|(at, _)| *at).max().unwrap_or(0);
                            let mut held = vec![String::new(); last + 1];
                            for (at, said) in &categories {
                                held[*at] = said.clone();
                            }
                            chart.categories = held;
                        }
                        let mut typed = vec![ChartCategoryValue::Blank;
                            category_count.max(typed_categories.iter().map(|(i, _)| i + 1).max().unwrap_or(0))];
                        for (i, value) in typed_categories.drain(..) { typed[i] = value; }
                        current.text_category_gaps = category_cache_seen && (typed.is_empty() || (!numeric_categories
                            && typed.iter().all(|value| matches!(value, ChartCategoryValue::Blank))));
                        current.category_values = Some(typed);
                        categories.clear();
                        series.push(current.clone());
                    }
                    "catAx" | "valAx" | "dateAx" => {
                        whose.pop();
                        if name == "valAx" {
                            chart.value_axis = Some(axis.clone());
                        } else {
                            chart.category_axis = Some(axis.clone());
                        }
                    }
                    "majorGridlines" | "dLbls" => {
                        whose.pop();
                    }
                    "legend" => {
                        whose.pop();
                        chart.legend = Some(legend.clone());
                    }
                    "plotArea" => {
                        whose.pop();
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    chart.category_reference = series.first().and_then(|series| series.category_reference.clone());
    chart.series = series;
    // An explicitly enabled automatic title is materialized on file open.
    // Pie and of-pie plots use the first series even when more are stored.
    if chart.title.is_none() && (automatic_title || explicit_default)
        && (chart.series.len() == 1 || matches!(chart.kind.as_str(), "pie" | "ofPie"))
    {
        chart.title = chart.series.first()
            .filter(|series| !series.name.is_empty() || series.explicit_empty_name || series.name_reference.is_some())
            .map(|series| series.name.clone());
    }
    if explicit_default && chart.title.is_none() {
        chart.title = Some("\u{30b0}\u{30e9}\u{30d5} \u{30bf}\u{30a4}\u{30c8}\u{30eb}".into());
        chart.title_from_data = false;
    }
    // Automatic titles and shared category labels use source order, while
    // the SeriesCollection follows each series' explicit order.
    if chart.title_from_data || explicit_default { chart.title_from_data_cache = chart.title.clone(); }
    let orders = series_positions(xml, "order");
    let mut ordered: Vec<_> = chart.series.into_iter().enumerate().collect();
    ordered.sort_by_key(|(index, _)| orders.get(*index).copied().flatten().unwrap_or(*index));
    let mut source_order = vec![0; ordered.len()];
    for (logical, (source, _)) in ordered.iter().enumerate() { source_order[*source] = logical; }
    if source_order.iter().copied().ne(0..source_order.len()) { chart.source_series_order = source_order; }
    chart.series = ordered.into_iter().map(|(_, series)| series).collect();
    // An empty chart is still an object: VBA can name it, set its properties
    // and add series later. Keep it even when there is nothing to plot yet.
    if chart.kind.is_empty() {
        return None;
    }
    chart.plot_options = plot_options(xml, &chart.kind);
    Some(chart)
}


/// Source positions are independent of the series' persistent identifiers.
pub(crate) fn series_positions(xml: &str, field: &str) -> Vec<Option<usize>> {
    use quick_xml::{reader::NsReader, name::ResolveResult};
    let mut reader = NsReader::from_str(xml);
    let mut parents: Vec<(String, bool)> = Vec::new();
    let mut positions = Vec::new();
    loop {
        let Ok((namespace, event)) = reader.read_resolved_event() else { break };
        let chart_namespace = matches!(namespace, ResolveResult::Bound(ref ns)
            if matches!(ns.as_ref(), b"http://schemas.openxmlformats.org/drawingml/2006/chart"
                | b"http://purl.oclc.org/ooxml/drawingml/chart"));
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                let name = local_name(element.name().as_ref());
                let at_plot = parents.len() >= 4 && parents.iter().all(|(_, chart)| *chart)
                    && parents[0].0 == "chartSpace" && parents[1].0 == "chart"
                    && parents[2].0 == "plotArea" && parents[3].0.ends_with("Chart");
                if chart_namespace && name == "ser" && parents.len() == 4 && at_plot {
                    positions.push(None);
                } else if chart_namespace && name == field && parents.len() == 5 && at_plot
                    && parents[4].0 == "ser" {
                    if let Some(position) = positions.last_mut() {
                        *position = get_attr(element, "val").and_then(|value| value.parse().ok());
                    }
                }
                if matches!(event, Event::Start(_)) { parents.push((name, chart_namespace)); }
            }
            Event::End(_) => { parents.pop(); }
            Event::Eof => break,
            _ => {}
        }
    }
    positions
}

/// Plot-wide settings and the first series jointly determine the visible plot.
/// A canonical scatter plot can say "lineMarker" while hiding its series line.
fn plot_options(xml: &str, kind: &str) -> crate::ir::ChartPlotOptions {
    use crate::ir::{ChartBarShape, ChartPlotOptions, ChartStacking};
    use quick_xml::{reader::NsReader, name::ResolveResult};
    let mut reader = NsReader::from_str(xml);
    let mut parents: Vec<(String, u8)> = Vec::new();
    let mut options = ChartPlotOptions::default();
    let mut selected = String::new();
    let mut active = false;
    let mut first_series = false;
    let mut in_series = false;
    let mut series_markers = None;
    let mut no_line = false;
    loop {
        let Ok((namespace, event)) = reader.read_resolved_event() else { break };
        let domain = match namespace {
            ResolveResult::Bound(ns) if matches!(ns.as_ref(), b"http://schemas.openxmlformats.org/drawingml/2006/chart" | b"http://purl.oclc.org/ooxml/drawingml/chart") => 1,
            ResolveResult::Bound(ns) if matches!(ns.as_ref(), b"http://schemas.openxmlformats.org/drawingml/2006/main" | b"http://purl.oclc.org/ooxml/drawingml/main") => 2,
            _ => 0,
        };
        match event {
            Event::Start(ref e) | Event::Empty(ref e) => {
                let name = local_name(e.name().as_ref());
                let value = get_attr(e, "val");
                let at_plot = active && parents.len() == 4 && parents.iter().all(|(_, ns)| *ns == 1);
                if selected.is_empty() && domain == 1 && name.ends_with("Chart") && parents.len() == 3
                    && parents.iter().map(|(n, _)| n.as_str()).eq(["chartSpace", "chart", "plotArea"])
                    && parents.iter().all(|(_, ns)| *ns == 1) {
                    selected = name.clone(); active = true;
                    options.three_dimensional = name.ends_with("3DChart");
                    if name == "ofPieChart" { options.of_pie = Some(Default::default()); }
                } else if domain == 1 && at_plot {
                    match name.as_str() {
                        "ofPieType" => if let Some(pie) = &mut options.of_pie {
                            pie.secondary = if value.as_deref() == Some("bar") { crate::ir::ChartSecondaryPlot::Bar }
                                else { crate::ir::ChartSecondaryPlot::Pie };
                        },
                        "splitType" => if let Some(pie) = &mut options.of_pie {
                            use crate::ir::ChartPieSplitMode;
                            pie.split_mode = match value.as_deref() {
                                Some("pos") => ChartPieSplitMode::Position,
                                Some("val") => ChartPieSplitMode::Value,
                                Some("percent") => ChartPieSplitMode::Percent,
                                Some("cust") => ChartPieSplitMode::Custom,
                                _ => ChartPieSplitMode::Automatic,
                            };
                        },
                        "splitPos" => if let Some(pie) = &mut options.of_pie { pie.split_position = number(&value); },
                        "secondPieSize" => if let Some(pie) = &mut options.of_pie {
                            pie.secondary_size = value.as_deref().and_then(|v| v.parse().ok());
                        },
                        "barDir" => options.horizontal = value.as_deref() == Some("bar"),
                        "grouping" => {
                            options.stacking = match value.as_deref() {
                                Some("stacked") => ChartStacking::Stacked,
                                Some("percentStacked") => ChartStacking::Percent,
                                _ => ChartStacking::None,
                            };
                            options.depth_series = selected == "bar3DChart" && value.as_deref() == Some("standard");
                        },
                        "shape" => options.bar_shape = match value.as_deref() {
                            Some("cylinder") => ChartBarShape::Cylinder,
                            Some("cone") => ChartBarShape::Cone,
                            Some("pyramid") => ChartBarShape::Pyramid,
                            _ => ChartBarShape::Box,
                        },
                        "radarStyle" => {
                            options.filled = value.as_deref() == Some("filled");
                            options.markers = Some(value.as_deref() != Some("standard"));
                        }
                        "scatterStyle" => {
                            options.smooth = value.as_deref().is_some_and(|s| s.starts_with("smooth"));
                            options.markers = Some(!matches!(value.as_deref(), Some("line" | "smooth")));
                        }
                        "marker" => options.markers = Some(value.as_deref() != Some("0")),
                        "smooth" => options.smooth = value.as_deref() == Some("1"),
                        "ser" if !first_series => { first_series = true; in_series = true; }
                        _ => {}
                    }
                }
                if active && domain == 1 && name == "secondPiePt" && parents.len() == 5
                    && parents.last().is_some_and(|(name, ns)| name == "custSplit" && *ns == 1)
                    && parents.iter().all(|(_, ns)| *ns == 1) {
                    if let (Some(pie), Some(index)) = (&mut options.of_pie, value.as_deref().and_then(|v| v.parse().ok())) {
                        pie.custom_points.push(index);
                    }
                }
                if active && in_series {
                    if domain == 1 && name == "explosion" && parents.len() == 5 {
                        options.explosion = value.as_deref().and_then(|v| v.parse().ok());
                    }
                    if domain == 1 && name == "symbol" && parents.len() == 6
                        && parents[4] == ("ser".into(), 1) && parents[5] == ("marker".into(), 1) {
                        series_markers = Some(value.as_deref() != Some("none"));
                    }
                    if domain == 2 && name == "noFill" && parents.len() == 7
                        && parents[4] == ("ser".into(), 1) && parents[5] == ("spPr".into(), 1)
                        && parents[6] == ("ln".into(), 2) { no_line = true; }
                }
                if matches!(event, Event::Start(_)) { parents.push((name, domain)); }
            }
            Event::End(ref e) => {
                let name = local_name(e.name().as_ref());
                if domain == 1 && name == "ser" && parents.len() == 5 { in_series = false; }
                if domain == 1 && name == selected && parents.len() == 4 { active = false; }
                parents.pop();
            }
            Event::Eof => break,
            _ => {}
        }
    }
    if matches!(kind, "line" | "scatter" | "radar") && !(kind == "line" && options.three_dimensional) {
        options.markers = Some(series_markers.or(options.markers).unwrap_or(true));
    }
    if kind == "scatter" { options.connected = Some(!first_series || !no_line); }
    options
}

fn number(value: &Option<String>) -> Option<f64> {
    value.as_deref()?.trim().parse().ok()
}

/// Only the chart's own title, not a series, axis or extension's text.
fn title_state(xml:&str)->Option<(Option<String>,bool,bool)> {
    use quick_xml::{reader::NsReader,name::ResolveResult};
    let mut reader=NsReader::from_str(xml);
    let mut parents:Vec<(String,bool,bool)>=Vec::new();
    let mut title=None;
    let mut title_present=false;
    let mut depth=None;
    let mut automatic=false;
    loop {
        let (namespace,event)=reader.read_resolved_event().ok()?;
        let chart=matches!(&namespace,ResolveResult::Bound(ns) if matches!(ns.as_ref(),b"http://schemas.openxmlformats.org/drawingml/2006/chart"|b"http://purl.oclc.org/ooxml/drawingml/chart"));
        let drawing=matches!(&namespace,ResolveResult::Bound(ns) if matches!(ns.as_ref(),b"http://schemas.openxmlformats.org/drawingml/2006/main"|b"http://purl.oclc.org/ooxml/drawingml/main"));
        match event {
            Event::Start(ref element)|Event::Empty(ref element)=> {
                let name=local_name(element.name().as_ref());
                if chart && name=="autoTitleDeleted" && parents.len()==2 && parents[0].0=="chartSpace" && parents[0].1 && parents[1].0=="chart" && parents[1].1 {
                    automatic=matches!(get_attr(element,"val").as_deref(),Some("0"|"false"));
                }
                if chart && name=="title" && parents.len()==2 && parents[0].0=="chartSpace" && parents[0].1 && parents[1].0=="chart" && parents[1].1 {
                    title_present=true;
                    if matches!(event,Event::Start(_)) {depth=Some(parents.len()+1);}
                }
                if chart && name=="tx" && depth==Some(parents.len()) {
                    title=Some(String::new());
                }
                if matches!(event,Event::Start(_)) {parents.push((name,chart,drawing));}
            }
            Event::Text(text) if depth.is_some() && parents.last().is_some_and(|(name,chart,drawing)|(name=="t" && *drawing)||(name=="v" && *chart))=> {
                let raw=std::str::from_utf8(text.as_ref()).ok()?.replace("\r\n","\n").replace('\r',"\n");
                title.as_mut()?.push_str(&quick_xml::escape::unescape(&raw).ok()?);
            }
            Event::CData(text) if depth.is_some() && parents.last().is_some_and(|(name,chart,drawing)|(name=="t" && *drawing)||(name=="v" && *chart))=> {
                title.as_mut()?.push_str(&String::from_utf8_lossy(text.as_ref()).replace("\r\n","\n").replace('\r',"\n"));
            }
            Event::End(_)=> {if depth==Some(parents.len()) {depth=None;} parents.pop();}
            Event::Eof=>break,
            _=>{}
        }
    }
    let explicit_default=title_present && title.is_none();
    Some((title,automatic,explicit_default))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn cold_file_titles_match_actual_excel_across_chart_families() {
        let controls: serde_json::Value = serde_json::from_str(include_str!(
            "../../../tests/fixtures/vba/chart_import_contracts.json")).unwrap();
        let mut differences = Vec::new();
        for case in controls["title_cases"].as_array().unwrap() {
            let chart = parse_chart_xml(case["xml"].as_str().unwrap(), &Theme::default()).unwrap();
            let expected = case["native_title"].as_str().unwrap().strip_prefix("True:").map(str::to_string);
            if chart.title != expected {
                differences.push(format!("{}: {:?} != {:?}", case["id"], chart.title, expected));
            }
        }
        assert!(differences.is_empty(), "{} differences: {:?}", differences.len(), differences);
    }



    const LINE: &str = r##"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:plotArea>
      <c:layout><c:manualLayout><c:layoutTarget val="inner"/>
        <c:x val="0.1"/><c:y val="0.2"/><c:w val="0.8"/><c:h val="0.6"/>
      </c:manualLayout></c:layout>
      <c:lineChart>
        <c:ser>
          <c:idx val="0"/>
          <c:tx><c:v>男女計</c:v></c:tx>
          <c:spPr><a:ln w="22225"><a:solidFill><a:srgbClr val="112233"/></a:solidFill>
            <a:prstDash val="dash"/></a:ln></c:spPr>
          <c:marker><c:symbol val="none"/></c:marker>
          <c:dPt><c:idx val="1"/><c:marker><c:symbol val="circle"/><c:size val="8"/>
            </c:marker></c:dPt>
          <c:dLbls><c:dLbl><c:idx val="1"/>
            <c:layout><c:manualLayout><c:x val="-0.05"/><c:y val="0.01"/>
              </c:manualLayout></c:layout>
            <c:txPr><a:p><a:pPr><a:defRPr sz="900"/></a:pPr></a:p></c:txPr>
          </c:dLbl></c:dLbls>
          <c:cat><c:strRef><c:strCache>
            <c:pt idx="0"><c:v>昭51</c:v></c:pt>
            <c:pt idx="1"><c:v>52</c:v></c:pt>
            <c:pt idx="2"><c:v>53</c:v></c:pt>
          </c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:numCache>
            <c:pt idx="0"><c:v>130.5</c:v></c:pt>
            <c:pt idx="2"><c:v>150</c:v></c:pt>
          </c:numCache></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:idx val="1"/>
          <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>男</c:v></c:pt>
            </c:strCache></c:strRef></c:tx>
          <c:val><c:numRef><c:numCache>
            <c:pt idx="0"><c:v>90</c:v></c:pt>
          </c:numCache></c:numRef></c:val>
        </c:ser>
      </c:lineChart>
      <c:catAx><c:axPos val="b"/><c:majorTickMark val="in"/>
        <c:tickLblPos val="nextTo"/>
        <c:spPr><a:ln w="3175"><a:solidFill><a:srgbClr val="000000"/></a:solidFill>
          </a:ln></c:spPr>
        <c:txPr><a:p><a:pPr><a:defRPr sz="1000"><a:latin typeface="ＭＳ 明朝"/>
          </a:defRPr></a:pPr></a:p></c:txPr>
      </c:catAx>
      <c:valAx><c:scaling><c:max val="350"/></c:scaling><c:axPos val="l"/>
        <c:majorTickMark val="in"/><c:numFmt formatCode="0_ " sourceLinked="0"/>
        <c:crossBetween val="midCat"/>
      </c:valAx>
      <c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></c:spPr>
    </c:plotArea>
    <c:legend><c:legendPos val="r"/>
      <c:layout><c:manualLayout><c:x val="0.5"/><c:y val="0.55"/>
        <c:w val="0.3"/><c:h val="0.2"/></c:manualLayout></c:layout>
      <c:txPr><a:p><a:pPr><a:defRPr sz="1200"/></a:pPr></a:p></c:txPr>
    </c:legend>
  </c:chart>
</c:chartSpace>"##;

    #[test]
    fn a_line_chart_gives_up_what_it_plots() {
        let chart = parse_chart_xml(LINE, &Theme::default()).expect("a line chart");
        assert_eq!(chart.kind, "line");
        assert_eq!(chart.series.len(), 2);

        let first = &chart.series[0];
        assert_eq!(first.name, "男女計");
        // A cached point that is missing leaves a gap, not a zero.
        assert_eq!(first.values, vec![Some(130.5), None, Some(150.0)]);
        assert_eq!(chart.categories, vec!["昭51", "52", "53"]);
        assert_eq!(chart.series[1].name, "男");

        let line = first.line.as_ref().expect("a series line");
        assert_eq!(line.color, "112233");
        assert_eq!(line.width, 22225);
        assert_eq!(line.dash.as_deref(), Some("dash"));
    }

    #[test]
    fn a_point_that_dresses_differently_is_kept_apart() {
        let chart = parse_chart_xml(LINE, &Theme::default()).unwrap();
        let first = &chart.series[0];
        // The series wears no marker; one of its points wears a circle.
        assert!(first.marker.as_ref().map(|m| m.symbol.as_str()) == Some("none"));
        assert_eq!(first.points.len(), 1);
        assert_eq!(first.points[0].index, 1);
        let marker = first.points[0].marker.as_ref().unwrap();
        assert_eq!(marker.symbol, "circle");
        assert_eq!(marker.size, 8);

        assert_eq!(first.labels.len(), 1);
        assert_eq!(first.labels[0].index, 1);
        assert_eq!(first.labels[0].offset, Some((-0.05, 0.01)));
        assert_eq!(first.labels[0].size, 9.0);
    }

    #[test]
    fn the_plot_the_axes_and_the_legend_keep_their_own_places() {
        let chart = parse_chart_xml(LINE, &Theme::default()).unwrap();
        let plot = chart.plot.expect("a pinned plot area");
        assert!((plot.x - 0.1).abs() < 1e-9 && (plot.h - 0.6).abs() < 1e-9);

        let value = chart.value_axis.expect("a value axis");
        assert_eq!(value.max, Some(350.0));
        assert_eq!(value.min, None);
        assert_eq!(value.major_tick, "in");
        assert_eq!(value.number_format.as_deref(), Some("0_ "));
        assert_eq!(value.cross_between.as_deref(), Some("midCat"));

        let category = chart.category_axis.expect("a category axis");
        assert_eq!(category.position, "b");
        assert_eq!(category.size, 10.0);
        assert_eq!(category.face.as_deref(), Some("ＭＳ 明朝"));
        assert_eq!(category.line.as_ref().map(|l| l.width), Some(3175));

        let legend = chart.legend.expect("a legend");
        assert_eq!(legend.position, "r");
        assert_eq!(legend.size, 12.0);
        assert!((legend.frame.unwrap().w - 0.3).abs() < 1e-9);
    }

    #[test]
    fn a_chart_of_a_kind_nothing_draws_yet_is_left_alone() {
        let bar = LINE.replace("lineChart", "barChart");
        let chart = parse_chart_xml(&bar, &Theme::default()).unwrap();
        assert_eq!(chart.kind, "bar");

        // Nothing to plot at all is not a chart.
        assert!(parse_chart_xml("<c:chartSpace/>", &Theme::default()).is_none());
    }
}
