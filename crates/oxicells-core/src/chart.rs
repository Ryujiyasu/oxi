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
use quick_xml::Reader;

use crate::ir::{
    Chart, ChartAxis, ChartPoint, ChartSeries, DataLabel, Frame, Legend, Marker, ShapeLine,
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
    let mut reader = Reader::from_str(xml);
    let mut buf = Vec::new();

    let mut chart = Chart::default();
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
    let mut categories: Vec<(usize, String)> = Vec::new();
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

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                let name = local_name(e.name().as_ref());
                let value = get_attr(e, "val");
                match name.as_str() {
                    "lineChart" | "barChart" | "pieChart" | "areaChart" | "scatterChart"
                    | "doughnutChart" | "radarChart" | "bubbleChart" | "ofPieChart"
                    | "surfaceChart" | "stockChart" => {
                        if chart.kind.is_empty() {
                            chart.kind = name.trim_end_matches("Chart").to_string();
                        }
                    }
                    "plotArea" => whose.push(Whose::Plot),
                    "ser" => {
                        whose.push(Whose::Series);
                        current = ChartSeries::default();
                        values.clear();
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
                    "spPr" | "txPr" => paint = Painted::default(),
                    "ln" => {
                        in_line = true;
                        paint.line = Some(ShapeLine {
                            color: "000000".into(),
                            width: get_attr(e, "w")
                                .and_then(|w| w.parse().ok())
                                .unwrap_or(9525),
                            dash: None,
                            head_end: None,
                            tail_end: None,
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
                    "val" => reading = Reading::Values,
                    "cat" | "xVal" => reading = Reading::Categories,
                    "tx" => reading = Reading::SeriesName,
                    "numCache" | "strCache" | "strLit" | "numLit" => in_cache = true,
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
                if in_value {
                    text.push_str(&e.unescape().unwrap_or_default());
                }
            }
            Ok(Event::End(ref e)) => {
                let name = local_name(e.name().as_ref());
                match name.as_str() {
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
                                Reading::Categories => {
                                    if let Some(index) = at {
                                        categories.push((index, text.trim().to_string()));
                                    }
                                }
                                Reading::SeriesName => current.name = text.trim().to_string(),
                                Reading::Nothing => {}
                            }
                        } else if reading == Reading::SeriesName && current.name.is_empty() {
                            // `<c:tx><c:v>` names the series outright.
                            current.name = text.trim().to_string();
                        }
                        text.clear();
                    }
                    "r" => in_run = false,
                    "numCache" | "strCache" | "strLit" | "numLit" => in_cache = false,
                    "val" | "cat" | "xVal" | "tx" => reading = Reading::Nothing,
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
                    "spPr" | "txPr" => match whose.last() {
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
                            if !paint.no_fill && chart.plot_fill.is_none() {
                                chart.plot_fill = paint.fill.take();
                            }
                        }
                        Some(Whose::Space) | None => {
                            if !paint.no_fill && chart.fill.is_none() {
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
                        if chart.categories.is_empty() && !categories.is_empty() {
                            let last =
                                categories.iter().map(|(at, _)| *at).max().unwrap_or(0);
                            let mut held = vec![String::new(); last + 1];
                            for (at, said) in &categories {
                                held[*at] = said.clone();
                            }
                            chart.categories = held;
                        }
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

    chart.series = series;
    if chart.kind.is_empty() || chart.series.is_empty() {
        return None;
    }
    Some(chart)
}

fn number(value: &Option<String>) -> Option<f64> {
    value.as_deref()?.trim().parse().ok()
}

#[cfg(test)]
mod tests {
    use super::*;

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
