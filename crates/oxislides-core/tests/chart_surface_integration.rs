// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Surface charts (`tools/metrics/gen_pptx_chart_surface.py`): both
//! `c:surfaceChart` and `c:surface3DChart` arrive as one chart type with their
//! series and the wireframe switch.

use oxislides_core::ir::ShapeContent;

#[test]
fn surface_charts_keep_their_series_and_wireframe_flag() {
    let data = include_bytes!("../../../tests/fixtures/pptx_chart_surface.pptx");
    let p = oxislides_core::parser::parse_pptx(data).expect("deck parses");
    let charts: Vec<_> = p
        .slides
        .iter()
        .filter_map(|s| {
            s.shapes.iter().find_map(|sh| match &sh.content {
                ShapeContent::Chart { chart } => Some(chart.clone()),
                _ => None,
            })
        })
        .collect();
    assert_eq!(charts.len(), 6);
    for c in &charts {
        assert_eq!(c.chart_type, "surface");
        assert_eq!(c.categories, vec!["Q1", "Q2", "Q3", "Q4"]);
    }
    let wire: Vec<bool> = charts.iter().map(|c| c.surface_wireframe).collect();
    assert_eq!(wire, vec![false, true, false, true, false, false]);
    let nser: Vec<usize> = charts.iter().map(|c| c.series.len()).collect();
    assert_eq!(nser, vec![2, 2, 2, 2, 2, 3]);
    assert_eq!(charts[5].series[2].values, vec![5.1, 6.4, 7.2, 4.8]);
}
