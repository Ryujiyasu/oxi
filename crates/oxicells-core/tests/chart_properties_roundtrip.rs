// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use oxicells_core::{ir::DrawingKind, parse_xlsx, XlsxEditor};
use std::io::{Cursor, Read};

#[test]
fn chart_properties_roundtrip_matches_native_excel_saved_parts() {
    for name in [
        "sheet_title_legend",
        "sheet_unicode",
        "sheet_hide_title",
        "sheet_default_title",
        "sheet_empty_title",
        "embedded_title_legend",
        "sheet_legend_left",
        "sheet_restore_title",
    ] {
        let fixtures =
            std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../tests/fixtures/vba");
        let original =
            std::fs::read(fixtures.join(format!("chart_persistence_{name}.xlsm"))).unwrap();
        let reference =
            std::fs::read(fixtures.join(format!("chart_persistence_{name}_saved.xlsm"))).unwrap();
        let expected = parse_xlsx(&reference).unwrap();
        let mut edited = parse_xlsx(&original).unwrap();
        let expected_chart = expected
            .sheets
            .iter()
            .flat_map(|sheet| &sheet.drawings)
            .find_map(|drawing| {
                if let DrawingKind::Chart(chart) = &drawing.kind {
                    Some(chart)
                } else {
                    None
                }
            })
            .unwrap();
        if name=="sheet_unicode" {
            assert_eq!(expected_chart.title.as_deref(),Some("売上 & <利益>\n2026"));
        }
        let chart = edited
            .sheets
            .iter_mut()
            .flat_map(|sheet| &mut sheet.drawings)
            .find_map(|drawing| {
                if let DrawingKind::Chart(chart) = &mut drawing.kind {
                    Some(chart)
                } else {
                    None
                }
            })
            .unwrap();
        chart.title = expected_chart.title.clone();
        chart.legend = expected_chart.legend.clone();
        let mut editor = XlsxEditor::new(&original).unwrap();
        editor.apply_workbook(&edited).unwrap();
        let saved = editor.save().unwrap();
        let actual = parse_xlsx(&saved).unwrap();
        let actual_chart = actual
            .sheets
            .iter()
            .flat_map(|sheet| &sheet.drawings)
            .find_map(|drawing| {
                if let DrawingKind::Chart(chart) = &drawing.kind {
                    Some(chart)
                } else {
                    None
                }
            })
            .unwrap();
        assert_eq!(actual_chart.title, expected_chart.title, "{name}");
        assert_eq!(
            actual_chart.legend.as_ref().map(|legend| &legend.position),
            expected_chart
                .legend
                .as_ref()
                .map(|legend| &legend.position),
            "{name}"
        );
        // The second embedded chart and every unrelated OOXML part must survive
        // byte for byte; this is not a whole-chart/package regeneration.
        let mut before = zip::ZipArchive::new(Cursor::new(&original)).unwrap();
        let mut after = zip::ZipArchive::new(Cursor::new(&saved)).unwrap();
        for index in 0..before.len() {
            let mut entry = before.by_index(index).unwrap();
            if entry.name() == "xl/charts/chart1.xml" {
                continue;
            }
            let mut old = Vec::new();
            entry.read_to_end(&mut old).unwrap();
            let mut current = Vec::new();
            after
                .by_name(entry.name())
                .unwrap()
                .read_to_end(&mut current)
                .unwrap();
            assert_eq!(old, current, "{name}: unrelated part {}", entry.name());
        }
    }
}
