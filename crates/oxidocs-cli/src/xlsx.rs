// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Spreadsheet print pages, drawn with this machine's font files. The layout
//! lives in `oxicells_core::print`, shared with the browser build.
use super::pptx::TextPages;
use oxidocs_core::ir::UnsupportedContent;
use oxipdf_core::ir::*;

pub use oxicells_core::print::is_workbook;

impl oxicells_core::print::PrintCanvas for TextPages {
    fn page(&mut self, width: f64, height: f64) {
        TextPages::page(self, width, height)
    }
    fn element(&mut self, element: ContentElement) {
        TextPages::element(self, element)
    }
    fn width(&self, text: &str, family: &str, size: f64, bold: bool, italic: bool) -> f64 {
        TextPages::width(self, text, family, size, bold, italic)
    }
    fn text(&mut self, text: &str, family: &str, size: f64, bold: bool, italic: bool, x: f64, y: f64, color: Color) {
        TextPages::text(self, text, family, size, bold, italic, x, y, color)
    }
}

pub fn pages(data: &[u8]) -> Result<(PdfDocument, Vec<UnsupportedContent>), String> {
    let mut canvas = TextPages::default();
    let unsupported = oxicells_core::print::pages(data, &mut canvas)?;
    let unsupported = unsupported
        .into_iter()
        .map(|u| UnsupportedContent { kind: u.kind, part: u.part, count: u.count })
        .collect();
    Ok((canvas.finish(), unsupported))
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::{Cursor, Write};
    fn package(sheets: &[(&str, &str, &str)], names: &str) -> Vec<u8> {
        package_with_styles(
            sheets,
            names,
            r#"<styleSheet><fonts count="1"><font><sz val="11"/><name val="Arial"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellXfs count="2"><xf fontId="0" fillId="0" borderId="0" numFmtId="0"/><xf fontId="0" fillId="0" borderId="0" numFmtId="4"/></cellXfs></styleSheet>"#,
        )
    }
    fn package_with_styles(sheets: &[(&str, &str, &str)], names: &str, styles: &str) -> Vec<u8> {
        let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
        let mut part = |name: &str, value: &str| {
            zip.start_file(name, zip::write::SimpleFileOptions::default())
                .unwrap();
            zip.write_all(value.as_bytes()).unwrap();
        };
        part("[Content_Types].xml", "<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Default Extension='xml' ContentType='application/xml'/></Types>");
        let mut book=String::from("<workbook xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'><sheets>");
        let mut rels = String::from("<Relationships>");
        for (i, (name, state, xml)) in sheets.iter().enumerate() {
            book.push_str(&format!(
                "<sheet name='{name}' sheetId='{}' r:id='r{}' state='{state}'/>",
                i + 1,
                i + 1
            ));
            rels.push_str(&format!("<Relationship Id='r{}' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet' Target='worksheets/sheet{}.xml'/>",i+1,i+1));
            part(&format!("xl/worksheets/sheet{}.xml", i + 1), xml);
        }
        book.push_str(&format!(
            "</sheets><definedNames>{names}</definedNames></workbook>"
        ));
        rels.push_str("<Relationship Id='styles' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles' Target='styles.xml'/></Relationships>");
        part("xl/workbook.xml", &book);
        part("xl/_rels/workbook.xml.rels", &rels);
        part("xl/styles.xml", styles);
        zip.finish().unwrap().into_inner()
    }
    fn strings(doc: &PdfDocument) -> Vec<String> {
        doc.pages
            .iter()
            .map(|p| {
                p.contents
                    .iter()
                    .filter_map(|e| e.text_span())
                    .map(|s| s.text.as_str())
                    .collect()
            })
            .collect()
    }

    #[test]
    fn xlsx_pdf_uses_cached_values_and_skips_hidden_sheets() {
        let xml = r#"<worksheet><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>日本語</t></is></c><c r="B1" s="1"><f>1+2</f><v>1234.5</v></c></row></sheetData></worksheet>"#;
        let bytes = package(
            &[
                ("Visible", "visible", xml),
                ("Hidden", "hidden", xml),
                ("Last", "visible", xml),
            ],
            "",
        );
        let (doc, unsupported) = pages(&bytes).unwrap();
        assert_eq!(doc.pages.len(), 2);
        assert!(unsupported.is_empty(), "{unsupported:?}");
        assert!(strings(&doc)[0].contains("1,234.50"));
        assert!(!doc.embedded_fonts.is_empty());
        assert!(oxipdf_core::write_pdf(&doc).starts_with(b"%PDF-"));
    }
    #[test]
    fn xlsx_pdf_respects_print_area_landscape_and_manual_breaks() {
        let xml = r#"<worksheet><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>OUTSIDE</t></is></c></row><row r="2"><c r="B2" t="inlineStr"><is><t>FIRST</t></is></c></row><row r="3"><c r="B3" t="inlineStr"><is><t>SECOND</t></is></c></row></sheetData><printOptions horizontalCentered="1"/><pageMargins left="0.5" right="0.5" top="0.5" bottom="0.5"/><pageSetup paperSize="9" orientation="landscape"/><rowBreaks><brk id="2" man="1"/></rowBreaks></worksheet>"#;
        let names = r#"<definedName name="_xlnm.Print_Area" localSheetId="0">'Printed'!$B$2:$C$3</definedName>"#;
        let bytes = package(&[("Printed", "visible", xml)], names);
        let (doc, unsupported) = pages(&bytes).unwrap();
        assert!(unsupported.is_empty(), "{unsupported:?}");
        assert_eq!(doc.pages.len(), 2);
        assert!((doc.pages[0].width - 841.88976).abs() < 0.01);
        assert_eq!(strings(&doc), ["FIRST", "SECOND"]);
    }
    #[test]
    fn xlsx_pdf_fits_large_sheet_and_keeps_multiple_print_areas() {
        let xml = r#"<worksheet><sheetPr><pageSetUpPr fitToPage="1"/></sheetPr><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>FIRST</t></is></c></row><row r="100"><c r="Z100" t="inlineStr"><is><t>LAST</t></is></c></row></sheetData><pageSetup fitToWidth="1" fitToHeight="1"/></worksheet>"#;
        let bytes = package(&[("Printed", "visible", xml)], "");
        let (doc, _) = pages(&bytes).unwrap();
        assert_eq!(doc.pages.len(), 1);
        assert!(strings(&doc)[0].contains("FIRST"));
        assert!(strings(&doc)[0].contains("LAST"));
        let names = r#"<definedName name="_xlnm.Print_Area" localSheetId="0">'Printed'!$A$1,'Printed'!$Z$100</definedName>"#;
        let bytes = package(&[("Printed", "visible", xml)], names);
        let (doc, _) = pages(&bytes).unwrap();
        assert_eq!(strings(&doc), ["FIRST", "LAST"]);
    }
    #[test]
    fn xlsx_pdf_reports_missing_cache_and_unsupported_print_content() {
        let xml = r#"<worksheet><sheetData><row r="1"><c r="A1"><f>NOW()</f></c></row></sheetData><printOptions headings="1"/><headerFooter><oddHeader>Title</oddHeader></headerFooter></worksheet>"#;
        let bytes = package(
            &[("Printed", "visible", xml)],
            r#"<definedName name="_xlnm.Print_Area" localSheetId="0">'Printed'!$A$1</definedName>"#,
        );
        let (_, unsupported) = pages(&bytes).unwrap();
        for name in [
            "print-headings",
            "header-footer",
            "formula-without-cached-value",
        ] {
            assert!(
                unsupported.iter().any(|u| u.kind == name),
                "{unsupported:?}"
            );
        }
    }
    #[test]
    fn xlsx_pdf_hides_rows_columns_and_warns_about_split_merges() {
        let xml = r#"<worksheet><cols><col min="2" max="2" hidden="1" width="10"/></cols><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>VISIBLE</t></is></c><c r="B1" t="inlineStr"><is><t>HIDDEN_COLUMN</t></is></c></row><row r="2" hidden="1"><c r="A2" t="inlineStr"><is><t>HIDDEN_ROW</t></is></c></row><row r="3"><c r="A3" t="inlineStr"><is><t>LAST</t></is></c></row></sheetData><mergeCells><mergeCell ref="A3:C3"/></mergeCells></worksheet>"#;
        let bytes = package(&[("Printed", "visible", xml)], "");
        let (doc, unsupported) = pages(&bytes).unwrap();
        assert!(unsupported.is_empty(), "{unsupported:?}");
        assert_eq!(strings(&doc), ["VISIBLELAST"]);
        let names = r#"<definedName name="_xlnm.Print_Area" localSheetId="0">Printed!$B$3:$C$3</definedName>"#;
        let bytes = package(&[("Printed", "visible", xml)], names);
        let (_, unsupported) = pages(&bytes).unwrap();
        assert!(unsupported
            .iter()
            .any(|u| u.kind == "merge-across-print-boundary"));
    }
    /// Styles with a fill, a dotted bottom border and distributed alignment.
    const FORM_STYLES: &str = r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFCCCC"/></patternFill></fill></fills><borders count="2"><border/><border><bottom style="dotted"><color auto="1"/></bottom></border></borders><cellXfs count="2"><xf fontId="0" fillId="0" borderId="0"/><xf fontId="0" fillId="2" borderId="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="distributed"/></xf></cellXfs></styleSheet>"#;
    fn form_sheet(page_setup: &str) -> String {
        format!(r#"<worksheet><cols><col min="1" max="1" width="30" customWidth="1"/></cols><sheetData><row r="1"><c r="A1" s="1" t="inlineStr"><is><t>金額</t></is></c></row></sheetData>{page_setup}</worksheet>"#)
    }
    fn fills(doc: &PdfDocument) -> usize {
        doc.pages[0].contents.iter().filter(|e| matches!(e, ContentElement::Path(p) if p.fill.is_some())).count()
    }

    #[test]
    fn xlsx_pdf_black_and_white_leaves_fills_out() {
        let colour = package_with_styles(&[("S", "visible", &form_sheet(""))], "", FORM_STYLES);
        let mono = package_with_styles(&[("S", "visible", &form_sheet(r#"<pageSetup blackAndWhite="1"/>"#))], "", FORM_STYLES);
        let (colour, _) = pages(&colour).unwrap();
        let (mono, unsupported) = pages(&mono).unwrap();
        assert_eq!(fills(&colour), 1);
        assert_eq!(fills(&mono), 0);
        assert!(unsupported.is_empty(), "{unsupported:?}");
    }

    #[test]
    fn xlsx_pdf_distributes_text_and_dots_borders() {
        let bytes = package_with_styles(&[("S", "visible", &form_sheet(""))], "", FORM_STYLES);
        let (doc, unsupported) = pages(&bytes).unwrap();
        assert!(unsupported.is_empty(), "{unsupported:?}");
        // 均等割り付け: each character is drawn on its own, far apart across the cell
        let xs: Vec<f64> = doc.pages[0].contents.iter().filter_map(|e| e.text_span()).map(|s| s.x).collect();
        assert_eq!(xs.len(), 2, "{xs:?}");
        assert!(xs[1] - xs[0] > 60., "{xs:?}");
        // a dotted rule is many short segments, not one line
        let dotted = doc.pages[0].contents.iter().filter_map(|e| match e {
            ContentElement::Path(p) if p.stroke.is_some() => Some(p.operations.len()),
            _ => None,
        }).max().unwrap();
        assert!(dotted > 20, "{dotted}");
    }

    #[test]
    fn xlsx_pdf_general_alignment_puts_numbers_right() {
        // horizontal="general" (LibreOffice writes it for the default) must
        // align like no setting: numbers right, text left.
        let styles = r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellXfs count="2"><xf fontId="0" fillId="0" borderId="0"/><xf fontId="0" fillId="0" borderId="0" applyAlignment="1"><alignment horizontal="general"/></xf></cellXfs></styleSheet>"#;
        let xml = r#"<worksheet><sheetData><row r="1"><c r="A1" s="1"><v>7</v></c><c r="B1" s="1" t="inlineStr"><is><t>x</t></is></c></row></sheetData></worksheet>"#;
        let bytes = package_with_styles(&[("Printed", "visible", xml)], "", styles);
        let (doc, _) = pages(&bytes).unwrap();
        let clips: Vec<(f64, f64)> = doc.pages[0].contents.iter().filter_map(|e| match e {
            ContentElement::ClipPath(c) => match (&c.operations[0], &c.operations[1]) {
                (PathOp::MoveTo(x0, _), PathOp::LineTo(x1, _)) => Some((*x0, *x1)),
                _ => None,
            },
            _ => None,
        }).collect();
        let xs: Vec<f64> = doc.pages[0].contents.iter().filter_map(|e| e.text_span()).map(|s| s.x).collect();
        // clips: page, A1, B1; the number sits in the right half of A1, the text in the left half of B1
        let (a0, a1) = clips[1];
        let (b0, b1) = clips[2];
        assert!(xs[0] > (a0 + a1) / 2., "{xs:?} {clips:?}");
        assert!(xs[1] < (b0 + b1) / 2., "{xs:?} {clips:?}");
    }
    #[test]
    fn xlsx_pdf_text_runs_on_over_empty_neighbours() {
        // A1 runs on over the empty B1 and stops at C1; C1 keeps its own width.
        let xml = r#"<worksheet><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>a long title</t></is></c><c r="C1" t="inlineStr"><is><t>X</t></is></c></row></sheetData></worksheet>"#;
        let bytes = package(&[("Printed", "visible", xml)], "");
        let (doc, _) = pages(&bytes).unwrap();
        let widths: Vec<f64> = doc.pages[0]
            .contents
            .iter()
            .filter_map(|e| match e {
                ContentElement::ClipPath(c) => match (&c.operations[0], &c.operations[1]) {
                    (PathOp::MoveTo(x0, _), PathOp::LineTo(x1, _)) => Some(x1 - x0),
                    _ => None,
                },
                _ => None,
            })
            .collect();
        // page clip, then A1 and C1
        assert_eq!(widths.len(), 3, "{widths:?}");
        assert!((widths[1] - 2. * widths[2]).abs() < 0.01, "{widths:?}");
    }
    #[test]
    fn xlsx_pdf_rejects_excessive_print_ranges() {
        let bytes = package(
            &[("Printed", "visible", "<worksheet><sheetData/></worksheet>")],
            r#"<definedName name="_xlnm.Print_Area" localSheetId="0">Printed!$A$1:$XFD$1048576</definedName>"#,
        );
        assert!(pages(&bytes).unwrap_err().contains("cell limit"));
    }
}
