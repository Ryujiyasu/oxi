// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Write rich cell content using authoritative value units and run formatting.
use crate::{editor::CellEditValue, ir::TextRun, parser::XlsxError};
use quick_xml::{
    events::{BytesEnd, BytesStart, BytesText, Event},
    Reader, Writer,
};
use std::{collections::HashMap, io::Cursor};

fn failure(error: impl std::fmt::Display) -> XlsxError {
    XlsxError::InvalidData(error.to_string())
}
fn local(name: &[u8]) -> &[u8] {
    name.rsplit(|b| *b == b':').next().unwrap_or(name)
}
fn property(
    writer: &mut Writer<Cursor<Vec<u8>>>,
    prefix: &str,
    name: &str,
    key: &str,
    value: &str,
) -> Result<(), XlsxError> {
    let name = format!("{prefix}{name}");
    let mut e = BytesStart::new(name);
    if !key.is_empty() {
        e.push_attribute((key, value));
    }
    writer.write_event(Event::Empty(e)).map_err(failure)
}
fn text(
    writer: &mut Writer<Cursor<Vec<u8>>>,
    prefix: &str,
    units: &[u16],
) -> Result<(), XlsxError> {
    let name = format!("{prefix}t");
    let mut text = BytesStart::new(&name);
    if units.first().is_some_and(|v|matches!(*v,9|10|13|32)) || units.last().is_some_and(|v|matches!(*v,9|10|13|32)) {
        text.push_attribute(("xml:space","preserve"));
    }
    writer.write_event(Event::Start(text)).map_err(failure)?;
    let escaped = crate::string_value::encode(units);
    writer
        .write_event(Event::Text(BytesText::new(&escaped)))
        .map_err(failure)?;
    writer
        .write_event(Event::End(BytesEnd::new(name)))
        .map_err(failure)?;
    Ok(())
}
fn run(
    writer: &mut Writer<Cursor<Vec<u8>>>,
    prefix: &str,
    units: &[u16],
    format: Option<&TextRun>,
) -> Result<(), XlsxError> {
    let r = format!("{prefix}r");
    writer
        .write_event(Event::Start(BytesStart::new(&r)))
        .map_err(failure)?;
    if let Some(format) = format.filter(|format| format.dressed) {
        let name = format!("{prefix}rPr");
        writer
            .write_event(Event::Start(BytesStart::new(&name)))
            .map_err(failure)?;
        for (name,set) in [("b",format.bold),("i",format.italic),("u",format.underline)] {
            if set {property(writer,prefix,name,"","")?;}
        }
        if let Some(size)=format.size {property(writer,prefix,"sz","val",&size.to_string())?;}
        let reference=format.font_reference.as_ref();
        let theme=reference.and_then(|r|r.theme).map(|v|v.to_string());
        let indexed=reference.and_then(|r|r.indexed).map(|v|v.to_string());
        let tint=reference.and_then(|r|r.tint).filter(|v|*v!=0.0).map(|v|v.to_string());
        let automatic=reference.is_some_and(|r|r.automatic);
        let rgb=format.color.as_ref().map(|value| {
            let value=value.trim_start_matches('#');
            if value.len()==6 {format!("FF{value}")} else {value.to_string()}
        });
        if theme.is_some() || indexed.is_some() || automatic || rgb.is_some() {
            let mut element=BytesStart::new(format!("{prefix}color"));
            if let Some(value)=theme.as_deref() {element.push_attribute(("theme",value));}
            else if let Some(value)=indexed.as_deref() {element.push_attribute(("indexed",value));}
            else if automatic {element.push_attribute(("auto","1"));}
            else if let Some(value)=rgb.as_deref() {element.push_attribute(("rgb",value));}
            if let Some(value)=tint.as_deref() {element.push_attribute(("tint",value));}
            writer.write_event(Event::Empty(element)).map_err(failure)?;
        }
        if let Some(font)=reference.and_then(|r|r.latin_name.as_ref()).or(format.font.as_ref()) {
            property(writer,prefix,"rFont","val",font)?;
        }
        if let Some(value)=format.font_family {property(writer,prefix,"family","val",&value.to_string())?;}
        if let Some(value)=format.font_charset.filter(|v|*v!=0) {property(writer,prefix,"charset","val",&value.to_string())?;}
        if let Some(value)=reference.and_then(|r|r.scheme.as_deref()).filter(|v|*v!="none") {
            property(writer,prefix,"scheme","val",value)?;
        }
        if let Some(align)=&format.vert_align {property(writer,prefix,"vertAlign","val",align)?;}
        writer
            .write_event(Event::End(BytesEnd::new(name)))
            .map_err(failure)?;
    }
    text(writer,prefix,units)?;
    writer
        .write_event(Event::End(BytesEnd::new(r)))
        .map_err(failure)
}
fn payload(
    writer: &mut Writer<Cursor<Vec<u8>>>,
    prefix: &str,
    value: &crate::ir::CellValue,
    runs: &[TextRun],
) -> Result<(), XlsxError> {
    let units = value
        .string_units()
        .ok_or_else(|| failure("rich text needs a string value"))?;
    let name = format!("{prefix}is");
    writer
        .write_event(Event::Start(BytesStart::new(&name)))
        .map_err(failure)?;
    if runs.is_empty() {
        text(writer,prefix,&units)?;
        writer.write_event(Event::End(BytesEnd::new(name))).map_err(failure)?;
        return Ok(());
    }
    let mut at = 0;
    for format in runs {
        let start = at;
        for display in format.text.encode_utf16() {
            if at >= units.len() {
                break;
            }
            // Parsed display runs normalize CRLF; value units retain both.
            at += if display == 10 && units[at] == 13 && units.get(at + 1) == Some(&10) {
                2
            } else {
                1
            };
        }
        if at > start {
            run(writer, prefix, &units[start..at], Some(format))?;
        }
    }
    if at < units.len() {
        run(writer, prefix, &units[at..], None)?;
    }
    writer
        .write_event(Event::End(BytesEnd::new(name)))
        .map_err(failure)
}

pub(crate) fn patch(
    xml: &str,
    edits: &HashMap<(u32, u32), &CellEditValue>,
) -> Result<String, XlsxError> {
    if !edits
        .values()
        .any(|value| matches!(value, CellEditValue::RichText { .. }))
    {
        return Ok(xml.into());
    }
    let mut reader = Reader::from_str(xml);
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    let mut target = false;
    let mut skip = 0_usize;
    loop {
        let event = reader.read_event().map_err(failure)?;
        if skip > 0 {
            match event {
                Event::Start(_) => skip += 1,
                Event::End(_) => skip -= 1,
                _ => {}
            }
            continue;
        }
        if let Event::Start(e) | Event::Empty(e) = &event {
            let name = e.name();
            let name = name.as_ref();
            if local(name) == b"c" {
                let reference = e
                    .attributes()
                    .flatten()
                    .find(|a| a.key.as_ref() == b"r")
                    .map(|a| String::from_utf8_lossy(&a.value).into_owned())
                    .unwrap_or_default();
                let (col, row) = crate::parser::parse_cell_ref(&reference);
                if let Some(CellEditValue::RichText { value, runs }) = edits.get(&(row + 1, col)) {
                    let name = std::str::from_utf8(name).map_err(failure)?;
                    let prefix = name
                        .rsplit_once(':')
                        .map(|(prefix, _)| format!("{prefix}:"))
                        .unwrap_or_default();
                    let mut cell = BytesStart::new(name);
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() != b"t" {
                            cell.push_attribute(attr);
                        }
                    }
                    cell.push_attribute(("t", "inlineStr"));
                    writer.write_event(Event::Start(cell)).map_err(failure)?;
                    payload(&mut writer, &prefix, value, runs)?;
                    if matches!(event, Event::Empty(_)) {
                        writer
                            .write_event(Event::End(BytesEnd::new(name)))
                            .map_err(failure)?;
                    } else {
                        target = true;
                    }
                    continue;
                }
            }
            if target && matches!(local(name), b"v" | b"is") {
                if matches!(event, Event::Start(_)) {
                    skip = 1;
                }
                continue;
            }
        }
        if let Event::End(e) = &event {
            if local(e.name().as_ref()) == b"c" {
                target = false;
            }
        }
        if matches!(event, Event::Eof) {
            break;
        }
        writer.write_event(event).map_err(failure)?;
    }
    String::from_utf8(writer.into_inner().into_inner()).map_err(failure)
}

#[cfg(test)]
mod font_roundtrip_tests {
    use crate::{editor::XlsxEditor, ir::TextRun, parse_xlsx_preserving_values};

    #[test]
    fn run_font_metadata_roundtrip_preserves_split_surrogates() {
        let seed=include_bytes!("../../../tests/fixtures/vba/array_argument_rank.xlsm");
        let mut editor=XlsxEditor::new(seed).unwrap();
        let mut edited=editor.workbook().clone();
        let cell=edited.sheets[0].rows.iter_mut().find(|r|r.index==1).unwrap()
            .cells.iter_mut().find(|c|c.col==0).unwrap();
        let style=cell.style.clone();
        let original=TextRun {size:style.font_size,font:style.font_name.clone(),
            bold:style.bold,italic:style.italic,underline:style.underline,color:style.font_color.clone(),
            dressed:true,font_reference:style.font_reference.clone(),font_family:style.font_family,
            font_charset:style.font_charset,..Default::default()};
        let mut painted=original.clone();
        painted.text="\u{FFFD}".into();painted.color=Some("FF0000".into());
        let reference=painted.font_reference.get_or_insert_with(Default::default);
        reference.theme=None;reference.indexed=None;reference.automatic=false;reference.tint=None;
        let mut tail=original;tail.text="b".into();
        cell.runs=vec![TextRun {text:"a\u{FFFD}".into(),..Default::default()},painted,tail];
        editor.apply_workbook(&edited).unwrap();
        let saved=editor.save().unwrap();
        let mut archive=oxidocs_common::archive::OoxmlArchive::new(&saved).unwrap();
        let xml=archive.read_part("xl/worksheets/sheet1.xml").unwrap();
        assert!(xml.contains("a_xD83D_"));
        assert!(xml.contains("_xDE00_"));
        let reloaded=parse_xlsx_preserving_values(&saved).unwrap();
        let cell=reloaded.sheets[0].rows.iter().find(|r|r.index==1).unwrap()
            .cells.iter().find(|c|c.col==0).unwrap();
        assert_eq!(cell.value.string_units().unwrap().as_ref(),&[97,0xD83D,0xDE00,98]);
        assert_eq!(cell.runs.len(),3);
        assert!(!cell.runs[0].dressed);
        let low=&cell.runs[1];
        assert_eq!(low.text,"\u{FFFD}");
        assert_eq!(low.color.as_deref(),Some("FF0000"));
        assert_eq!(low.font_family,Some(2));
        let reference=low.font_reference.as_ref().unwrap();
        assert_eq!(reference.theme,None);
        assert_eq!(reference.scheme.as_deref(),Some("minor"));
        assert_eq!(reference.latin_name.as_deref().or(low.font.as_deref()),Some("Aptos Narrow"));
        assert_eq!(cell.runs[2].font_reference.as_ref().unwrap().theme,Some(1));
    }
}
