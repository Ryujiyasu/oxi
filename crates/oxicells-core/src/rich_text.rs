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
        if let Some(font) = &format.font {
            property(writer, prefix, "rFont", "val", font)?;
        }
        if let Some(size) = format.size {
            property(writer, prefix, "sz", "val", &size.to_string())?;
        }
        for (name, set) in [
            ("b", format.bold),
            ("i", format.italic),
            ("u", format.underline),
        ] {
            if set {
                property(writer, prefix, name, "", "")?;
            }
        }
        if let Some(color) = &format.color {
            let color = color.trim_start_matches('#');
            let color = if color.len() == 6 {
                format!("FF{color}")
            } else {
                color.to_owned()
            };
            property(writer, prefix, "color", "rgb", &color)?;
        }
        if let Some(align) = &format.vert_align {
            property(writer, prefix, "vertAlign", "val", align)?;
        }
        writer
            .write_event(Event::End(BytesEnd::new(name)))
            .map_err(failure)?;
    }
    let name = format!("{prefix}t");
    let mut text = BytesStart::new(&name);
    text.push_attribute(("xml:space", "preserve"));
    writer.write_event(Event::Start(text)).map_err(failure)?;
    let escaped = crate::string_value::encode(units);
    writer
        .write_event(Event::Text(BytesText::new(&escaped)))
        .map_err(failure)?;
    writer
        .write_event(Event::End(BytesEnd::new(name)))
        .map_err(failure)?;
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
