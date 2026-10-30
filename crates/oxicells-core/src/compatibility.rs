// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use quick_xml::{events::Event, name::ResolveResult, NsReader};
const NAMESPACE: &[u8] =
    b"http://schemas.microsoft.com/office/spreadsheetml/2024/workbookCompatibilityVersion";

pub(crate) fn read_version(xml: &str) -> u8 {
    version(xml).unwrap_or(1)
}

fn version(xml: &str) -> Option<u8> {
    let mut reader = NsReader::from_str(xml);
    while let Ok((resolved, event)) = reader.read_resolved_event() {
        match event {
            Event::Start(ref e) | Event::Empty(ref e)
                if matches!(resolved, ResolveResult::Bound(ref namespace) if namespace.as_ref()==NAMESPACE)
                    && e.local_name().as_ref() == b"version" =>
            {
                for attr in e.attributes().flatten() {
                    if attr.key.as_ref() == b"setVersion" {
                        return std::str::from_utf8(&attr.value)
                            .ok()
                            .and_then(|text| text.parse().ok());
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    None
}

pub(crate) fn patch_version(xml: &str, value: u8) -> Result<String, crate::parser::XlsxError> {
    use quick_xml::{
        events::{BytesEnd, BytesStart},
        Writer,
    };
    use std::io::Cursor;
    let had_version = version(xml).is_some();
    if !had_version && value == 1 {
        return Ok(xml.into());
    }
    let mut reader = NsReader::from_str(xml);
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    let mut depth = 0_usize;
    let mut written = false;
    let mut prefix = String::new();
    fn failure(error: impl std::fmt::Display) -> crate::parser::XlsxError {
        crate::parser::XlsxError::InvalidData(error.to_string())
    }
    fn extension(
        writer: &mut Writer<Cursor<Vec<u8>>>,
        prefix: &str,
        value: u8,
    ) -> Result<(), quick_xml::Error> {
        let name = format!("{prefix}ext");
        let mut ext = BytesStart::new(&name);
        ext.push_attribute(("uri", "{D14903EA-33C4-47F7-8F05-3474C54BE107}"));
        ext.push_attribute(("xmlns:xlwcv", std::str::from_utf8(NAMESPACE).unwrap()));
        writer.write_event(Event::Start(ext))?;
        let mut version = BytesStart::new("xlwcv:version");
        let number = value.to_string();
        version.push_attribute(("setVersion", number.as_str()));
        writer.write_event(Event::Empty(version))?;
        writer.write_event(Event::End(BytesEnd::new(name)))?;
        Ok(())
    }
    loop {
        let (resolved, event) = reader.read_resolved_event().map_err(failure)?;
        if let Event::Start(e) = &event {
            if depth == 0 {
                let name = e.name();
                let name = std::str::from_utf8(name.as_ref()).unwrap_or("");
                prefix = name
                    .rsplit_once(':')
                    .map(|(prefix, _)| format!("{prefix}:"))
                    .unwrap_or_default();
            }
            depth += 1;
        }
        if matches!(resolved,ResolveResult::Bound(ref namespace)if namespace.as_ref()==NAMESPACE) {
            if let Event::Start(e) | Event::Empty(e) = &event {
                if e.local_name().as_ref() == b"version" {
                    let name = e.name();
                    let name = std::str::from_utf8(name.as_ref()).unwrap();
                    let mut element = BytesStart::new(name);
                    for attribute in e.attributes().flatten() {
                        if attribute.key.as_ref() != b"setVersion" {
                            element.push_attribute(attribute);
                        }
                    }
                    let number = value.to_string();
                    element.push_attribute(("setVersion", number.as_str()));
                    writer
                        .write_event(if matches!(event, Event::Empty(_)) {
                            Event::Empty(element)
                        } else {
                            Event::Start(element)
                        })
                        .map_err(failure)?;
                    written = true;
                    continue;
                }
            }
        }
        if let Event::End(e) = &event {
            if !had_version && !written && depth == 2 && e.local_name().as_ref() == b"extLst" {
                extension(&mut writer, &prefix, value).map_err(failure)?;
                written = true;
            } else if !had_version
                && !written
                && depth == 1
                && e.local_name().as_ref() == b"workbook"
            {
                let name = format!("{prefix}extLst");
                writer
                    .write_event(Event::Start(BytesStart::new(&name)))
                    .map_err(failure)?;
                extension(&mut writer, &prefix, value).map_err(failure)?;
                writer
                    .write_event(Event::End(BytesEnd::new(name)))
                    .map_err(failure)?;
                written = true;
            }
            depth = depth.saturating_sub(1);
        }
        if let Event::Empty(e) = &event {
            if !had_version && !written && depth == 1 && e.local_name().as_ref() == b"extLst" {
                writer
                    .write_event(Event::Start(e.clone()))
                    .map_err(failure)?;
                extension(&mut writer, &prefix, value).map_err(failure)?;
                writer
                    .write_event(Event::End(BytesEnd::new(
                        std::str::from_utf8(e.name().as_ref()).unwrap(),
                    )))
                    .map_err(failure)?;
                written = true;
                continue;
            }
        }
        if matches!(event, Event::Eof) {
            break;
        }
        writer.write_event(event).map_err(failure)?;
    }
    String::from_utf8(writer.into_inner().into_inner())
        .map_err(|error| crate::parser::XlsxError::InvalidData(error.to_string()))
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn version_extension_roundtrip_preserves_namespaces_and_other_extensions() {
        for xml in [
            "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheets/></workbook>",
            "<s:workbook xmlns:s=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><s:sheets/><s:extLst/></s:workbook>",
            "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheets/><extLst><ext uri=\"keep\"><x:version xmlns:x=\"urn:other\" setVersion=\"99\"/></ext></extLst></workbook>",
        ] {
            assert_eq!(read_version(xml), 1);
            let modern = patch_version(xml, 2).unwrap();assert_eq!(read_version(&modern), 2);
            let legacy = patch_version(&modern, 1).unwrap();assert_eq!(read_version(&legacy), 1);
            if xml.contains("uri=\"keep\"") {assert!(legacy.contains("uri=\"keep\""));assert!(legacy.contains("setVersion=\"99\""));}
        }
    }
}
