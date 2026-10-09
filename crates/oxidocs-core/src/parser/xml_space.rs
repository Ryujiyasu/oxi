// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use std::borrow::Cow;
use quick_xml::{Reader, Writer, events::{BytesStart, Event}};

fn declared_space(element: &BytesStart<'_>) -> Option<bool> {
    element.attributes().flatten().find_map(|a| {
        if a.key.as_ref() != b"xml:space" { return None; }
        match a.value.as_ref() {
            b"preserve" => Some(true),
            b"default" => Some(false),
            _ => None,
        }
    })
}

/// Materialize the measured Word text-part whitespace default.
/// Word uses the XML root's setting and each text element's explicit setting.
/// Container declarations do not alter this default: the twelve Word controls
/// cover body, paragraph, run, hyperlink, table, and explicit text resets.
pub(super) fn materialize(xml: &str) -> Result<Cow<'_, str>, quick_xml::Error> {
    if !xml.contains("xml:space") { return Ok(Cow::Borrowed(xml)); }
    let mut reader = Reader::from_str(xml);
    let mut writer = Writer::new(Vec::with_capacity(xml.len()));
    let mut root_default = None;
    loop {
        match reader.read_event()? {
            Event::Start(mut e) => {
                let preserve = *root_default.get_or_insert_with(|| declared_space(&e).unwrap_or(false));
                if preserve && declared_space(&e).is_none()
                    && matches!(e.local_name().as_ref(), b"t" | b"delText") {
                    e.push_attribute(("xml:space", "preserve"));
                }
                writer.write_event(Event::Start(e))?;
            }
            Event::Empty(mut e) => {
                let preserve = *root_default.get_or_insert_with(|| declared_space(&e).unwrap_or(false));
                if preserve && declared_space(&e).is_none()
                    && matches!(e.local_name().as_ref(), b"t" | b"delText") {
                    e.push_attribute(("xml:space", "preserve"));
                }
                writer.write_event(Event::Empty(e))?;
            }
            Event::Eof => break,
            other => writer.write_event(other)?,
        }
    }
    Ok(Cow::Owned(String::from_utf8(writer.into_inner())
        .expect("UTF-8 XML reader and writer preserve UTF-8")))
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn root_default_survives_container_reset_and_nested_cells() {
        let xml = r#"<w:document xml:space="preserve"><w:body xml:space="default"><w:tbl><w:tr><w:tc><w:p><w:hyperlink><w:r><w:t> </w:t></w:r></w:hyperlink></w:p></w:tc></w:tr></w:tbl></w:body></w:document>"#;
        assert!(materialize(xml).unwrap().contains(r#"<w:t xml:space="preserve"> </w:t>"#));
    }
    #[test]
    fn container_preserve_does_not_change_part_default() {
        let xml = r#"<r><p xml:space="preserve"><t> A </t></p><t> B </t></r>"#;
        assert_eq!(materialize(xml).unwrap(), xml);
    }
    #[test]
    fn explicit_text_reset_and_entities_are_preserved() {
        let xml = r#"<r xml:space="preserve"><t xml:space="default"> &amp; </t><t xml:space="preserve"/></r>"#;
        assert_eq!(materialize(xml).unwrap(), xml);
    }
    #[test]
    fn parts_without_space_declarations_are_borrowed() {
        assert!(matches!(materialize("<r><t> A </t></r>").unwrap(), Cow::Borrowed(_)));
    }
}
