// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use std::collections::HashMap;

use quick_xml::events::Event;
use quick_xml::reader::Reader;

use crate::xml_utils::local_name;
use crate::OxiError;

#[derive(Debug, Clone)]
pub struct Relationship {
    pub id: String,
    pub rel_type: String,
    pub target: String,
    pub target_mode: Option<String>,
}

/// Parse a .rels XML file into a map of Id -> Relationship
pub fn parse_relationships(xml: &str) -> Result<HashMap<String, Relationship>, OxiError> {
    let mut reader = Reader::from_str(xml);
    let mut rels = HashMap::new();

    loop {
        match reader.read_event()? {
            Event::Empty(e) | Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                if local == "Relationship" {
                    let mut id = String::new();
                    let mut rel_type = String::new();
                    let mut target = String::new();
                    let mut target_mode = None;

                    for attr in e.attributes().flatten() {
                        let key = std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                        let val = attr.decode_and_unescape_value(reader.decoder())?.into_owned();
                        match key {
                            "Id" => id = val,
                            "Type" => rel_type = val,
                            "Target" => target = val,
                            "TargetMode" => target_mode = Some(val),
                            _ => {}
                        }
                    }

                    if !id.is_empty() {
                        rels.insert(
                            id.clone(),
                            Relationship {
                                id,
                                rel_type,
                                target,
                                target_mode,
                            },
                        );
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(rels)
}

#[cfg(test)]
mod tests {
    use super::parse_relationships;

    #[test]
    fn relationship_target_mode_is_preserved_for_external_ole_links() {
        let relationships = parse_relationships(r#"<Relationships><Relationship Id="rId1" Type="oleObject" Target="embedded.bin"/><Relationship Id="rId2" Type="oleObject" Target="C:/linked.bin" TargetMode="External"/></Relationships>"#).unwrap();
        assert_eq!(relationships["rId1"].target_mode, None);
        assert_eq!(relationships["rId2"].target_mode.as_deref(), Some("External"));
    }
}
