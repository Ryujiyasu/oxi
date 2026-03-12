use std::collections::HashMap;

use quick_xml::events::Event;
use quick_xml::reader::Reader;

use super::ParseError;

#[derive(Debug, Clone)]
pub struct Relationship {
    pub id: String,
    pub rel_type: String,
    pub target: String,
}

/// Parse a .rels XML file into a map of Id -> Relationship
pub fn parse_relationships(xml: &str) -> Result<HashMap<String, Relationship>, ParseError> {
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

                    for attr in e.attributes().flatten() {
                        let key = std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                        let val = String::from_utf8_lossy(&attr.value).to_string();
                        match key {
                            "Id" => id = val,
                            "Type" => rel_type = val,
                            "Target" => target = val,
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

fn local_name(name: &[u8]) -> String {
    let s = std::str::from_utf8(name).unwrap_or("");
    match s.rfind(':') {
        Some(pos) => s[pos + 1..].to_string(),
        None => s.to_string(),
    }
}
