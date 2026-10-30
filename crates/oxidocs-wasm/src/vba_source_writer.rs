// SPDX-License-Identifier: MIT OR Apache-2.0

//! Update source in existing VBA modules while retaining project identities.

use crate::vba_module_directory::{identities, module_source, module_table_offset};
use std::collections::{BTreeMap, BTreeSet, HashMap};
use std::io::{Cursor, Write};

/// MS-OVBA 2.4.1: chunks decompress to 4096 bytes except the final chunk.
/// Raw fallback, including final zero padding, follows section 2.4.1.3.10.
pub(crate) fn compress(bytes: &[u8]) -> Vec<u8> {
    let mut output = vec![1];
    for chunk in bytes.chunks(4096) {
        let mut encoded = Vec::new();
        let mut positions: HashMap<[u8; 3], Vec<usize>> = HashMap::new();
        let mut at = 0usize;
        while at < chunk.len() && encoded.len() <= 4096 {
            let flag_at = encoded.len();
            encoded.push(0);
            for bit in 0..8 {
                if at == chunk.len() {
                    break;
                }
                let offset_bits = if at <= 16 {
                    4
                } else {
                    usize::BITS - (at - 1).leading_zeros()
                };
                let length_bits = 16 - offset_bits;
                let max_length = ((1usize << length_bits) + 2).min(chunk.len() - at);
                let mut best = (0usize, 0usize);
                if at + 3 <= chunk.len() {
                    let key: [u8; 3] = chunk[at..at + 3].try_into().unwrap();
                    if let Some(candidates) = positions.get(&key) {
                        // Bounded work for long comments and repetitive source.
                        for &previous in candidates.iter().rev().take(64) {
                            let mut length = 3;
                            while length < max_length
                                && chunk[previous + length] == chunk[at + length]
                            {
                                length += 1;
                            }
                            if length > best.1 {
                                best = (at - previous, length);
                            }
                            if length == max_length {
                                break;
                            }
                        }
                    }
                }
                let length = if best.1 >= 3 {
                    encoded[flag_at] |= 1 << bit;
                    let token = (((best.0 - 1) << length_bits) | (best.1 - 3)) as u16;
                    encoded.extend_from_slice(&token.to_le_bytes());
                    best.1
                } else {
                    encoded.push(chunk[at]);
                    1
                };
                for position in at..at + length {
                    if position + 3 <= chunk.len() {
                        positions
                            .entry(chunk[position..position + 3].try_into().unwrap())
                            .or_default()
                            .push(position);
                    }
                }
                at += length;
            }
        }
        if at == chunk.len() && encoded.len() <= 4096 {
            let header = 0xb000 | (encoded.len() as u16 - 1);
            output.extend_from_slice(&header.to_le_bytes());
            output.extend_from_slice(&encoded);
        } else {
            output.extend_from_slice(&0x3fffu16.to_le_bytes());
            output.extend_from_slice(chunk);
            output.resize(output.len() + 4096 - chunk.len(), 0);
        }
    }
    output
}

fn compress_source(bytes: &[u8]) -> Vec<u8> {
    // Keep literal blocks below the native stored-size limit. Source streams
    // must avoid raw blocks and zero padding, including incompressible text.
    let mut output = vec![1];
    for chunk in bytes.chunks(3600) {
        let mut body = Vec::new();
        for literal in chunk.chunks(8) {
            body.push(0);
            body.extend_from_slice(literal);
        }
        let size = body.len();
        output.extend_from_slice(&(0xb000 | (size as u16-1)).to_le_bytes());
        output.extend_from_slice(&body);
    }
    output
}

fn source_body(text: &str, class: bool) -> Result<String, String> {
    let text = text.replace("\r\n", "\n");
    let mut lines = text.lines();
    let mut first = lines.next().unwrap_or("");
    if class && first.trim().to_ascii_uppercase().starts_with("VERSION ") {
        first = lines.next().unwrap_or("");
        if first.trim().eq_ignore_ascii_case("BEGIN") {
            loop {
                let line = lines.next().ok_or("unterminated exported class header")?;
                if line.trim().eq_ignore_ascii_case("END") {
                    break;
                }
            }
            first = lines.next().unwrap_or("");
        }
    }
    Ok(std::iter::once(first).chain(lines).collect::<Vec<_>>().join("\n"))
}

fn comparable(text: &str) -> String {
    text.replace("\r\n", "\n").trim_end_matches(['\n', '\0']).to_owned()
}

fn module_record(output: &mut Vec<u8>, tag: u16, bytes: &[u8]) -> Result<(), String> {
    let length = u32::try_from(bytes.len()).map_err(|_| "VBA module record is too large")?;
    output.extend_from_slice(&tag.to_le_bytes());
    output.extend_from_slice(&length.to_le_bytes());
    output.extend_from_slice(bytes);
    Ok(())
}

fn unicode_name(name: &str) -> Vec<u8> {
    name.encode_utf16().flat_map(u16::to_le_bytes).collect()
}

fn valid_module_name(name: &str) -> bool {
    let mut characters = name.chars();
    name.encode_utf16().count() <= 31
        && characters.next().is_some_and(char::is_alphabetic)
        && characters.all(|c| c.is_alphanumeric() || c == '_')
}

fn filtered_project_name_map(bytes: &[u8], removed: &BTreeSet<String>) -> Result<Vec<u8>, String> {
    if !bytes.ends_with(&[0, 0]) { return Err("invalid VBA module name map terminator".into()); }
    let mut output = Vec::with_capacity(bytes.len());
    let mut at = 0;
    let end = bytes.len() - 2;
    while at < end {
        let ansi_end = bytes[at..end].iter().position(|byte| *byte == 0)
            .map(|offset| at + offset).ok_or("invalid VBA ANSI name map entry")?;
        let wide_start = ansi_end + 1;
        let mut wide_end = wide_start;
        while wide_end + 1 < end && bytes[wide_end..wide_end + 2] != [0, 0] {
            wide_end += 2;
        }
        if wide_end + 1 >= end { return Err("invalid VBA Unicode name map entry".into()); }
        let units = bytes[wide_start..wide_end].chunks_exact(2)
            .map(|pair| u16::from_le_bytes([pair[0], pair[1]])).collect::<Vec<_>>();
        let name = String::from_utf16(&units).map_err(|_| "invalid VBA Unicode module name map")?;
        if !removed.contains(&name.to_ascii_lowercase()) {
            output.extend_from_slice(&bytes[at..wide_end + 2]);
        }
        at = wide_end + 2;
    }
    if at != end { return Err("invalid VBA module name map length".into()); }
    output.extend_from_slice(&[0, 0]);
    Ok(output)
}

/// Update an existing project's source. Document modules can be omitted as
/// they retain their native identity; omitted ordinary modules are removed.
pub(crate) fn update_project(project_bytes: &[u8], source: &str, source_code_page: Option<u16>) -> Result<Vec<u8>, String> {
    if source.trim().is_empty() {
        return Err("VBA project source is empty".into());
    }
    if source.contains('\0') {
        return Err("VBA source contains a NUL character".into());
    }
    let project = ovba::open_project(project_bytes.to_vec())
        .map_err(|error| format!("cannot read VBA project: {error}"))?;
    let names = identities(&project, source_code_page)?;
    let declared_encoding = codepage::to_encoding(project.information.code_page)
        .ok_or_else(|| format!("unsupported VBA code page {}", project.information.code_page))?;
    let source_code_page_value = source_code_page.unwrap_or(project.information.code_page);
    let encoding = codepage::to_encoding(source_code_page_value)
        .ok_or_else(|| format!("unsupported VBA source code page {source_code_page_value}"))?;
    let old_sources: Vec<String> = project.modules.iter().zip(&names)
        .map(|(module, name)| module_source(&project, module, name, source_code_page))
        .collect::<Result<_, _>>()?;
    let project_text = project.read_stream("/PROJECT").map_err(|e|e.to_string())?;
    let project_text = declared_encoding.decode_without_bom_handling_and_without_replacement(&project_text)
        .ok_or("invalid VBA PROJECT text")?;
    let documents: BTreeSet<_> = project_text.lines()
        .filter_map(|line|line.strip_prefix("Document=")?.split_once("/&H").map(|(name,_)|name.to_owned()))
        .collect();
    let mut edits = BTreeMap::new();
    let mut seen = BTreeSet::new();
    let mut new_names = BTreeSet::new();
    // Logical name, ordinary class flag, declared-code-page name, source bytes.
    let mut additions: Vec<(String, bool, Vec<u8>, Vec<u8>)> = Vec::new();
    for module in oxivba_core::parser::source_modules(source) {
        let index = if let Some(name) = &module.named {
            let matches: Vec<usize> = names.iter().enumerate()
                .filter(|(_, identity)| identity.name.eq_ignore_ascii_case(name))
                .map(|(index, _)| index).collect();
            if matches.is_empty() {
                if !valid_module_name(name) {
                    return Err(format!("invalid VBA module name: {name}"));
                }
                if !new_names.insert(name.to_ascii_lowercase()) {
                    return Err(format!("duplicate VBA module {name}"));
                }
                let mut text = source_body(&module.text, module.class)?;
                let class_base = "Attribute VB_Base = \"0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}\"";
                if text.lines().any(|line| {
                    line.trim_start().to_ascii_lowercase().starts_with("attribute vb_base")
                        && (!module.class || !line.trim().eq_ignore_ascii_case(class_base))
                }) {
                    return Err(format!("VBA document module creation is not implemented: {name}"));
                }
                if module.class {
                    let mut missing = Vec::new();
                    if !text.lines().any(|line| line.trim_start().to_ascii_lowercase().starts_with("attribute vb_base")) {
                        missing.push(class_base.to_string());
                    }
                    for attribute in ["VB_GlobalNameSpace", "VB_Creatable", "VB_PredeclaredId", "VB_Exposed", "VB_TemplateDerived", "VB_Customizable"] {
                        let prefix = format!("attribute {attribute}").to_ascii_lowercase();
                        if !text.lines().any(|line| line.trim_start().to_ascii_lowercase().starts_with(&prefix)) {
                            missing.push(format!("Attribute {attribute} = False"));
                        }
                    }
                    if !missing.is_empty() {
                        let mut lines = Vec::new();
                        for line in text.lines() {
                            lines.push(line.to_owned());
                            if line.trim_start().to_ascii_lowercase().starts_with("attribute vb_name") {
                                lines.extend(missing.iter().cloned());
                            }
                        }
                        text = lines.join("\n");
                    }
                }
                let text = text.trim_end_matches('\n').replace('\n', "\r\n") + "\r\n";
                let (encoded, _, replaced) = encoding.encode(&text);
                if replaced {
                    return Err(format!("VBA source in {name} cannot be represented in code page {source_code_page_value}"));
                }
                let (name_bytes, _, replaced) = declared_encoding.encode(name);
                if replaced {
                    return Err(format!("VBA module name {name} cannot be represented in project code page {}", project.information.code_page));
                }
                additions.push((name.clone(), module.class, name_bytes.into_owned(), encoded.into_owned()));
                continue;
            }
            if matches.len() != 1 {
                return Err(format!("ambiguous VBA module: {name}"));
            }
            matches[0]
        } else {
            let matches: Vec<usize> = project.modules.iter().enumerate()
                .filter(|(_, existing)| matches!(existing.module_type, ovba::ModuleType::Procedural))
                .map(|(index, _)| index).collect();
            if matches.len() != 1 || module.class {
                return Err("VBA source must identify its module with Attribute VB_Name".into());
            }
            matches[0]
        };
        if !seen.insert(index) {
            return Err(format!("duplicate VBA module {}", names[index].name));
        }
        let class = matches!(project.modules[index].module_type, ovba::ModuleType::DocClsDesigner);
        if module.class && !class {
            return Err(format!("VBA module kind changed: {}", names[index].name));
        }
        let mut text = source_body(&module.text, class)?;
        if module.named.is_none() {
            let attributes = old_sources[index].lines()
                .take_while(|line| line.trim().is_empty() || line.trim_start().starts_with("Attribute "))
                .filter(|line| line.trim_start().starts_with("Attribute "))
                .collect::<Vec<_>>().join("\n");
            text = format!("{attributes}\n{text}");
        } else {
            // CodePane/exported class text can omit hidden module metadata.
            // In particular VB_Base is required to rebuild a source-only class.
            let attribute_key = |line: &str| -> Option<String> {
                let rest = line.trim().strip_prefix("Attribute ")?;
                let key = rest.split_once('=')?.0.trim();
                (!key.contains('.')).then(|| key.to_ascii_lowercase())
            };
            let present: BTreeSet<_> = text.lines().filter_map(attribute_key).collect();
            let missing: Vec<_> = old_sources[index].lines()
                .take_while(|line| line.trim().is_empty() || line.trim_start().starts_with("Attribute "))
                .filter(|line| attribute_key(line).is_some_and(|key| !present.contains(&key)))
                .collect();
            if !missing.is_empty() {
                let mut lines = Vec::new();
                for line in text.lines() {
                    lines.push(line.to_owned());
                    if attribute_key(line).is_some_and(|key| key=="vb_name") {
                        lines.extend(missing.iter().map(|line| (*line).to_owned()));
                    }
                }
                text = lines.join("\n");
            }
        }
        if comparable(&text) != comparable(&old_sources[index]) {
            let text = text.trim_end_matches('\n').replace('\n', "\r\n") + "\r\n";
            let (encoded, _, replaced) = encoding.encode(&text);
            if replaced {
                return Err(format!("VBA source in {} cannot be represented in code page {}", names[index].name, source_code_page_value));
            }
            edits.insert(index, encoded.into_owned());
        }
    }
    let removed: BTreeSet<usize> = project.modules.iter().enumerate()
        .filter(|(index, module)| !seen.contains(index) && !documents.contains(&module.name))
        .map(|(index, _)| index).collect();
    if edits.is_empty() && additions.is_empty() && removed.is_empty() {
        return Ok(project_bytes.to_vec());
    }
    let mut directory = crate::vba_compression::read_stream(&project, "/VBA/dir", 0)
        .map_err(|error| format!("cannot read VBA directory: {error}"))?;
    let table_offset = if additions.is_empty() && removed.is_empty() { None }
        else { Some(module_table_offset(&directory, &project, source_code_page)?) };
    let mut sources = Vec::with_capacity(names.len() - removed.len() + additions.len() + 4);
    let mut removed_streams = Vec::new();
    for (index, (module, name)) in project.modules.iter().zip(&names).enumerate() {
        let path = format!("/VBA/{}", name.stream_name);
        if removed.contains(&index) {
            removed_streams.push(path);
            continue;
        }
        let bytes = if let Some(edited) = edits.remove(&index) {
            edited
        } else {
            crate::vba_compression::read_stream(&project, &path, module.text_offset)
                .map_err(|error| format!("cannot read VBA source {}: {error}", name.name))?
        };
        sources.push((path, compress_source(&bytes)));
        directory.get_mut(name.source_offset_position..name.source_offset_position + 4)
            .ok_or("invalid VBA source offset field")?.fill(0);
    }
    if let Some(table) = table_offset {
        let total = u16::try_from(names.len() - removed.len() + additions.len()).map_err(|_| "too many VBA modules")?;
        directory.get_mut(table + 6..table + 8).ok_or("invalid VBA module count")?
            .copy_from_slice(&total.to_le_bytes());
        if !directory.ends_with(&[0x10, 0, 0, 0, 0, 0]) {
            return Err("invalid VBA module table terminator".into());
        }
        let mut rebuilt = directory.get(..table + 16).ok_or("invalid VBA module table header")?.to_vec();
        for (index, name) in names.iter().enumerate() {
            if !removed.contains(&index) {
                rebuilt.extend_from_slice(directory.get(name.record_start..name.record_end).ok_or("invalid VBA module record bounds")?);
            }
        }
        directory = rebuilt;
        let old_map = project.read_stream("/PROJECTwm")
            .map_err(|error| format!("cannot read VBA module name map: {error}"))?;
        let removed_names: BTreeSet<String> = removed.iter().map(|index| names[*index].name.to_ascii_lowercase()).collect();
        let mut map = filtered_project_name_map(&old_map, &removed_names)?;
        map.truncate(map.len() - 2);
        let mut properties = project_text.to_string();
        if !removed_names.is_empty() {
            properties = properties.split_inclusive('\n').filter(|line| {
                let Some((kind, name)) = line.trim_end_matches(['\r', '\n']).split_once('=') else { return true; };
                !matches!(kind, "Module" | "Class") || !removed_names.contains(&name.to_ascii_lowercase())
            }).collect();
        }
        let mut insertion = None;
        let mut position = 0;
        for line in properties.split_inclusive('\n') {
            position += line.len();
            if ["Document=", "Module=", "Class=", "BaseClass="].iter()
                .any(|prefix| line.starts_with(prefix))
            {
                insertion = Some(position);
            }
        }
        let insertion = insertion.ok_or("VBA PROJECT has no module declarations")?;
        let mut added_properties = String::new();
        let mut streams: BTreeSet<String> = names.iter().map(|name| name.stream_name.to_ascii_lowercase()).collect();
        let mut serial = names.len();
        for (name, class, encoded_name, bytes) in additions {
            let stream_name = loop {
                serial += 1;
                let candidate = format!("OxiModule{serial}");
                if streams.insert(candidate.to_ascii_lowercase())
                    && project.read_stream(format!("/VBA/{candidate}")).is_err()
                {
                    break candidate;
                }
            };
            let wide_name = unicode_name(&name);
            module_record(&mut directory, 0x0019, &encoded_name)?;
            module_record(&mut directory, 0x0047, &wide_name)?;
            module_record(&mut directory, 0x001a, stream_name.as_bytes())?;
            module_record(&mut directory, 0x0032, &unicode_name(&stream_name))?;
            module_record(&mut directory, 0x001c, &[])?;
            module_record(&mut directory, 0x0048, &[])?;
            module_record(&mut directory, 0x0031, &0u32.to_le_bytes())?;
            module_record(&mut directory, 0x001e, &0u32.to_le_bytes())?;
            module_record(&mut directory, 0x002c, &0xffffu16.to_le_bytes())?;
            module_record(&mut directory, if class { 0x0022 } else { 0x0021 }, &[])?;
            module_record(&mut directory, 0x002b, &[])?;
            map.extend_from_slice(&encoded_name);
            map.push(0);
            map.extend_from_slice(&wide_name);
            map.extend_from_slice(&[0, 0]);
            added_properties.push_str(if class { "Class=" } else { "Module=" });
            added_properties.push_str(&name);
            added_properties.push_str("\r\n");
            sources.push((format!("/VBA/{stream_name}"), compress_source(&bytes)));
        }
        module_record(&mut directory, 0x0010, &[])?;
        map.extend_from_slice(&[0, 0]);
        properties.insert_str(insertion, &added_properties);
        let (properties, _, replaced) = declared_encoding.encode(&properties);
        if replaced { return Err("VBA PROJECT declarations cannot be represented in its code page".into()); }
        sources.push(("/PROJECT".into(), properties.into_owned()));
        sources.push(("/PROJECTwm".into(), map));
    }
    let mut header = project.read_stream("/VBA/_VBA_PROJECT")
        .map_err(|error| format!("cannot read VBA project version: {error}"))?;
    if header.len() < 7 || header[..2] != [0xcc, 0x61] {
        return Err("invalid VBA project version header".into());
    }
    header.truncate(7);
    header[2..4].copy_from_slice(&0xffffu16.to_le_bytes());
    header[4] = 0;
    drop(project);
    let mut compound = cfb::CompoundFile::open(Cursor::new(project_bytes.to_vec()))
        .map_err(|error| format!("cannot open VBA compound file: {error}"))?;
    let caches: Vec<_> = compound.walk().filter(|entry| entry.is_stream())
        .filter(|entry| entry.path().parent().and_then(|p|p.to_str()).is_some_and(|p|p.eq_ignore_ascii_case("/VBA")))
        .filter(|entry| entry.name().to_ascii_uppercase().starts_with("__SRP_"))
        .map(|entry| entry.path().to_path_buf()).collect();
    for path in caches {
        compound.remove_stream(path).map_err(|error| error.to_string())?;
    }
    for path in removed_streams {
        compound.remove_stream(path).map_err(|error| error.to_string())?;
    }
    sources.push(("/VBA/dir".into(), compress(&directory)));
    sources.push(("/VBA/_VBA_PROJECT".into(), header));
    for (path, bytes) in sources {
        compound.create_stream(path).map_err(|error| error.to_string())?
            .write_all(&bytes).map_err(|error| error.to_string())?;
    }
    Ok(compound.into_inner().into_inner())
}
