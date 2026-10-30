// SPDX-License-Identifier: MIT OR Apache-2.0

//! Module and CFB stream identities from the Unicode directory records.

pub(crate) struct ModuleIdentity {
    pub declared_name: String,
    pub name: String,
    pub stream_name: String,
    pub source_offset_position: usize,
    pub record_start: usize,
    pub record_end: usize,
}

struct Records<'a> {
    bytes: &'a [u8],
}

impl<'a> Records<'a> {
    fn take(&mut self, count: usize) -> Option<&'a [u8]> {
        let (value, rest) = self.bytes.split_at_checked(count)?;
        self.bytes = rest;
        Some(value)
    }

    fn peek(&self) -> Option<u16> {
        Some(u16::from_le_bytes(self.bytes.get(..2)?.try_into().ok()?))
    }

    fn record(&mut self, expected: u16) -> Option<&'a [u8]> {
        let header = self.take(6)?;
        if u16::from_le_bytes(header[..2].try_into().ok()?) != expected {
            return None;
        }
        let size = u32::from_le_bytes(header[2..].try_into().ok()?) as usize;
        self.take(size)
    }

    fn fixed(&mut self, expected: u16, size: usize) -> Option<&'a [u8]> {
        let value = self.record(expected)?;
        (value.len() == size).then_some(value)
    }
}

fn unicode(bytes: &[u8]) -> Option<String> {
    if bytes.len() % 2 != 0 {
        return None;
    }
    let units: Vec<_> = bytes
        .chunks_exact(2)
        .map(|b| u16::from_le_bytes([b[0], b[1]]))
        .collect();
    let value = String::from_utf16(&units).ok()?;
    (!value.contains('\0')).then_some(value)
}

fn candidate(
    bytes: &[u8],
    project: &ovba::Project,
    source_code_page: Option<u16>,
) -> Option<Vec<ModuleIdentity>> {
    let declared = codepage::to_encoding(project.information.code_page)?;
    let target = codepage::to_encoding(source_code_page.unwrap_or(project.information.code_page))?;
    let decode = |bytes: &[u8]| declared.decode_without_bom_handling(bytes).0.into_owned();
    let mut records = Records { bytes };
    let count = u16::from_le_bytes(records.fixed(0x000f, 2)?.try_into().ok()?) as usize;
    if count != project.modules.len() {
        return None;
    }
    records.fixed(0x0013, 2)?; // PROJECTCOOKIE
    let mut identities = Vec::with_capacity(count);
    for module in &project.modules {
        let record_start = bytes.len() - records.bytes.len();
        let ansi_name = records.record(0x0019)?;
        let declared_name = decode(ansi_name);
        if declared_name != module.name {
            return None;
        }
        let name = if records.peek() == Some(0x0047) {
            unicode(records.record(0x0047)?)?
        } else {
            target
                .decode_without_bom_handling_and_without_replacement(ansi_name)?
                .into_owned()
        };
        let stream_ansi = records.record(0x001a)?;
        if decode(stream_ansi) != module.stream_name {
            return None;
        }
        let stream_name = unicode(records.record(0x0032)?)?;
        if name.is_empty()
            || stream_name.is_empty()
            || matches!(stream_name.as_str(), "." | "..")
            || stream_name.contains(['/', '\\'])
        {
            return None;
        }
        if decode(records.record(0x001c)?) != module.doc_string {
            return None;
        }
        records.record(0x0048)?; // The Unicode doc string is not an identity.
        let source_offset_position = bytes.len() - records.bytes.len() + 6;
        let offset = u32::from_le_bytes(records.fixed(0x0031, 4)?.try_into().ok()?) as usize;
        let help = u32::from_le_bytes(records.fixed(0x001e, 4)?.try_into().ok()?);
        if offset != module.text_offset || help != module.help_context {
            return None;
        }
        records.fixed(0x002c, 2)?;
        let kind = match module.module_type {
            ovba::ModuleType::Procedural => 0x0021,
            ovba::ModuleType::DocClsDesigner => 0x0022,
        };
        records.fixed(kind, 0)?;
        let read_only = records.peek() == Some(0x0025);
        if read_only {
            records.fixed(0x0025, 0)?;
        }
        let private = records.peek() == Some(0x0028);
        if private {
            records.fixed(0x0028, 0)?;
        }
        if read_only != module.read_only || private != module.private {
            return None;
        }
        records.fixed(0x002b, 0)?;
        let record_end = bytes.len() - records.bytes.len();
        identities.push(ModuleIdentity {
            declared_name,
            name,
            stream_name,
            source_offset_position,
            record_start,
            record_end,
        });
    }
    records.fixed(0x0010, 0)?;
    records.bytes.is_empty().then_some(identities)
}

pub(crate) fn identities(
    project: &ovba::Project,
    source_code_page: Option<u16>,
) -> Result<Vec<ModuleIdentity>, String> {
    if let Some(code_page) = source_code_page {
        if codepage::to_encoding(code_page).is_none() {
            return Err(format!("unsupported VBA source code page: {code_page}"));
        }
    }
    let path = std::path::Path::new("/VBA").join("dir");
    let bytes = crate::vba_compression::read_stream(project, path, 0)
        .map_err(|e| format!("cannot read VBA directory: {e}"))?;
    from_directory(&bytes, project, source_code_page)
}

fn from_directory(
    bytes: &[u8],
    project: &ovba::Project,
    source_code_page: Option<u16>,
) -> Result<Vec<ModuleIdentity>, String> {
    locate_module_table(bytes, project, source_code_page).map(|(_, modules)| modules)
}

pub(crate) fn module_table_offset(
    bytes: &[u8], project: &ovba::Project, source_code_page: Option<u16>,
) -> Result<usize, String> {
    locate_module_table(bytes, project, source_code_page).map(|(offset, _)| offset)
}

fn locate_module_table(
    bytes: &[u8],
    project: &ovba::Project,
    source_code_page: Option<u16>,
) -> Result<(usize, Vec<ModuleIdentity>), String> {
    let count = u16::try_from(project.modules.len())
        .map_err(|_| "too many VBA modules")?
        .to_le_bytes();
    let marker = [0x0f, 0, 2, 0, 0, 0, count[0], count[1], 0x13, 0, 2, 0, 0, 0];
    let mut found = None;
    let mut attempts = 0;
    // The dependency parses the complete directory but drops Unicode names.
    // Locate the module suffix by validating EVERY descriptor against that
    // parse, and require the final terminator at EOF. Never accept a marker
    // in an opaque reference string as a module table by itself.
    for (offset, window) in bytes.windows(marker.len()).enumerate() {
        if window != marker {
            continue;
        }
        attempts += 1;
        if attempts > 32 {
            return Err("ambiguous VBA module directory".into());
        }
        if let Some(mut value) = candidate(&bytes[offset..], project, source_code_page) {
            if found.is_some() {
                return Err("ambiguous VBA module directory".into());
            }
            for identity in &mut value {
                identity.source_offset_position += offset;
                identity.record_start += offset;
                identity.record_end += offset;
            }
            found = Some((offset, value));
        }
    }
    found.ok_or_else(|| "invalid VBA Unicode module directory".into())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn unicode_directory_checks_complete_records_and_rejects_invalid_utf16() {
        let bytes =
            include_bytes!("../../../tests/fixtures/vba/unicode_modules_unicode_mixed.xlsm");
        let mut archive = oxidocs_common::archive::OoxmlArchive::new(bytes).unwrap();
        let mut project =
            ovba::open_project(archive.read_binary_part("xl/vbaProject.bin").unwrap()).unwrap();
        let dir = project
            .decompress_stream_from(std::path::Path::new("/VBA").join("dir"), 0)
            .unwrap();
        let good = from_directory(&dir, &project, Some(932)).unwrap();
        assert!(good.iter().any(|name| name.name == "帳票"));
        let count = (project.modules.len() as u16).to_le_bytes();
        let marker = [0x0f, 0, 2, 0, 0, 0, count[0], count[1], 0x13, 0, 2, 0, 0, 0];
        let mut decoy = marker.to_vec();
        decoy.extend_from_slice(&[0, 0, 0x19, 0, 0xff, 0xff, 0xff, 0xff]);
        decoy.extend_from_slice(&dir);
        let names: Vec<_> = from_directory(&decoy, &project, Some(932))
            .unwrap()
            .into_iter()
            .map(|name| name.name)
            .collect();
        assert_eq!(
            names,
            good.into_iter().map(|name| name.name).collect::<Vec<_>>()
        );
        let start = dir
            .windows(marker.len())
            .position(|window| window == marker)
            .unwrap();
        let mut records = Records {
            bytes: &dir[start..],
        };
        records.fixed(0x000f, 2).unwrap();
        records.fixed(0x0013, 2).unwrap();
        records.record(0x0019).unwrap();
        let name = records.record(0x0047).unwrap();
        let position = dir.len() - records.bytes.len() - name.len();
        let mut broken = dir.clone();
        broken[position..position + 2].copy_from_slice(&0xd800_u16.to_le_bytes());
        assert!(from_directory(&broken, &project, Some(932)).is_err());
        let identity = identities(&project, Some(932)).unwrap().remove(0);
        project.modules[0].text_offset = usize::MAX;
        assert!(module_source(&project, &project.modules[0], &identity, Some(932)).is_err());
    }
}

pub(crate) fn module_source(
    project: &ovba::Project,
    module: &ovba::Module,
    identity: &ModuleIdentity,
    source_code_page: Option<u16>,
) -> Result<String, String> {
    let code_page = source_code_page.unwrap_or(project.information.code_page);
    let encoding = codepage::to_encoding(code_page)
        .ok_or_else(|| format!("unsupported VBA source code page: {code_page}"))?;
    let path = std::path::Path::new("/VBA").join(&identity.stream_name);
    let stream_size = project
        .read_stream(&path)
        .map_err(|e| format!("cannot read VBA module {}: {e}", identity.name))?
        .len();
    if module.text_offset >= stream_size {
        return Err(format!(
            "VBA source offset is outside module {}",
            identity.name
        ));
    }
    let bytes = crate::vba_compression::read_stream(project, path, module.text_offset)
        .map_err(|e| format!("cannot read VBA module {}: {e}", identity.name))?;
    if source_code_page.is_some() {
        encoding
            .decode_without_bom_handling_and_without_replacement(&bytes)
            .map(|text| text.into_owned())
            .ok_or_else(|| {
                format!(
                    "invalid VBA source in module {} for code page {code_page}",
                    identity.name
                )
            })
    } else {
        Ok(encoding.decode_without_bom_handling(&bytes).0.into_owned())
    }
}
