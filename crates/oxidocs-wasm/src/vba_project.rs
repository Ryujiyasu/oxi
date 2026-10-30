// SPDX-License-Identifier: MIT OR Apache-2.0

//! VBA project metadata from an OOXML workbook, without running its macros.

use crate::vba_module_directory::{identities, module_source};
use oxidocs_common::archive::OoxmlArchive;
use oxidocs_common::relationships::parse_relationships;
use wasm_bindgen::prelude::*;

const VBA_RELATIONSHIP: &str = "http://schemas.microsoft.com/office/2006/relationships/vbaProject";

/// Document code names and native types come from workbook/worksheet properties,
/// never from a sheet's display name or the exported class attributes.
#[wasm_bindgen]
pub fn spreadsheet_vba_document_modules(package: &[u8]) -> Result<JsValue, JsError> {
    let documents = read_document_modules(package).map_err(|error| JsError::new(&error))?;
    serde_wasm_bindgen::to_value(&documents).map_err(|error| JsError::new(&error.to_string()))
}

pub(crate) fn read_document_modules(
    package: &[u8],
) -> Result<Vec<crate::spreadsheet_vba::VbaDocumentModule>, String> {
    use crate::spreadsheet_vba::VbaDocumentModule;
    let mut archive = OoxmlArchive::new(package).map_err(|error| error.to_string())?;
    let workbook = archive
        .read_part("xl/workbook.xml")
        .map_err(|error| error.to_string())?;
    let entries = document_properties(&workbook)?;
    let mut documents = Vec::new();
    // The source stored in vbaProject.bin retains VB_Base even though VBE's
    // exported .cls source omits it. WorkbookPr can retain the old code name
    // after its document component is renamed; prefer the project identity.
    // Only PROJECT's Document records qualify, never ordinary source classes.
    let mut project_documents = Vec::new();
    let mut declared_names = None;
    if let Some(project) = open_workbook_project(package)? {
        let module_names = identities(&project, None)?;
        let bytes = project
            .read_stream("/PROJECT")
            .map_err(|error| error.to_string())?;
        let encoding = codepage::to_encoding(project.information.code_page)
            .ok_or_else(|| "unsupported VBA document code page".to_string())?;
        let text = encoding
            .decode_without_bom_handling_and_without_replacement(&bytes)
            .ok_or_else(|| "invalid VBA document names".to_string())?;
        let names: std::collections::BTreeSet<_> = text
            .lines()
            .take_while(|line| !line.starts_with('['))
            .filter_map(|line| {
                line.strip_prefix("Document=")?
                    .split_once("/&H")
                    .map(|(name, _)| {
                        module_names
                            .iter()
                            .find(|identity| identity.declared_name.eq_ignore_ascii_case(name))
                            .map_or_else(
                                || name.to_ascii_lowercase(),
                                |identity| identity.name.to_ascii_lowercase(),
                            )
                    })
            })
            .collect();
        for (module, identity) in project.modules.iter().zip(&module_names) {
            if !names.contains(&identity.name.to_ascii_lowercase()) {
                continue;
            }
            let source = module_source(&project, module, identity, None)?;
            if let Some(base) = native_document_base(&source) {
                project_documents.push(VbaDocumentModule {
                    name: identity.name.clone(),
                    base_type: base.into(),
                    ..VbaDocumentModule::default()
                });
            }
        }
        declared_names = Some(names);
    }
    for (element, name) in &entries {
        if element == "workbookPr" && !name.is_empty() {
            for document in project_documents
                .iter_mut()
                .filter(|document| document.base_type == "Excel.Workbook")
            {
                document.code_name = Some(name.clone());
            }
            if !project_documents
                .iter()
                .any(|document| document.base_type == "Excel.Workbook")
                && declared_names
                    .as_ref()
                    .is_none_or(|names| names.contains(&name.to_ascii_lowercase()))
            {
                documents.push(VbaDocumentModule {
                    name: name.clone(),
                    base_type: "Excel.Workbook".into(),
                    code_name: Some(name.clone()),
                    sheet_index: None,
                });
            }
        }
    }
    let relationships = archive
        .try_read_part("xl/_rels/workbook.xml.rels")
        .map_err(|error| error.to_string())?
        .map(|xml| parse_relationships(&xml))
        .transpose()
        .map_err(|error| error.to_string())?
        .unwrap_or_default();
    for (sheet_index, (_, id)) in entries
        .into_iter()
        .filter(|(element, _)| element == "sheet")
        .enumerate()
    {
        let relationship = relationships
            .get(&id)
            .ok_or_else(|| format!("missing worksheet relationship: {id}"))?;
        // A chart sheet is not a Worksheet, even when its module has the same
        // exported attributes. Other sheet kinds need their own host surfaces.
        let base = match relationship.rel_type.as_str() {
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
            | "http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet" => {
                "Excel.Worksheet"
            }
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet"
            | "http://purl.oclc.org/ooxml/officeDocument/relationships/chartsheet" => "Excel.Chart",
            _ => continue,
        };
        let target = project_part(&relationship.target)?;
        let xml = archive
            .read_part(&target)
            .map_err(|error| error.to_string())?;
        for (element, name) in document_properties(&xml)? {
            if element == "sheetPr" && !name.is_empty() {
                if let Some(document) = project_documents
                    .iter_mut()
                    .find(|document| document.name.eq_ignore_ascii_case(&name))
                {
                    document.sheet_index = Some(sheet_index);
                    document.code_name = Some(name);
                } else if declared_names
                    .as_ref()
                    .is_none_or(|names| names.contains(&name.to_ascii_lowercase()))
                {
                    documents.push(VbaDocumentModule {
                        code_name: Some(name.clone()),
                        name,
                        base_type: base.into(),
                        sheet_index: Some(sheet_index),
                    });
                }
            }
        }
    }
    documents.extend(project_documents);
    let mut names = std::collections::BTreeSet::new();
    if documents
        .iter()
        .any(|document| !names.insert(document.name.to_ascii_lowercase()))
    {
        return Err("duplicate VBA document code name".into());
    }
    Ok(documents)
}

fn native_document_base(source: &str) -> Option<&'static str> {
    // Read the attribute header independently of executable source. A saved
    // module with a lexical error must still supply its document metadata when
    // the editor is compiling a corrected version of that source.
    for line in source
        .lines()
        .map(str::trim)
        .filter(|line| !line.is_empty())
    {
        let (left, right) = line.split_once('=')?;
        let mut words = left.split_whitespace();
        if !words.next()?.eq_ignore_ascii_case("Attribute") {
            return None;
        }
        let name = words.next()?;
        if words.next().is_some() || !name.eq_ignore_ascii_case("VB_Base") {
            continue;
        }
        // Coclass identities measured from Excel's native reference library.
        return match right.trim().to_ascii_lowercase().as_str() {
            "\"0{00020819-0000-0000-c000-000000000046}\"" => Some("Excel.Workbook"),
            "\"0{00020820-0000-0000-c000-000000000046}\"" => Some("Excel.Worksheet"),
            "\"0{00020821-0000-0000-c000-000000000046}\"" => Some("Excel.Chart"),
            _ => None,
        };
    }
    None
}

fn document_properties(xml: &str) -> Result<Vec<(String, String)>, String> {
    use oxidocs_common::xml_utils::{get_attr, local_name};
    use quick_xml::{events::Event, name::ResolveResult, reader::NsReader};
    let mut reader = NsReader::from_str(xml);
    let mut parents = Vec::new();
    let mut properties = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| error.to_string())?;
        let spreadsheet = matches!(namespace,ResolveResult::Bound(ns) if matches!(ns.as_ref(),
            b"http://schemas.openxmlformats.org/spreadsheetml/2006/main" | b"http://purl.oclc.org/ooxml/spreadsheetml/main"));
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                let name = local_name(element.name().as_ref());
                if spreadsheet {
                    let attribute = match (parents.as_slice(), name.as_str()) {
                        ([Some(root)], "workbookPr") if root == "workbook" => Some("codeName"),
                        ([Some(root)], "sheetPr")
                            if root == "worksheet" || root == "chartsheet" =>
                        {
                            Some("codeName")
                        }
                        ([Some(root), Some(sheets)], "sheet")
                            if root == "workbook" && sheets == "sheets" =>
                        {
                            Some("id")
                        }
                        _ => None,
                    };
                    if let Some(value) =
                        attribute.and_then(|attribute| get_attr(element, attribute))
                    {
                        properties.push((name.clone(), value));
                    }
                }
                if matches!(event, Event::Start(_)) {
                    parents.push(spreadsheet.then_some(name));
                }
            }
            Event::End(_) => {
                parents.pop();
            }
            Event::Eof => {
                if !parents.is_empty() {
                    return Err("unclosed XML element in document properties".into());
                }
                break;
            }
            _ => {}
        }
    }
    Ok(properties)
}

/// Read a workbook's VBA project name. A workbook without VBA returns undefined.
/// `name_code_page` can supply the target host's ANSI code page when its stored
/// project-name encoding differs from PROJECTCODEPAGE.
#[wasm_bindgen]
pub fn spreadsheet_vba_project_name(
    package: &[u8],
    name_code_page: Option<u16>,
) -> Result<Option<String>, JsError> {
    read_project_name(package, name_code_page).map_err(|error| JsError::new(&error))
}

/// Read the workbook's source modules without running project entry points.
/// `source_code_page` supplies the target host's ANSI code page when Excel's
/// source bytes use a different encoding from the stored PROJECTCODEPAGE.
#[wasm_bindgen]
pub fn spreadsheet_vba_source(
    package: &[u8],
    source_code_page: Option<u16>,
) -> Result<Option<String>, JsError> {
    read_project_source_with_code_page(package, source_code_page)
        .map_err(|error| JsError::new(&error))
}

#[cfg(test)]
fn read_project_source(package: &[u8]) -> Result<Option<String>, String> {
    read_project_source_with_code_page(package, None)
}

pub(crate) fn read_project_source_with_code_page(
    package: &[u8],
    source_code_page: Option<u16>,
) -> Result<Option<String>, String> {
    let Some(project) = open_workbook_project(package)? else {
        return Ok(None);
    };
    let mut source = String::new();
    let module_names = identities(&project, source_code_page)?;
    for (module, identity) in project.modules.iter().zip(&module_names) {
        let text = module_source(&project, module, identity, source_code_page)?;
        if matches!(module.module_type, ovba::ModuleType::DocClsDesigner)
            && !text.trim_start().starts_with("VERSION ")
        {
            source.push_str("VERSION 1.0 CLASS\n");
        }
        source.push_str(&text);
        source.push('\n');
    }
    Ok(Some(source))
}

pub(crate) fn read_project_name(
    package: &[u8],
    name_code_page: Option<u16>,
) -> Result<Option<String>, String> {
    let Some(project) = open_workbook_project(package)? else {
        return Ok(None);
    };
    let bytes = project
        .read_stream("/PROJECT")
        .map_err(|error| format!("cannot read VBA PROJECT stream: {error}"))?;
    project_name_from_stream(&project, &bytes, name_code_page).map(Some)
}

fn open_workbook_project(package: &[u8]) -> Result<Option<ovba::Project>, String> {
    let Some((_, bytes)) = workbook_project_part(package)? else { return Ok(None) };
    ovba::open_project(bytes)
        .map(Some)
        .map_err(|error| format!("cannot read VBA project: {error}"))
}

fn workbook_project_part(package: &[u8]) -> Result<Option<(String, Vec<u8>)>, String> {
    let mut archive = OoxmlArchive::new(package).map_err(|error| error.to_string())?;
    let Some(xml) = archive
        .try_read_part("xl/_rels/workbook.xml.rels")
        .map_err(|error| error.to_string())?
    else {
        return Ok(None);
    };
    let relationships = parse_relationships(&xml).map_err(|error| error.to_string())?;
    let mut projects = relationships
        .values()
        .filter(|relationship| relationship.rel_type == VBA_RELATIONSHIP);
    let Some(relationship) = projects.next() else {
        return Ok(None);
    };
    if projects.next().is_some() {
        return Err("multiple VBA project relationships in workbook".to_string());
    }
    let target = project_part(&relationship.target)?;
    let bytes = archive
        .read_binary_part(&target)
        .map_err(|error| error.to_string())?;
    Ok(Some((target, bytes)))
}

/// Save edits to the source of an existing workbook VBA project. Unchanged
/// source preserves the entire package byte-for-byte. The host code page has
/// the same meaning as in `spreadsheet_vba_source`.
#[wasm_bindgen]
pub fn save_spreadsheet_vba_source(package: &[u8], source: &str, source_code_page: Option<u16>) -> Result<Vec<u8>, JsError> {
    write_project_source(package, source, source_code_page).map_err(|error| JsError::new(&error))
}

pub(crate) fn write_project_source(package: &[u8], source: &str, source_code_page: Option<u16>) -> Result<Vec<u8>, String> {
    // The complete exported source can contain comments before a module's
    // Attribute marker. Keep an unchanged export byte-identical instead of
    // reparsing those comments as the preceding module's trailing text.
    if !source.contains('\0') && read_project_source_with_code_page(package, source_code_page)?.as_deref() == Some(source) {
        return Ok(package.to_vec());
    }
    use std::io::{Cursor, Write};
    let Some((target, original)) = workbook_project_part(package)? else {
        return Err("This workbook has no VBA project to edit.".into());
    };
    let edited = crate::vba_source_writer::update_project(&original, source, source_code_page)?;
    if edited == original { return Ok(package.to_vec()); }
    let mut parts = OoxmlArchive::new(package).map_err(|error|error.to_string())?;
    let names = parts.file_names();
    if names.iter().filter(|name| **name == target).count() != 1 {
        return Err("ambiguous VBA project part".into());
    }
    let path = std::path::Path::new(&target);
    let parent = path.parent().and_then(|p|p.to_str()).unwrap_or("");
    let name = path.file_name().and_then(|p|p.to_str()).ok_or("invalid VBA project part")?;
    let rels = format!("{parent}/_rels/{name}.rels");
    if parts.try_read_part(&rels).map_err(|e|e.to_string())?
        .is_some_and(|xml|xml.contains("vbaProjectSignature"))
    {
        return Err("Editing signed VBA projects is not supported yet.".into());
    }
    let mut archive = zip::ZipArchive::new(Cursor::new(package)).map_err(|e|e.to_string())?;
    let mut output = zip::ZipWriter::new(Cursor::new(Vec::new()));
    for index in 0..archive.len() {
        let file = archive.by_index(index).map_err(|e|e.to_string())?;
        if file.name() == target {
            let options = zip::write::SimpleFileOptions::default()
                .compression_method(file.compression())
                .last_modified_time(file.last_modified().unwrap_or_default());
            output.start_file(file.name(), options).map_err(|e|e.to_string())?;
            output.write_all(&edited).map_err(|e|e.to_string())?;
        } else {
            output.raw_copy_file(file).map_err(|e|e.to_string())?;
        }
    }
    Ok(output.finish().map_err(|e|e.to_string())?.into_inner())
}

/// Real workbook-copy serialization may rebuild compound allocation while
/// retaining all live project objects. Source-only no-op saves keep their
/// separate byte-preservation contract.
pub(crate) fn compact_project_storage(package: &[u8]) -> Result<Vec<u8>, String> {
    use std::io::{Cursor, Write};
    let Some((target, original)) = workbook_project_part(package)? else { return Ok(package.to_vec()); };
    let mut parts = OoxmlArchive::new(package).map_err(|error| error.to_string())?;
    if parts.file_names().iter().filter(|name| **name == target).count() != 1 {
        return Err("ambiguous VBA project part".into());
    }
    let path = std::path::Path::new(&target);
    let parent = path.parent().and_then(|path| path.to_str()).unwrap_or("");
    let name = path.file_name().and_then(|name| name.to_str()).ok_or("invalid VBA project part")?;
    let rels = format!("{parent}/_rels/{name}.rels");
    // Preserve signed payloads until their signature's canonicalization is
    // handled explicitly; allocation rewriting must not invalidate a signature.
    if parts.try_read_part(&rels).map_err(|error| error.to_string())?
        .is_some_and(|xml| xml.contains("vbaProjectSignature")) {
        return Ok(package.to_vec());
    }
    let compact = crate::vba_compound::compact(&original)?;
    if compact == original { return Ok(package.to_vec()); }
    let mut archive = zip::ZipArchive::new(Cursor::new(package)).map_err(|error| error.to_string())?;
    let mut output = zip::ZipWriter::new(Cursor::new(Vec::new()));
    output.set_raw_comment(archive.comment().into());
    for index in 0..archive.len() {
        let file = archive.by_index(index).map_err(|error| error.to_string())?;
        if file.name() == target {
            let mut options = zip::write::SimpleFileOptions::default()
                .compression_method(file.compression())
                .last_modified_time(file.last_modified().unwrap_or_default());
            if let Some(mode) = file.unix_mode() { options = options.unix_permissions(mode); }
            output.start_file(file.name(), options).map_err(|error| error.to_string())?;
            output.write_all(&compact).map_err(|error| error.to_string())?;
        } else { output.raw_copy_file(file).map_err(|error| error.to_string())?; }
    }
    Ok(output.finish().map_err(|error| error.to_string())?.into_inner())
}

fn project_name_from_stream(
    project: &ovba::Project,
    bytes: &[u8],
    name_code_page: Option<u16>,
) -> Result<String, String> {
    let code_page = name_code_page.unwrap_or(project.information.code_page);
    let encoding = codepage::to_encoding(code_page)
        .ok_or_else(|| format!("unsupported VBA code page: {code_page}"))?;
    let text = encoding
        .decode_without_bom_handling_and_without_replacement(bytes)
        .ok_or_else(|| "invalid text in VBA PROJECT stream".to_string())?;
    // MS-OVBA 2.3.1: PROJECT properties use the project's MBCS code page.
    // https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/d88cb9d8-a475-423d-b370-cc0caaf78628
    let mut names = text
        .lines()
        .take_while(|line| !line.starts_with('['))
        .filter_map(|line| {
            line.strip_prefix("Name=\"")
                .and_then(|name| name.strip_suffix('"'))
        });
    let name = names
        .next()
        .ok_or_else(|| "missing VBA project Name property".to_string())?;
    if name.is_empty() || name.contains('\0') || names.next().is_some() {
        return Err("invalid VBA project Name property".to_string());
    }
    Ok(name.to_string())
}

fn project_part(target: &str) -> Result<String, String> {
    if target.contains(':') || target.contains('\\') || target.starts_with("//") {
        return Err("invalid VBA project relationship target".to_string());
    }
    let mut parts = if target.starts_with('/') {
        Vec::new()
    } else {
        vec!["xl"]
    };
    for part in target.split('/') {
        match part {
            "" | "." => {}
            ".." => {
                parts
                    .pop()
                    .ok_or_else(|| "VBA project target escapes package root".to_string())?;
            }
            part => parts.push(part),
        }
    }
    if parts.is_empty() {
        return Err("empty VBA project relationship target".to_string());
    }
    Ok(parts.join("/"))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn truncated_module_source_returns_an_error_without_poisoning_the_reader() {
        use std::io::{Cursor, Read, Write};
        let original = include_bytes!("../../../tests/fixtures/vba/save_calculation.xlsm");
        let mut archive = zip::ZipArchive::new(Cursor::new(original)).unwrap();
        let mut bytes = Vec::new();
        archive.by_name("xl/vbaProject.bin").unwrap().read_to_end(&mut bytes).unwrap();
        let project = ovba::open_project(bytes.clone()).unwrap();
        let module = project.modules.iter().find(|module|matches!(module.module_type,ovba::ModuleType::Procedural)).unwrap();
        let path = format!("/VBA/{}",module.stream_name);
        let mut source = project.read_stream(&path).unwrap();
        source.truncate(module.text_offset);
        source.extend_from_slice(&[1,0xff,0xbf,0]);
        let mut compound = cfb::CompoundFile::open(Cursor::new(bytes)).unwrap();
        compound.create_stream(&path).unwrap().write_all(&source).unwrap();
        let bytes = compound.into_inner().into_inner();
        let mut output = zip::ZipWriter::new(Cursor::new(Vec::new()));
        for index in 0..archive.len() {
            let file = archive.by_index(index).unwrap();
            if file.name()=="xl/vbaProject.bin" {
                output.start_file(file.name(),zip::write::SimpleFileOptions::default()).unwrap();
                output.write_all(&bytes).unwrap();
            } else { output.raw_copy_file(file).unwrap(); }
        }
        let broken = output.finish().unwrap().into_inner();
        assert!(read_project_source(&broken).unwrap_err().contains("truncated VBA compressed chunk"));
        assert!(read_project_source(original).unwrap().unwrap().contains("Function Twice"));
    }

    #[test]
    fn copied_project_storage_is_compact_without_changing_live_data() {
        let original = include_bytes!("../../../tests/fixtures/vba/save_copy_lifecycle.xlsm");
        let source = read_project_source(original).unwrap().unwrap();
        assert_eq!(write_project_source(original, &source, None).unwrap(), original);
        let packed = compact_project_storage(original).unwrap();
        assert!(packed.len() < original.len());
        assert_eq!(read_project_source(&packed).unwrap().unwrap(), source);
        assert_eq!(read_document_modules(original).unwrap(), read_document_modules(&packed).unwrap());
        let mut before = OoxmlArchive::new(original).unwrap();
        let mut after = OoxmlArchive::new(&packed).unwrap();
        assert_eq!(before.file_names(), after.file_names());
        for name in before.file_names() {
            if name != "xl/vbaProject.bin" {
                assert_eq!(before.read_binary_part(&name).unwrap(), after.read_binary_part(&name).unwrap(), "{name}");
            }
        }
        assert_eq!(compact_project_storage(&packed).unwrap(), packed);
        assert_eq!(write_project_source(&packed, &source, None).unwrap(), packed);
        let comment = [0, 255, b'V', b'B', b'A'];
        let mut commented = zip::ZipWriter::new_append(std::io::Cursor::new(original.to_vec())).unwrap();
        commented.set_raw_comment(comment.into());
        let commented = commented.finish().unwrap().into_inner();
        let packed = compact_project_storage(&commented).unwrap();
        let archive = zip::ZipArchive::new(std::io::Cursor::new(packed)).unwrap();
        assert_eq!(archive.comment(), comment);
    }

    #[test]
    fn copy_storage_compaction_retains_signed_packages_verbatim() {
        use std::io::Write;
        let original = include_bytes!("../../../tests/fixtures/vba/save_copy_lifecycle.xlsm");
        let mut archive = zip::ZipArchive::new(std::io::Cursor::new(original)).unwrap();
        let mut output = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
        for index in 0..archive.len() { output.raw_copy_file(archive.by_index(index).unwrap()).unwrap(); }
        output.start_file("xl/_rels/vbaProject.bin.rels", zip::write::SimpleFileOptions::default()).unwrap();
        output.write_all(br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rSig" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature" Target="vbaProjectSignature.bin"/></Relationships>"#).unwrap();
        output.start_file("xl/vbaProjectSignature.bin", zip::write::SimpleFileOptions::default()).unwrap();
        output.write_all(b"opaque signature payload").unwrap();
        let signed = output.finish().unwrap().into_inner();
        assert_eq!(compact_project_storage(&signed).unwrap(), signed);
    }

    #[test]
    fn edited_source_survives_package_reload_and_preserves_other_parts() {
        let original = include_bytes!("../../../tests/fixtures/vba/save_calculation.xlsm");
        let source = read_project_source(original).unwrap().unwrap();
        assert_eq!(write_project_source(original, &source, None).unwrap(), original);
        let edited = source.replace("Twice = amount * 2", "Twice = amount * 3");
        assert_ne!(source, edited);
        let saved = write_project_source(original, &edited, None).unwrap();
        let loaded = read_project_source(&saved).unwrap().unwrap();
        assert!(loaded.contains("Twice = amount * 3"));
        assert!(!loaded.contains("Twice = amount * 2"));
        let mut before = OoxmlArchive::new(original).unwrap();
        let mut after = OoxmlArchive::new(&saved).unwrap();
        assert_eq!(before.file_names(), after.file_names());
        for name in before.file_names() {
            if name != "xl/vbaProject.bin" {
                assert_eq!(before.read_binary_part(&name).unwrap(), after.read_binary_part(&name).unwrap(), "{name}");
            }
        }
        assert_eq!(read_document_modules(original).unwrap(), read_document_modules(&saved).unwrap());
    }

    #[test]
    fn removed_standard_module_is_deleted_from_the_compound_project() {
        let original = include_bytes!("../../../tests/fixtures/vba/save_copy_lifecycle.xlsm");
        let source = read_project_source(original).unwrap().unwrap();
        let modules = oxivba_core::parser::source_modules(&source);
        let remove = modules.iter().find(|module| !module.class && module.named.is_some())
            .and_then(|module| module.named.clone()).expect("fixture has a named standard module");
        let (_, original_bytes) = workbook_project_part(original).unwrap().unwrap();
        let original_project = ovba::open_project(original_bytes).unwrap();
        let removed_stream = crate::vba_module_directory::identities(&original_project, None).unwrap()
            .into_iter().find(|identity| identity.name.eq_ignore_ascii_case(&remove))
            .expect("module identity exists").stream_name;
        let retained = modules.iter().filter(|module| module.named.as_deref() != Some(remove.as_str()))
            .map(|module| module.text.as_str()).collect::<Vec<_>>().join("\r\n");
        let saved = write_project_source(original, &retained, None).unwrap();
        let reloaded = read_project_source(&saved).unwrap().unwrap();
        assert!(!reloaded.contains(&format!("Attribute VB_Name = \"{remove}\"")));
        assert_eq!(write_project_source(&saved, &reloaded, None).unwrap(), saved);
        let (_, project_bytes) = workbook_project_part(&saved).unwrap().unwrap();
        let project = ovba::open_project(project_bytes).unwrap();
        assert!(!project.modules.iter().any(|module| module.name.eq_ignore_ascii_case(&remove)));
        assert!(project.read_stream(format!("/VBA/{removed_stream}")).is_err());
    }

    #[test]
    fn source_save_resolves_nonstandard_project_part() {
        use std::io::{Cursor, Read, Write};
        let original = include_bytes!("../../../tests/fixtures/vba/save_calculation.xlsm");
        let mut archive = zip::ZipArchive::new(Cursor::new(original)).unwrap();
        let mut output = zip::ZipWriter::new(Cursor::new(Vec::new()));
        for index in 0..archive.len() {
            let mut file = archive.by_index(index).unwrap();
            let name = file.name().to_owned();
            if name == "xl/vbaProject.bin" || name == "xl/_rels/workbook.xml.rels" || name == "[Content_Types].xml" {
                let mut bytes = Vec::new(); file.read_to_end(&mut bytes).unwrap();
                let target = if name == "xl/vbaProject.bin" { "xl/code/project.bin" } else { &name };
                if name != "xl/vbaProject.bin" {
                    bytes = String::from_utf8(bytes).unwrap().replace("vbaProject.bin", "code/project.bin").into_bytes();
                }
                output.start_file(target, zip::write::SimpleFileOptions::default()).unwrap();
                output.write_all(&bytes).unwrap();
            } else { output.raw_copy_file(file).unwrap(); }
        }
        let original = output.finish().unwrap().into_inner();
        let source = read_project_source(&original).unwrap().unwrap().replace("Twice = amount * 2", "Twice = amount * 3");
        let saved = write_project_source(&original, &source, None).unwrap();
        assert!(read_project_source(&saved).unwrap().unwrap().contains("Twice = amount * 3"));
        let mut archive = OoxmlArchive::new(&saved).unwrap();
        assert!(archive.try_read_bytes("xl/vbaProject.bin").unwrap().is_none());
        assert!(archive.try_read_bytes("xl/code/project.bin").unwrap().is_some());
    }

    #[test]
    fn excel_document_code_names_supply_the_native_base_type() {
        for (bytes, name, base) in [
            (
                include_bytes!("../../../tests/fixtures/vba/document_context_sheet.xlsm")
                    .as_slice(),
                "ReportHost",
                "Excel.Worksheet",
            ),
            (
                include_bytes!("../../../tests/fixtures/vba/document_context_book.xlsm").as_slice(),
                "ArchiveHost",
                "Excel.Workbook",
            ),
        ] {
            let documents = read_document_modules(bytes).unwrap();
            assert!(
                documents
                    .iter()
                    .any(|document| document.name == name && document.base_type == base),
                "{documents:?}; source: {:?}",
                read_project_source(bytes)
            );
            let source = read_project_source(bytes).unwrap().unwrap();
            let mut references = crate::spreadsheet_vba::builtin_compilation_references().clone();
            references.document_modules = documents
                .into_iter()
                .map(|document| (document.name, document.base_type))
                .collect();
            oxivba_core::compile_project_with_references(&source, &references).unwrap();
        }
        assert!(
            read_document_modules(include_bytes!("../../../tests/fixtures/basic_test.xlsx"))
                .unwrap()
                .is_empty()
        );
        assert!(read_document_modules(b"not a workbook").is_err());
    }

    #[test]
    fn document_properties_require_the_spreadsheet_namespace_and_structural_parent() {
        assert_eq!(native_document_base("Attribute VB_Name = \"Ledger\"\nAttribute VB_Base = \"0{00020820-0000-0000-C000-000000000046}\"\nInvalid \""),Some("Excel.Worksheet"));
        assert_eq!(native_document_base("Attribute VB_Name = \"Ledger\"\nOption Explicit\nAttribute VB_Base = \"0{00020820-0000-0000-C000-000000000046}\""),None);
        let xml = r#"<x:workbook xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:other="urn:other"><x:workbookPr codeName="ArchiveHost"/><x:sheets><x:sheet name="Report" r:id="rId7"/></x:sheets><x:extLst><x:workbookPr codeName="Ignored"/></x:extLst><other:workbookPr codeName="Ignored"/></x:workbook>"#;
        assert_eq!(
            document_properties(xml).unwrap(),
            vec![
                ("workbookPr".into(), "ArchiveHost".into()),
                ("sheet".into(), "rId7".into())
            ]
        );
        let xml = r#"<worksheet xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main"><sheetPr codeName="Ledger"/></worksheet>"#;
        assert_eq!(
            document_properties(xml).unwrap(),
            vec![("sheetPr".into(), "Ledger".into())]
        );
        assert!(
            document_properties("<worksheet><sheetPr codeName=\"Unknown\"/></worksheet>")
                .unwrap()
                .is_empty()
        );
        assert!(document_properties("<broken>").is_err());
    }

    #[test]
    fn source_import_keeps_standard_and_document_modules_distinct() {
        let bytes = include_bytes!("../../../tests/fixtures/vba/save_calculation.xlsm");
        let source = read_project_source(bytes).unwrap().unwrap();
        assert!(source.contains("Function Twice"));
        let (module, classes) = oxivba_core::parse_project(&source).unwrap();
        assert!(module.items.iter().any(|item| matches!(item, oxivba_core::ModuleItem::Procedure(procedure) if procedure.name == "Twice")));
        assert!(classes.iter().any(|(name, _)| name == "ThisWorkbook"));
        assert!(classes.iter().any(|(name, _)| name == "Sheet1"));
        assert_eq!(
            read_project_source(include_bytes!("../../../tests/fixtures/basic_test.xlsx")).unwrap(),
            None
        );
    }

    #[test]
    fn native_japanese_chart_source_respects_target_host_code_page() {
        let bytes =
            include_bytes!("../../../tests/fixtures/vba/chart_persistence_sheet_unicode.xlsm");
        let project = open_workbook_project(bytes).unwrap().unwrap();
        assert_eq!(project.information.code_page, 1252);
        // Native Excel executes and exports Japanese literals although its
        // stored code-page record still says 1252. The measured host uses 932.
        let source = read_project_source_with_code_page(bytes, Some(932))
            .unwrap()
            .unwrap();
        assert!(source.contains("売上 & <利益>"));
        assert!(!read_project_source(bytes)
            .unwrap()
            .unwrap()
            .contains("売上 & <利益>"));
        assert!(read_project_source_with_code_page(bytes, Some(0)).is_err());
    }

    #[test]
    fn excel_project_names_respect_the_target_host_code_page() {
        for (data, name) in [
            (
                include_bytes!("../../../tests/fixtures/vba/named_project_ascii.xlsm").as_slice(),
                "OxiNamedProject",
            ),
            (
                include_bytes!("../../../tests/fixtures/vba/named_project_mixed.xlsm").as_slice(),
                "Another_Project42",
            ),
            (
                include_bytes!("../../../tests/fixtures/vba/named_project_unicode.xlsm").as_slice(),
                "日本語検証",
            ),
        ] {
            assert_eq!(
                read_project_name(data, Some(932)),
                Ok(Some(name.to_string()))
            );
        }
    }

    #[test]
    fn declared_encoding_is_used_without_a_host_override() {
        assert_eq!(
            read_project_name(
                include_bytes!("../../../tests/fixtures/vba/named_project_ascii.xlsm"),
                None
            ),
            Ok(Some("OxiNamedProject".to_string()))
        );
        assert!(read_project_name(
            include_bytes!("../../../tests/fixtures/vba/named_project_unicode.xlsm"),
            Some(0)
        )
        .is_err());
    }

    #[test]
    fn a_workbook_without_vba_has_no_project_name() {
        assert_eq!(
            read_project_name(
                include_bytes!("../../../tests/fixtures/basic_test.xlsx"),
                None
            ),
            Ok(None)
        );
        assert!(read_project_name(b"not a workbook", None).is_err());
    }
}
