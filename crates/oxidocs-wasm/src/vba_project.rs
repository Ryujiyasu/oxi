// SPDX-License-Identifier: MIT OR Apache-2.0

//! VBA project metadata from an OOXML workbook, without running its macros.

use crate::vba_module_directory::{identities, module_source};
use oxidocs_common::archive::OoxmlArchive;
use oxidocs_common::relationships::parse_relationships;
use wasm_bindgen::prelude::*;

const VBA_RELATIONSHIP: &str = "http://schemas.microsoft.com/office/2006/relationships/vbaProject";

/// Read declared document classes and their attachment to workbook/sheet objects.
/// Display names never identify VBA document components.
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
    // PROJECT identifies document classes and their stored VB_Base identifies
    // their native type. Package code names independently identify the actual
    // host objects: a class with a different name has no default instance.
    // An attached object's package type overrides a stale stored VB_Base.
    // Only PROJECT's Document records qualify, never ordinary source classes.
    let mut project_documents = Vec::new();
    let mut component_names = std::collections::BTreeSet::new();
    let project = open_workbook_project(package)?;
    let has_project = project.is_some();
    if let Some(project) = project {
        let module_names = identities(&project, None)?;
        component_names.extend(module_names.iter().map(|identity| identity.name.to_ascii_lowercase()));
        let bytes = project
            .read_stream("/PROJECT")
            .map_err(|error| error.to_string())?;
        let encoding = codepage::to_encoding(project.information.code_page)
            .ok_or_else(|| "unsupported VBA document code page".to_string())?;
        let text = encoding
            .decode_without_bom_handling_and_without_replacement(&bytes)
            .ok_or_else(|| "invalid VBA document names".to_string())?;
        let names = crate::vba_module_directory::document_module_names(&project, &module_names, &text);
        for (module, identity) in project.modules.iter().zip(&module_names) {
            if !names.contains(&identity.name.to_ascii_lowercase()) {
                continue;
            }
            let source = module_source(&project, module, identity, None)?;
            // PROJECT declares the document even when VB_Base is unknown.
            // Excel normalizes an unknown unbound base to Workbook on save.
            // Matching package code names still supply the actual native type.
            project_documents.push(VbaDocumentModule {
                name: identity.name.clone(),
                base_type: native_document_base(&source).unwrap_or("Excel.Workbook").into(),
                unbound: true,
                ..VbaDocumentModule::default()
            });
        }
    }
    for (element, name) in &entries {
        if element == "workbookPr" && !name.is_empty() {
            if let Some(document) = project_documents.iter_mut().find(|document|
                document.name.eq_ignore_ascii_case(name))
            {
                document.base_type = "Excel.Workbook".into();
                document.code_name = Some(name.clone());
                document.unbound = false;
            } else if !has_project {
                documents.push(VbaDocumentModule {
                    name: name.clone(),
                    base_type: "Excel.Workbook".into(),
                    code_name: Some(name.clone()),
                    sheet_index: None,
                    host_only: has_project,
                    ..VbaDocumentModule::default()
                });
            }
        }
    }
    // Excel supplies the workbook's document component on Open, even when
    // workbookPr is absent or names a component that PROJECT does not declare.
    // A sheet's codeName alone does not cause the same repair for that sheet.
    let mut generated_workbook = None;
    if has_project && !project_documents.iter().any(|document| !document.unbound) {
        let mut serial = 0u32;
        let name = loop {
            let name = if serial == 0 { "ThisWorkbook".to_owned() } else { format!("ThisWorkbook{serial}") };
            if !component_names.contains(&name.to_ascii_lowercase()) { break name; }
            serial = serial.checked_add(1).ok_or("no available workbook document name")?;
        };
        generated_workbook = Some(VbaDocumentModule {
            code_name: Some(name.clone()), name, base_type: "Excel.Workbook".into(),
            ..VbaDocumentModule::default()
        });
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
                    document.base_type = base.into();
                    document.sheet_index = Some(sheet_index);
                    document.code_name = Some(name);
                    document.unbound = false;
                } else {
                    documents.push(VbaDocumentModule {
                        code_name: Some(name.clone()),
                        name,
                        base_type: base.into(),
                        sheet_index: Some(sheet_index),
                        host_only: has_project,
                        ..VbaDocumentModule::default()
                    });
                }
            }
        }
    }
    documents.extend(project_documents);
    documents.extend(generated_workbook);
    let mut names = std::collections::BTreeSet::new();
    if documents
        .iter()
        .any(|document| !names.insert(document.name.to_ascii_lowercase()))
    {
        return Err("duplicate VBA document code name".into());
    }
    Ok(documents)
}

/// A macro-free SaveAs removes the persisted project without unloading the
/// open one. Package properties must not replace that live project's bindings.
pub(crate) fn automatic_document_modules(
    package: &[u8],
    loaded: Option<&[crate::spreadsheet_vba::VbaDocumentModule]>,
) -> Result<Vec<crate::spreadsheet_vba::VbaDocumentModule>, String> {
    let current = read_document_modules(package)?;
    if let Some(loaded) = loaded.filter(|_| open_workbook_project(package).is_ok_and(|project| project.is_none())) {
        let mut retained = loaded.to_vec();
        for document in retained.iter_mut().filter(|document| !document.unbound) {
            let code_name = document.code_name.as_deref().unwrap_or(&document.name);
            if let Some(property) = current.iter().find(|property|
                property.code_name.as_deref().is_some_and(|name|name.eq_ignore_ascii_case(code_name))
                && property.base_type.rsplit('.').next().unwrap_or_default()
                    .eq_ignore_ascii_case(document.base_type.rsplit('.').next().unwrap_or_default()))
            {
                document.sheet_index = property.sheet_index;
            }
        }
        return Ok(retained);
    }
    Ok(current)
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

/// VBA payloads remain live when Excel saves a macro-free package. Retain only
/// the project and its related parts, not a second copy of worksheet data.
#[derive(Clone)]
pub(crate) struct RetainedVbaProject {
    target: String,
    parts: std::collections::BTreeMap<String, Vec<u8>>,
    types: std::collections::BTreeMap<String, String>,
}

fn related_part(base: &str, target: &str) -> Result<String, String> {
    if target.contains([':', '\\', '?', '#']) || target.starts_with("//") {
        return Err("invalid internal project relationship target".into());
    }
    let mut parts: Vec<_> = if target.starts_with('/') { Vec::new() }
        else { base.rsplit_once('/').map_or(Vec::new(), |(parent, _)| parent.split('/').collect()) };
    for part in target.split('/') {
        match part {
            "" | "." => {}
            ".." => { parts.pop().ok_or("project relationship escapes package root")?; }
            _ => parts.push(part),
        }
    }
    if parts.is_empty() { return Err("empty project relationship target".into()); }
    Ok(parts.join("/"))
}

fn part_relationships(part: &str) -> String {
    let (parent, leaf) = part.rsplit_once('/').unwrap_or(("", part));
    if parent.is_empty() { format!("_rels/{leaf}.rels") } else { format!("{parent}/_rels/{leaf}.rels") }
}

pub(crate) fn retain_workbook_project(package: &[u8]) -> Result<Option<RetainedVbaProject>, String> {
    use std::collections::BTreeMap;
    let Some((target, _)) = workbook_project_part(package)? else { return Ok(None); };
    let mut archive = OoxmlArchive::new(package).map_err(|error| error.to_string())?;
    let types_xml = archive.read_part("[Content_Types].xml").map_err(|error| error.to_string())?;
    let mut defaults = BTreeMap::new();
    let mut overrides = BTreeMap::new();
    let mut reader = quick_xml::Reader::from_str(&types_xml);
    loop {
        match reader.read_event().map_err(|error| error.to_string())? {
            quick_xml::events::Event::Start(element) | quick_xml::events::Event::Empty(element) => {
                let attribute = |key| oxidocs_common::xml_utils::get_attr(&element, key);
                match element.local_name().as_ref() {
                    b"Default" => { if let (Some(extension), Some(value)) = (attribute("Extension"), attribute("ContentType")) {
                        defaults.insert(extension.to_ascii_lowercase(), value);
                    }}
                    b"Override" => { if let (Some(part), Some(value)) = (attribute("PartName"), attribute("ContentType")) {
                        overrides.insert(part.trim_start_matches('/').to_string(), value);
                    }}
                    _ => {}
                }
            }
            quick_xml::events::Event::Eof => break,
            _ => {}
        }
    }
    let mut result = RetainedVbaProject { target: target.clone(), parts: BTreeMap::new(), types: BTreeMap::new() };
    let mut pending = vec![target];
    while let Some(part) = pending.pop() {
        if result.parts.contains_key(&part) { continue; }
        let bytes = archive.read_binary_part(&part).map_err(|error| error.to_string())?;
        let content_type = overrides.get(&part).or_else(|| {
            part.rsplit_once('.').and_then(|(_, extension)| defaults.get(&extension.to_ascii_lowercase()))
        }).cloned().ok_or_else(|| format!("missing project content type: {part}"))?;
        result.types.insert(part.clone(), content_type);
        result.parts.insert(part.clone(), bytes);
        let rels_part = part_relationships(&part);
        if let Some(xml) = archive.try_read_part(&rels_part).map_err(|error| error.to_string())? {
            for relationship in parse_relationships(&xml).map_err(|error| error.to_string())?.values() {
                if relationship.target_mode.as_deref() != Some("External") {
                    pending.push(related_part(&part, &relationship.target)?);
                }
            }
            result.parts.insert(rels_part, xml.into_bytes());
        }
    }
    Ok(Some(result))
}

// Preserve the metadata root's namespace prefix and unrelated XML events.
// Removing a paired child also consumes its complete subtree.
fn rewrite_package_metadata(
    xml: &str,
    mut child: impl FnMut(quick_xml::events::BytesStart<'static>) -> Result<Option<quick_xml::events::BytesStart<'static>>, String>,
    additions: &[(&str, Vec<(&str, String)>)],
) -> Result<Vec<u8>, String> {
    use quick_xml::events::{BytesStart, Event};
    let mut reader = quick_xml::Reader::from_str(xml);
    let mut writer = quick_xml::Writer::new(Vec::new());
    let mut depth = 0;
    let mut prefix = String::new();
    loop {
        let event = reader.read_event().map_err(|error| error.to_string())?;
        let event = match event {
            Event::Start(element) => {
                if depth == 0 {
                    let name = String::from_utf8(element.name().as_ref().to_vec()).map_err(|error| error.to_string())?;
                    prefix = name.rsplit_once(':').map_or(String::new(), |(prefix,_)| format!("{prefix}:"));
                }
                let original_name = element.name().as_ref().to_vec();
                let element = if depth == 1 { child(element.into_owned())? } else { Some(element.into_owned()) };
                let Some(element) = element else {
                    reader.read_to_end(quick_xml::name::QName(&original_name)).map_err(|error| error.to_string())?;
                    continue;
                };
                depth += 1;
                Event::Start(element)
            }
            Event::Empty(element) if depth == 1 => {
                let Some(element) = child(element.into_owned())? else { continue; };
                Event::Empty(element)
            }
            Event::End(element) => {
                if depth == 1 {
                    for (name, attributes) in additions {
                        let qualified = format!("{prefix}{name}");
                        let mut addition = BytesStart::new(qualified);
                        for (key, value) in attributes { addition.push_attribute((*key, value.as_str())); }
                        writer.write_event(Event::Empty(addition)).map_err(|error| error.to_string())?;
                    }
                }
                depth -= 1;
                Event::End(element.into_owned())
            }
            Event::Eof => break,
            other => other.into_owned(),
        };
        writer.write_event(event).map_err(|error| error.to_string())?;
    }
    if depth != 0 { return Err("unclosed package metadata".into()); }
    Ok(writer.into_inner())
}

/// Serialize the requested OOXML format without changing the live project's
/// lifetime. A reopened macro-free book supplies no retained project.
pub(crate) fn convert_workbook_format(
    package: &[u8], format: i64, live_project: Option<&RetainedVbaProject>,
) -> Result<Vec<u8>, String> {
    use std::{collections::{BTreeMap, BTreeSet}, io::{Cursor, Write}};
    use oxidocs_common::xml_utils::get_attr;
    let main_type = match format {
        51 => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
        52 => "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
        53 => "application/vnd.ms-excel.template.macroEnabled.main+xml",
        54 => "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml",
        _ => return Err("unsupported OOXML workbook format".into()),
    };
    let macro_free = matches!(format, 51 | 54);
    let existing = retain_workbook_project(package)?;
    let retained = existing.as_ref().or(live_project);
    let removed: BTreeSet<_> = if macro_free {
        existing.as_ref().map(|project| project.parts.keys().cloned().collect()).unwrap_or_default()
    } else { BTreeSet::new() };
    let restore = if macro_free || existing.is_some() { None } else { retained };
    let mut parts = OoxmlArchive::new(package).map_err(|error| error.to_string())?;
    let names = parts.file_names();
    let mut replacements = BTreeMap::new();
    let types = parts.read_part("[Content_Types].xml").map_err(|error| error.to_string())?;
    let mut type_additions = Vec::new();
    if let Some(project) = restore {
        for (part, content_type) in &project.types {
            type_additions.push(("Override", vec![("PartName", format!("/{part}")), ("ContentType", content_type.clone())]));
        }
        for (part, bytes) in &project.parts {
            if names.contains(part) { return Err(format!("retained VBA part collides with workbook part: {part}")); }
            replacements.insert(part.clone(), bytes.clone());
        }
    }
    let types = rewrite_package_metadata(&types, |mut element| {
        if macro_free && element.local_name().as_ref() == b"Default"
            && get_attr(&element, "ContentType").as_deref() == Some("application/vnd.ms-office.vbaProject")
            && get_attr(&element, "Extension").is_some_and(|extension| !names.iter().any(|name|
                !removed.contains(name) && name.rsplit_once('.').is_some_and(|(_, ext)| ext.eq_ignore_ascii_case(&extension)))) {
            return Ok(None);
        }
        if element.local_name().as_ref() == b"Override" {
            let part = get_attr(&element, "PartName").unwrap_or_default();
            if removed.contains(part.trim_start_matches('/')) { return Ok(None); }
            if part == "/xl/workbook.xml" {
                let attributes: Vec<_> = element.attributes().map(|attribute| {
                    let attribute = attribute.map_err(|error| error.to_string())?;
                    Ok((attribute.key.as_ref().to_vec(), attribute.value.into_owned()))
                }).collect::<Result<_, String>>()?;
                element.clear_attributes();
                for (key, value) in attributes {
                    element.push_attribute((key.as_slice(), if key == b"ContentType" { main_type.as_bytes() } else { value.as_slice() }));
                }
            }
        }
        Ok(Some(element))
    }, &type_additions)?;
    replacements.insert("[Content_Types].xml".to_string(), types);
    let rels_path = "xl/_rels/workbook.xml.rels";
    let rels = parts.read_part(rels_path).map_err(|error| error.to_string())?;
    let mut rel_additions = Vec::new();
    if let Some(project) = restore {
        let relationships = parse_relationships(&rels).map_err(|error| error.to_string())?;
        let mut id = "rIdVbaProject".to_string();
        while relationships.contains_key(&id) { id.push('_'); }
        rel_additions.push(("Relationship", vec![("Id", id), ("Type", VBA_RELATIONSHIP.to_string()),
            ("Target", format!("/{}", project.target))]));
    }
    let rels = rewrite_package_metadata(&rels, |element| {
        if macro_free && get_attr(&element, "Type").as_deref() == Some(VBA_RELATIONSHIP) { Ok(None) }
        else { Ok(Some(element)) }
    }, &rel_additions)?;
    replacements.insert(rels_path.to_string(), rels);
    let mut archive = zip::ZipArchive::new(Cursor::new(package)).map_err(|error| error.to_string())?;
    let mut writer = zip::ZipWriter::new(Cursor::new(Vec::new()));
    writer.set_raw_comment(archive.comment().into());
    let mut seen = BTreeSet::new();
    for index in 0..archive.len() {
        let file = archive.by_index(index).map_err(|error| error.to_string())?;
        let name = file.name().to_string();
        if !seen.insert(name.clone()) { return Err("ambiguous workbook package part".into()); }
        if removed.contains(&name) { continue; }
        if let Some(bytes) = replacements.remove(&name) {
            let mut options = zip::write::SimpleFileOptions::default().compression_method(file.compression())
                .last_modified_time(file.last_modified().unwrap_or_default());
            if let Some(mode) = file.unix_mode() { options = options.unix_permissions(mode); }
            writer.start_file(name, options).map_err(|error| error.to_string())?;
            writer.write_all(&bytes).map_err(|error| error.to_string())?;
        } else { writer.raw_copy_file(file).map_err(|error| error.to_string())?; }
    }
    for (name, bytes) in replacements {
        writer.start_file(name, zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated)).map_err(|error| error.to_string())?;
        writer.write_all(&bytes).map_err(|error| error.to_string())?;
    }
    writer.finish().map(|cursor| cursor.into_inner()).map_err(|error| error.to_string())
}

/// Save edits to the source of an existing workbook VBA project. Unchanged
/// source preserves the entire package byte-for-byte, including LF/CRLF-only
/// differences in the editor. The host code page has
/// the same meaning as in `spreadsheet_vba_source`.
#[wasm_bindgen]
pub fn save_spreadsheet_vba_source(package: &[u8], source: &str, source_code_page: Option<u16>) -> Result<Vec<u8>, JsError> {
    write_project_source(package, source, source_code_page).map_err(|error| JsError::new(&error))
}

pub(crate) fn write_project_source(package: &[u8], source: &str, source_code_page: Option<u16>) -> Result<Vec<u8>, String> {
    write_project_source_inner(package, source, source_code_page, None, false)
}

fn write_project_source_inner(package: &[u8], source: &str, source_code_page: Option<u16>, workbook_document: Option<&str>, insert_only: bool) -> Result<Vec<u8>, String> {
    // The complete exported source can contain comments before a module's
    // Attribute marker. Keep an unchanged export byte-identical instead of
    // reparsing those comments as the preceding module's trailing text.
    if workbook_document.is_none() && !source.contains('\0') {
        if let Some(existing) = read_project_source_with_code_page(package, source_code_page)? {
            if existing == source || existing.replace("\r\n", "\n") == source.replace("\r\n", "\n") {
                return Ok(package.to_vec());
            }
        }
    }
    use std::io::{Cursor, Write};
    let Some((target, original)) = workbook_project_part(package)? else {
        return Err("This workbook has no VBA project to edit.".into());
    };
    let edited = if insert_only {
        crate::vba_source_writer::insert_workbook_document(&original, source, source_code_page,
            workbook_document.ok_or("missing workbook document name")?)?
    } else if workbook_document.is_some() {
        crate::vba_source_writer::update_project_with_workbook_document(&original, source, source_code_page, workbook_document)?
    } else { crate::vba_source_writer::update_project(&original, source, source_code_page)? };
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

/// Supply only the missing component created by the workbook loader. Existing
/// document bodies retain the separate source-import omission contract.
pub(crate) fn bootstrap_workbook_source(package: &[u8], source: &str) -> Result<(String, Option<String>), String> {
    let Some(project) = open_workbook_project(package)? else { return Ok((source.into(), None)); };
    let documents = read_document_modules(package)?;
    let Some(workbook) = documents.iter().find(|document| !document.unbound && !document.host_only && document.sheet_index.is_none()) else {
        return Ok((source.into(), None));
    };
    let names = identities(&project, None)?;
    let existing = project.modules.iter().zip(&names).find(|(_, name)| name.name.eq_ignore_ascii_case(&workbook.name));
    let generated = existing.is_none();
    let name = workbook.name.clone();
    if oxivba_core::parser::source_modules(source).iter().any(|module|
        module.named.as_deref().is_some_and(|module| module.eq_ignore_ascii_case(&name))) {
        return Ok((source.into(), generated.then_some(name)));
    }
    let header = if let Some((module, identity)) = existing {
        // After Save, an editor may retain the source it loaded before the
        // empty component was persisted. Keeping that source must not unload
        // the component or reset live globals. Never replace an existing body.
        let text = module_source(&project, module, identity, None)?;
        let default_name = name.get(..12).is_some_and(|prefix| prefix.eq_ignore_ascii_case("ThisWorkbook"))
            && name.get(12..).is_some_and(|suffix| suffix.bytes().all(|byte| byte.is_ascii_digit()));
        let inserted_stream = identity.stream_name.strip_prefix("OxiModule")
            .is_some_and(|suffix| !suffix.is_empty() && suffix.bytes().all(|byte| byte.is_ascii_digit()));
        if !default_name || !inserted_stream || native_document_base(&text) != Some("Excel.Workbook")
            || !text.lines().all(|line| line.trim().is_empty() || line.trim_start().to_ascii_lowercase().starts_with("attribute ")) {
            return Ok((source.into(), None));
        }
        format!("VERSION 1.0 CLASS\n{text}")
    } else { workbook_document_header(&name) };
    Ok((format!("{}{header}", if source.ends_with('\n') { source.to_owned() } else { format!("{source}\n") }), generated.then_some(name)))
}

fn workbook_document_header(name: &str) -> String {
    format!("VERSION 1.0 CLASS\nAttribute VB_Name = \"{name}\"\n\
        Attribute VB_Base = \"0{{00020819-0000-0000-C000-000000000046}}\"\n\
        Attribute VB_GlobalNameSpace = False\nAttribute VB_Creatable = False\n\
        Attribute VB_PredeclaredId = True\nAttribute VB_Exposed = True\n\
        Attribute VB_TemplateDerived = False\nAttribute VB_Customizable = True\n")
}

/// Workbook Save materializes the loader's component and its native code name.
/// A source-only no-op export remains byte-identical through write_project_source.
pub(crate) fn save_workbook_project(package: &[u8], source: &str, source_code_page: Option<u16>) -> Result<Vec<u8>, String> {
    let (source, generated) = bootstrap_workbook_source(package, source)?;
    let unchanged = if generated.is_some() {
        let original = read_project_source_with_code_page(package, source_code_page)?.ok_or("missing VBA project source")?;
        let original = bootstrap_workbook_source(package, &original)?.0;
        original.replace("\r\n", "\n") == source.replace("\r\n", "\n")
    } else { false };
    let header = generated.as_deref().filter(|_| unchanged).map(workbook_document_header);
    let saved = write_project_source_inner(package, header.as_deref().unwrap_or(&source), source_code_page, generated.as_deref(), unchanged)?;
    let Some(name) = generated else { return Ok(saved); };
    set_workbook_code_name(&saved, &name)
}

fn set_workbook_code_name(package: &[u8], code_name: &str) -> Result<Vec<u8>, String> {
    use oxidocs_common::xml_utils::{get_attr, local_name};
    use quick_xml::{events::{BytesStart, Event}, name::ResolveResult, reader::NsReader, Writer};
    use std::io::{Cursor, Write};
    let mut parts = OoxmlArchive::new(package).map_err(|error| error.to_string())?;
    let xml = parts.read_part("xl/workbook.xml").map_err(|error| error.to_string())?;
    let mut reader = NsReader::from_str(&xml);
    let mut writer = Writer::new(Vec::new());
    let mut depth = 0usize;
    let mut property_name = String::new();
    let mut written = false;
    loop {
        let (ns, event) = reader.read_resolved_event().map_err(|error| error.to_string())?;
        let spreadsheet = matches!(ns, ResolveResult::Bound(ns) if matches!(ns.as_ref(),
            b"http://schemas.openxmlformats.org/spreadsheetml/2006/main" | b"http://purl.oclc.org/ooxml/spreadsheetml/main"));
        match event {
            event @ (Event::Start(_) | Event::Empty(_)) => {
                let empty = matches!(event, Event::Empty(_));
                let mut element = match event { Event::Start(element) | Event::Empty(element) => element, _ => unreachable!() };
                let name = local_name(element.name().as_ref());
                if depth == 0 && spreadsheet && name == "workbook" {
                    let qualified = String::from_utf8(element.name().as_ref().to_vec()).map_err(|error| error.to_string())?;
                    property_name = qualified.rsplit_once(':').map_or_else(|| "workbookPr".into(), |(prefix, _)| format!("{prefix}:workbookPr"));
                }
                if depth == 1 && spreadsheet && name == "workbookPr" {
                    if written { return Err("duplicate workbook properties".into()); }
                    if get_attr(&element, "codeName").as_deref() == Some(code_name) { return Ok(package.to_vec()); }
                    let attributes = element.attributes().map(|attribute| {
                        let attribute = attribute.map_err(|error| error.to_string())?;
                        Ok((attribute.key.as_ref().to_vec(), attribute.value.into_owned()))
                    }).collect::<Result<Vec<_>, String>>()?;
                    element.clear_attributes();
                    for (key, value) in attributes {
                        if key != b"codeName" { element.push_attribute((key.as_slice(), value.as_slice())); }
                    }
                    element.push_attribute(("codeName", code_name));
                    written = true;
                } else if depth == 1 && !written && !matches!(name.as_str(), "fileVersion" | "fileSharing") {
                    let mut property = BytesStart::new(property_name.as_str());
                    property.push_attribute(("codeName", code_name));
                    writer.write_event(Event::Empty(property)).map_err(|error| error.to_string())?;
                    written = true;
                }
                if !empty { depth += 1; }
                writer.write_event(if empty { Event::Empty(element) } else { Event::Start(element) }).map_err(|error| error.to_string())?;
            }
            Event::End(element) => {
                if depth == 1 && !written {
                    let mut property = BytesStart::new(property_name.as_str());
                    property.push_attribute(("codeName", code_name));
                    writer.write_event(Event::Empty(property)).map_err(|error| error.to_string())?;
                    written = true;
                }
                depth = depth.checked_sub(1).ok_or("invalid workbook XML depth")?;
                writer.write_event(Event::End(element)).map_err(|error| error.to_string())?;
            }
            Event::Eof => break,
            event => writer.write_event(event).map_err(|error| error.to_string())?,
        }
    }
    let xml = writer.into_inner();
    let mut archive = zip::ZipArchive::new(Cursor::new(package)).map_err(|error| error.to_string())?;
    let mut output = zip::ZipWriter::new(Cursor::new(Vec::new()));
    output.set_raw_comment(archive.comment().into());
    for index in 0..archive.len() {
        let file = archive.by_index(index).map_err(|error| error.to_string())?;
        if file.name() == "xl/workbook.xml" {
            let options = zip::write::SimpleFileOptions::default().compression_method(file.compression())
                .last_modified_time(file.last_modified().unwrap_or_default());
            output.start_file(file.name(), options).map_err(|error| error.to_string())?;
            output.write_all(&xml).map_err(|error| error.to_string())?;
        } else { output.raw_copy_file(file).map_err(|error| error.to_string())?; }
    }
    Ok(output.finish().map_err(|error| error.to_string())?.into_inner())
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
    fn workbook_document_name_map_survives_native_ansi_replacement() {
        use std::io::{Cursor, Write};
        for (package, document) in [
            (include_bytes!("../../../tests/fixtures/vba/unicode_modules_unicode_workbook.xlsm").as_slice(), true),
            (include_bytes!("../../../tests/fixtures/vba/unicode_modules_unicode_class.xlsm").as_slice(), false),
        ] {
            let (target, binary) = workbook_project_part(package).unwrap().unwrap();
            let project = ovba::open_project(binary.clone()).unwrap();
            let identities = identities(&project, Some(932)).unwrap();
            let identity = identities.iter().find(|name| !name.name.is_ascii()).unwrap();
            let name = identity.name.clone();
            let mut directory = crate::vba_compression::read_stream(&project, "/VBA/dir", 0).unwrap();
            let record = identity.record_start;
            assert_eq!(&directory[record..record + 2], &[0x19, 0]);
            let size = u32::from_le_bytes(directory[record + 2..record + 6].try_into().unwrap()) as usize;
            directory.splice(record + 6..record + 6 + size, b"?????".iter().copied());
            directory[record + 2..record + 6].copy_from_slice(&5u32.to_le_bytes());
            let mut compound = cfb::CompoundFile::open(Cursor::new(binary)).unwrap();
            compound.create_stream("/VBA/dir").unwrap().write_all(&crate::vba_source_writer::compress(&directory)).unwrap();
            let binary = compound.into_inner().into_inner();
            let mut archive = zip::ZipArchive::new(Cursor::new(package)).unwrap();
            let mut writer = zip::ZipWriter::new(Cursor::new(Vec::new()));
            for index in 0..archive.len() {
                let file = archive.by_index(index).unwrap();
                if file.name() == target {
                    writer.start_file(file.name(), zip::write::SimpleFileOptions::default()).unwrap();
                    writer.write_all(&binary).unwrap();
                } else { writer.raw_copy_file(file).unwrap(); }
            }
            let package = writer.finish().unwrap().into_inner();
            let documents = read_document_modules(&package).unwrap();
            assert_eq!(documents.iter().any(|doc| doc.name == name), document);
            if document {
                if let Some(path) = std::env::var_os("OXI_VBA_NATIVE_DOCUMENT_MAP_PROBE") {
                    let mut native = read_document_modules(&std::fs::read(path).unwrap()).unwrap();
                    let mut expected = documents.clone();
                    native.sort_by(|a, b| a.name.cmp(&b.name));
                    expected.sort_by(|a, b| a.name.cmp(&b.name));
                    assert_eq!(native, expected, "actual Native-saved component identities");
                }
                let saved = write_project_source(&package,
                    "Attribute VB_Name = \"Driver\"\nPublic Function Probe() As Long\nProbe = 1\nEnd Function\n", Some(932)).unwrap();
                assert!(read_document_modules(&saved).unwrap().iter().any(|doc| doc.name == name && doc.unbound),
                    "omitting a document body from a source import retains the component");
            }
        }
    }

    #[test]
    fn missing_workbook_document_save_keeps_unicode_component_sources() {
        let package = include_bytes!("../../../tests/fixtures/vba/unicode_modules_unicode_workbook.xlsm");
        let source = read_project_source_with_code_page(package, Some(932)).unwrap().unwrap();
        let saved = save_workbook_project(package, &source, Some(932)).unwrap();
        let declared = read_project_source_with_code_page(package, None).unwrap().unwrap();
        assert_eq!(save_workbook_project(package, &declared, None).unwrap(), saved,
            "a metadata-only repair must not reinterpret existing source bytes");
        let reloaded = read_project_source_with_code_page(&saved, Some(932)).unwrap().unwrap();
        assert!(reloaded.replace("\r\n", "\n").starts_with(source.replace("\r\n", "\n").trim_end_matches('\n')));
        let documents = read_document_modules(&saved).unwrap();
        assert!(documents.iter().any(|doc| doc.name == "ThisWorkbook" && !doc.host_only && !doc.unbound));
        if std::env::var_os("OXI_VBA_CAPTURE_BOOTSTRAP").is_some() {
            println!("OXI_VBA_BOOTSTRAP_PACKAGE {}", serde_json::json!({
                "input":package.as_slice(), "saved":saved, "expected":"ThisWorkbook",
                "xml_name":"ThisWorkbook", "collision":"unicode", "procedure":"Driver.Probe"
            }));
        }
    }

    #[test]
    fn native_workbook_document_bootstrap_and_save_preserve_component_kinds() {
        use std::io::{Read, Write};
        let seed = include_bytes!("../../../tests/fixtures/vba/document_context_book.xlsm");
        let original = read_project_source_with_code_page(seed, Some(932)).unwrap().unwrap()
            + "\nAttribute VB_Name = \"BootstrapProbe\"\nPublic Function Meta() As String\n\
               Meta = ActiveWorkbook.CodeName & \"|\" & CStr(ActiveWorkbook.Saved)\nEnd Function\n";
        let fixture = write_project_source(seed, &original, Some(932)).unwrap();
        let fixture = fixture.as_slice();
        for (xml_name, extra, expected) in [
            (Some("ThisWorkbook"), "", "ThisWorkbook"),
            (None, "", "ThisWorkbook"),
            (Some("FreshHost"), "", "ThisWorkbook"),
            (Some("GroupContractProbe"), "", "ThisWorkbook"),
            (Some("ThisWorkbook"), "\nVERSION 1.0 CLASS\nAttribute VB_Name = \"ThisWorkbook\"\nPublic Token As Long\n", "ThisWorkbook1"),
            (Some("ThisWorkbook"), "\nAttribute VB_Name = \"ThisWorkbook\"\nPublic Token As Long\n", "ThisWorkbook1"),
            (Some("ArchiveHost"), "", "ArchiveHost"),
        ] {
            let mut input = zip::ZipArchive::new(std::io::Cursor::new(fixture)).unwrap();
            let mut writer = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
            for index in 0..input.len() {
                let mut entry = input.by_index(index).unwrap();
                if entry.name() == "xl/workbook.xml" {
                    let mut xml = String::new();
                    entry.read_to_string(&mut xml).unwrap();
                    assert_eq!(xml.matches("codeName=\"ThisWorkbook\"").count(), 1);
                    let attribute = xml_name.map(|name| format!("codeName=\"{name}\"")).unwrap_or_default();
                    let xml = xml.replace("codeName=\"ThisWorkbook\"", &attribute);
                    writer.start_file(entry.name(), zip::write::SimpleFileOptions::default()).unwrap();
                    writer.write_all(xml.as_bytes()).unwrap();
                } else { writer.raw_copy_file(entry).unwrap(); }
            }
            let mut package = writer.finish().unwrap().into_inner();
            if !extra.is_empty() {
                package = write_project_source(&package, &(original.clone() + extra), Some(932)).unwrap();
            }
            let documents = read_document_modules(&package).unwrap();
            let workbook = documents.iter().find(|doc| !doc.unbound && doc.sheet_index.is_none()).unwrap();
            assert_eq!(workbook.name, expected, "{xml_name:?}, {extra}");
            assert!(!workbook.host_only, "Native creates a workbook document before Run");
            let source = read_project_source_with_code_page(&package, Some(932)).unwrap().unwrap();
            assert_eq!(write_project_source(&package, &source, Some(932)).unwrap(), package,
                "source-only no-op exports retain their byte-preservation contract");
            let saved = save_workbook_project(&package, &source, Some(932)).unwrap();
            let project = open_workbook_project(&saved).unwrap().unwrap();
            assert!(project.modules.iter().any(|module| module.name == expected));
            let properties = project.read_stream("/PROJECT").unwrap();
            let properties = String::from_utf8(properties).unwrap();
            assert!(properties.lines().any(|line| line == format!("Document={expected}/&H00000000")));
            if !extra.is_empty() {
                let prefix = if extra.contains("VERSION 1.0 CLASS") { "Class" } else { "Module" };
                assert!(properties.lines().any(|line| line == format!("{prefix}=ThisWorkbook")));
            }
            let reloaded = read_document_modules(&saved).unwrap();
            assert_eq!(documents, reloaded, "Saving retains live native identities");
            if std::env::var_os("OXI_VBA_CAPTURE_BOOTSTRAP").is_some() {
                println!("OXI_VBA_BOOTSTRAP_PACKAGE {}", serde_json::json!({
                    "input":package, "saved":saved, "expected":expected,
                    "xml_name":xml_name, "collision":if extra.is_empty() { "none" } else if extra.contains("VERSION 1.0 CLASS") { "class" } else { "standard" }
                }));
            }
            assert_eq!(write_project_source(&saved, &read_project_source_with_code_page(&saved, Some(932)).unwrap().unwrap(), Some(932)).unwrap(), saved);
        }
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
    fn source_line_endings_preserve_unchanged_package_bytes() {
        let packages: [&[u8]; 3] = [
            include_bytes!("../../../tests/fixtures/vba/save_calculation.xlsm"),
            include_bytes!("../../../tests/fixtures/vba/save_copy_filename_shapes.xlsm"),
            include_bytes!("../../../tests/fixtures/vba/save_copy_filename_leaves.xlsm"),
        ];
        for (index, package) in packages.iter().enumerate() {
            let source = read_project_source(package).unwrap().unwrap();
            let lines = source.replace("\r\n", "\n");
            for ending in ["\n", "\r\n"] {
                let text = lines.replace('\n', ending);
                let written = write_project_source(package, &text, None).unwrap();
                assert!(written == *package, "unchanged package {index}, ending {ending:?}");
            }
            let edited = format!("{lines}\n' Retain this source edit\n");
            let written = write_project_source(package, &edited, None).unwrap();
            assert!(written != *package);
            assert!(read_project_source(&written).unwrap().unwrap().contains("' Retain this source edit"));
        }
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

        let retained = retain_workbook_project(&saved).unwrap().unwrap();
        for format in [51, 54] {
            let macro_free = convert_workbook_format(&saved, format, Some(&retained)).unwrap();
            assert!(read_project_source(&macro_free).unwrap().is_none());
            let mut archive = OoxmlArchive::new(&macro_free).unwrap();
            assert!(archive.try_read_bytes("xl/code/project.bin").unwrap().is_none());
            let restored = convert_workbook_format(&macro_free, 52, Some(&retained)).unwrap();
            assert_eq!(read_project_source(&restored).unwrap(), read_project_source(&saved).unwrap());
            let without_live_project = convert_workbook_format(&macro_free, 52, None).unwrap();
            assert!(read_project_source(&without_live_project).unwrap().is_none(),
                "reopening macro-free disk bytes cannot recreate a discarded project");
            let mut original_parts = OoxmlArchive::new(&saved).unwrap();
            let mut restored_parts = OoxmlArchive::new(&restored).unwrap();
            for name in original_parts.file_names() {
                if !matches!(name.as_str(), "[Content_Types].xml" | "xl/_rels/workbook.xml.rels") {
                    assert_eq!(original_parts.read_binary_part(&name).unwrap(),
                        restored_parts.read_binary_part(&name).unwrap(), "{name}");
                }
            }
        }
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
