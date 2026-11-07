// SPDX-License-Identifier: MIT OR Apache-2.0

use super::*;
use oxivba_core::files::{FileError, Files};

/// Input package and filesystem facts belong to the host environment rather
/// than to a VBA project's global variables.
#[derive(Clone, Default, Deserialize)]
#[serde(default, rename_all = "camelCase", deny_unknown_fields)]
pub(super) struct FileContext {
    pub package_bytes: Option<Vec<u8>>,
    pub full_name: Option<String>,
    pub read_only: Option<bool>,
    /// Monotonic host editor revision, including edits restored to their old values.
    pub edit_revision: Option<u64>,
    /// Source decoding profile used when the editor loaded this package.
    pub source_code_page: Option<u16>,
    pub folders: Vec<String>,
    pub files: Vec<ProducedFile>,
}

#[derive(Clone, Deserialize, Serialize)]
#[serde(deny_unknown_fields)]
pub(super) struct ProducedFile {
    pub path: String,
    pub bytes: Vec<u8>,
}

/// Complete closed namespace plus mount identities. A consumed initial mount
/// stays consumed after Kill, Name or RmDir, rather than resurrecting its input.
#[derive(Clone, Deserialize, Serialize)]
#[serde(rename_all="camelCase",deny_unknown_fields)]
pub(super) struct FileSystemState {
    folders: Vec<String>,
    files: Vec<ProducedFile>,
    mounted_folders: Vec<String>,
    mounted_files: Vec<String>,
    #[serde(default)]
    attributes: Vec<FileAttributes>,
    #[serde(default)]
    volumes: Vec<VolumeInfo>,
}

#[derive(Clone, Deserialize, Serialize)]
#[serde(deny_unknown_fields)]
struct VolumeInfo {
    root: String,
    label: String,
}

#[derive(Clone, Deserialize, Serialize)]
#[serde(deny_unknown_fields)]
struct FileAttributes {
    path: String,
    value: i64,
}

#[derive(Clone, Copy)]
enum SharedSaveOperation {
    Save,
    SaveAs,
    Copy,
}

#[derive(Clone, Default)]
pub(super) struct WorkbookFiles {
    pub package: Option<Vec<u8>>,
    imported_package: bool,
    pub retained_project: Option<crate::vba_project::RetainedVbaProject>,
    pub full_name: Option<String>,
    pub read_only: bool,
    pub format: Option<f64>,
    pub read_only_recommended: bool,
    pub create_backup: bool,
    pub shared: bool,
    pub shared_input_formats: BTreeMap<(usize,u32,u32), SharedInputFormat>,
    pub shared_automatic_formats: BTreeSet<(usize,u32,u32)>,
    pub shared_revision_sequence: u64,
    pub shared_format_edits: Vec<SharedFormatEdit>,
    pub shared_cell_edits: Vec<SharedCellEdit>,
    pub shared_format_ids: BTreeMap<String,u32>,
    pub source: String,
    pub edit_revision: Option<u64>,
    pub files: Files,
    pub exports: BTreeSet<String>,
    mounted_folders: BTreeSet<String>,
    mounted_files: BTreeSet<String>,
}

#[derive(Clone)]
pub(super) struct SharedInputFormat {
    pub sequence: u64,
    pub before: String,
    pub after: String,
    pub changed: bool,
    pub before_automatic: bool,
    pub before_recorded: bool,
    pub style_dxf: Option<(String,String)>,
    pub style_pasted: bool,
}

#[derive(Clone)]
pub(super) struct SharedFormatEdit {
    sequence: u64,
    range: CellRange,
    references: Option<String>,
    start: u32,
    length: Option<u32>,
    dxf: String,
    automatic: bool,
}

pub(super) struct SharedFormatRequest {
    range: CellRange,
    property: &'static str,
    before: CellStyle,
    mixed: bool,
    span: Option<(u32,u32)>,
}

#[derive(Clone)]
pub(super) struct SharedCellEdit {
    at: CellAddress,
    before: (Option<&'static str>,String),
    before_exists: bool,
    after: (Option<&'static str>,String),
    format: SharedInputFormat,
    previous_format: Option<SharedInputFormat>,
}

struct SharedCellChange {
    sequence: u64,
    sheet: usize,
    row: u32,
    column: u32,
    attributes: String,
    body: String,
}

struct SharedXmlChild<'a> {
    name: &'a str,
    attributes: Vec<(&'a str, String)>,
}

enum CopyInput {
    Filename(String),
    Error,
}

enum CopyFailure {
    Method(String),
    Excel(String),
}

impl WorkbookFiles {

    /// Move the shared namespace together with consumed mount identities.
    /// Package bytes, source and filename continue to belong to each book.
    pub(super) fn swap_application_files(&mut self, other: &mut Self) {
        std::mem::swap(&mut self.files, &mut other.files);
        std::mem::swap(&mut self.exports, &mut other.exports);
        std::mem::swap(&mut self.mounted_folders, &mut other.mounted_folders);
        std::mem::swap(&mut self.mounted_files, &mut other.mounted_files);
    }

    /// A fresh package gives a created workbook the same differential save
    /// path as an opened workbook, including unsupported-part preservation.
    pub(super) fn blank_package(
        kind: oxicells_core::ir::SheetKind,
        count: usize,
        font: &CellStyle,
        dimensions: (f32, f32),
    ) -> Result<Vec<u8>, String> {
        use std::io::Write;
        use oxicells_core::ir::SheetKind;
        let main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        let rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        let xlm = "http://schemas.microsoft.com/office/excel/2006/main";
        let (folder, relationship, content_type, root, name) = match kind {
            SheetKind::Worksheet => ("worksheets", format!("{rel}/worksheet"),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml", "worksheet", "Sheet"),
            SheetKind::Chart => ("chartsheets", format!("{rel}/chartsheet"),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml", "chartsheet", "Chart"),
            SheetKind::MacroSheet => ("macrosheets", "http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet".into(),
                "application/vnd.ms-excel.macrosheet+xml", "xm:macrosheet", "Macro"),
            SheetKind::InternationalMacroSheet => ("macrosheets", "http://schemas.microsoft.com/office/2006/relationships/xlIntlMacrosheet".into(),
                "application/vnd.ms-excel.intlmacrosheet+xml", "xm:macrosheet", "Macro"),
        };
        let mut parts = Vec::new();
        let mut sheets = String::new();
        let mut relationships = format!(r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="{rel}/styles" Target="styles.xml"/>"#);
        let mut types = r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>"#.to_string();
        if matches!(kind, SheetKind::MacroSheet | SheetKind::InternationalMacroSheet) {
            types = types.replace("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml", "application/vnd.ms-excel.sheet.macroEnabled.main+xml");
        }
        for index in 1..=count {
            sheets.push_str(&format!(r#"<sheet name="{name}{index}" sheetId="{index}" r:id="sheet{index}"/>"#));
            relationships.push_str(&format!(r#"<Relationship Id="sheet{index}" Type="{relationship}" Target="{folder}/sheet{index}.xml"/>"#));
            types.push_str(&format!(r#"<Override PartName="/xl/{folder}/sheet{index}.xml" ContentType="{content_type}"/>"#));
            let body = if kind == SheetKind::Chart {
                format!(r#"<sheetPr codeName="Chart{index}"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><drawing r:id="drawing"/>"#)
            } else {
                format!(r#"<sheetPr codeName="{name}{index}"/><dimension ref="A1"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultColWidth="{}" defaultRowHeight="{}"/><sheetData/>"#, dimensions.0, dimensions.1)
            };
            parts.push((format!("xl/{folder}/sheet{index}.xml"),
                format!(r#"<{root} xmlns="{main}" xmlns:xm="{xlm}" xmlns:r="{rel}">{body}</{root}>"#)));
            if kind == SheetKind::Chart {
                parts.push((format!("xl/{folder}/_rels/sheet{index}.xml.rels"),
                    format!(r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="drawing" Type="{rel}/drawing" Target="../drawings/drawing{index}.xml"/></Relationships>"#)));
                parts.push((format!("xl/drawings/drawing{index}.xml"),
                    format!(r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="{rel}"><xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="9144000" cy="6858000"/><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="1" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="9144000" cy="6858000"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="chart"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:absoluteAnchor></xdr:wsDr>"#)));
                parts.push((format!("xl/drawings/_rels/drawing{index}.xml.rels"),
                    format!(r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="chart" Type="{rel}/chart" Target="../charts/chart{index}.xml"/></Relationships>"#)));
                parts.push((format!("xl/charts/chart{index}.xml"),
                    r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:axId val="10"/><c:axId val="20"/></c:barChart><c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="20"/></c:catAx><c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="10"/></c:valAx></c:plotArea><c:legend><c:legendPos val="r"/></c:legend><c:plotVisOnly val="1"/></c:chart></c:chartSpace>"#.into()));
                types.push_str(&format!(r#"<Override PartName="/xl/drawings/drawing{index}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/><Override PartName="/xl/charts/chart{index}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>"#));
            }
        }
        relationships.push_str("</Relationships>");
        types.push_str("</Types>");
        parts.push(("[Content_Types].xml".into(), types));
        parts.push(("_rels/.rels".into(), format!(r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="workbook" Type="{rel}/officeDocument" Target="xl/workbook.xml"/></Relationships>"#)));
        parts.push(("xl/workbook.xml".into(), format!(r#"<workbook xmlns="{main}" xmlns:r="{rel}"><workbookPr codeName="ThisWorkbook"/><bookViews><workbookView/></bookViews><sheets>{sheets}</sheets><calcPr calcId="0" fullCalcOnLoad="1"/></workbook>"#)));
        parts.push(("xl/_rels/workbook.xml.rels".into(), relationships));
        parts.push(("xl/styles.xml".into(), format!(r#"<styleSheet xmlns="{main}"><fonts count="1">{}</fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>"#, oxicells_core::editor::serialize_cell_font(font))));
        let mut zip = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
        for (name, xml) in parts {
            zip.start_file(name, zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Deflated)).map_err(|error| error.to_string())?;
            zip.write_all(xml.as_bytes()).map_err(|error| error.to_string())?;
        }
        zip.finish().map(|cursor| cursor.into_inner()).map_err(|error| error.to_string())
    }
    pub(super) fn snapshot(&self) -> FileSystemState {
        let state=self.files.closed_snapshot();
        FileSystemState {folders:state.folders,files:state.files.into_iter().map(|(path,bytes)|ProducedFile {path,bytes}).collect(),
            mounted_folders:self.mounted_folders.iter().cloned().collect(),mounted_files:self.mounted_files.iter().cloned().collect(),
            attributes:state.attributes.into_iter().map(|(path,value)|FileAttributes {path,value}).collect(),
            volumes:self.files.volume_facts().into_iter().map(|(root,label)|VolumeInfo {root,label}).collect()}
    }

    pub(super) fn restore(&mut self, state: &FileSystemState) -> Result<(), String> {
        let mut next=self.files.clone();
        next.restore_closed_snapshot(&oxivba_core::files::FileSystemSnapshot {
            folders:state.folders.clone(),files:state.files.iter().map(|file|(file.path.clone(),file.bytes.clone())).collect(),
            attributes:state.attributes.iter().map(|entry|(entry.path.clone(),entry.value)).collect(),
        }).map_err(|error|format!("invalid VBA fileSystemState: {}",error.description()))?;
        next.restore_volume_facts(&state.volumes.iter().map(|entry|(entry.root.clone(),entry.label.clone())).collect::<Vec<_>>())
            .map_err(|error|format!("invalid VBA fileSystemState volumes: {}",error.description()))?;
        let keys=|paths:&[String]| -> Result<BTreeSet<String>,String> {
            let mut keys=BTreeSet::new();
            for path in paths {
                let path=next.normalize_snapshot_path(path).map_err(|error|format!("invalid VBA fileSystemState: {}",error.description()))?;
                if !keys.insert(path.to_lowercase()) {return Err("invalid VBA fileSystemState: duplicate mount".into());}
            }
            Ok(keys)
        };
        let mounted_folders=keys(&state.mounted_folders)?;
        let mounted_files=keys(&state.mounted_files)?;
        self.files=next;self.mounted_folders=mounted_folders;self.mounted_files=mounted_files;
        Ok(())
    }

    fn mount_folder(&mut self, folder: &str) -> Result<(), String> {
        let key=self.files.absolute_path(folder.trim()).to_lowercase();
        if !self.mounted_folders.contains(&key) {
            self.files.mount_folder(folder).map_err(|error|error.description().to_string())?;
            self.mounted_folders.insert(key);
        }
        Ok(())
    }

    fn mount_file(&mut self, path: &str, bytes: Vec<u8>) -> Result<(), String> {
        let key=self.files.absolute_path(path.trim()).to_lowercase();
        if !self.mounted_files.contains(&key) {
            if !self.files.exists(path) {self.files.replace_bytes(path,bytes).map_err(|error|error.description().to_string())?;}
            self.mounted_files.insert(key);
        }
        Ok(())
    }

    fn package_preferences(package: &[u8]) -> Result<(bool, bool), String> {
        use quick_xml::events::Event;
        use oxidocs_common::xml_utils::get_attr;
        let mut archive = oxidocs_common::archive::OoxmlArchive::new(package).map_err(|error| error.to_string())?;
        let xml = archive.read_part("xl/workbook.xml").map_err(|error| error.to_string())?;
        let mut reader = quick_xml::Reader::from_str(&xml);
        let (mut recommended, mut backup, mut depth) = (false, false, 0);
        loop {
            let event = reader.read_event().map_err(|error| error.to_string())?;
            if let Event::Start(element) | Event::Empty(element) = &event {
                if depth == 1 {
                    match element.local_name().as_ref() {
                        b"fileSharing" => recommended = get_attr(element, "readOnlyRecommended")
                            .is_some_and(|value| matches!(value.as_str(), "1" | "true")),
                        b"workbookPr" => backup = get_attr(element, "backupFile")
                            .is_some_and(|value| matches!(value.as_str(), "1" | "true")),
                        _ => {}
                    }
                }
            }
            match event {
                Event::Start(_) => depth += 1,
                Event::End(_) => depth -= 1,
                Event::Eof => break,
                _ => {}
            }
        }
        Ok((recommended, backup))
    }

    fn package_with_preferences(package: &[u8], recommended: bool, backup: bool) -> Result<Vec<u8>, String> {
        use std::io::{Cursor, Write};
        use quick_xml::events::{BytesStart, Event};
        fn element(original: Option<&BytesStart<'_>>, name: String, key: &str, value: bool) -> Result<BytesStart<'static>, String> {
            let mut next = BytesStart::new(name);
            if let Some(original) = original {
                for attr in original.attributes() {
                    let attr = attr.map_err(|error| error.to_string())?;
                    if attr.key.as_ref() != key.as_bytes() { next.push_attribute(attr); }
                }
            }
            if value { next.push_attribute((key, "1")); }
            Ok(next.into_owned())
        }
        let mut archive = oxidocs_common::archive::OoxmlArchive::new(package).map_err(|error| error.to_string())?;
        let xml = archive.read_part("xl/workbook.xml").map_err(|error| error.to_string())?;
        let mut scan = quick_xml::Reader::from_str(&xml);
        let (mut has_sharing, mut has_properties, mut depth) = (false, false, 0);
        loop {
            let event = scan.read_event().map_err(|error| error.to_string())?;
            if let Event::Start(child) | Event::Empty(child) = &event {
                if depth == 1 {
                    has_sharing |= child.local_name().as_ref() == b"fileSharing";
                    has_properties |= child.local_name().as_ref() == b"workbookPr";
                }
            }
            match event {
                Event::Start(_) => depth += 1,
                Event::End(_) => depth -= 1,
                Event::Eof => break,
                _ => {}
            }
        }
        let mut reader = quick_xml::Reader::from_str(&xml);
        let mut writer = quick_xml::Writer::new(Vec::new());
        let (mut sharing_written, mut properties_written, mut prefix, mut depth) = (false, false, String::new(), 0);
        loop {
            let event = reader.read_event().map_err(|error| error.to_string())?;
            if let Event::Start(root) = &event {
                if depth == 0 {
                    let name = String::from_utf8(root.name().as_ref().to_vec()).map_err(|error| error.to_string())?;
                    prefix = name.rsplit_once(':').map_or(String::new(), |(prefix, _)| format!("{prefix}:"));
                }
            }
            let child_name = match &event {
                Event::Start(child) | Event::Empty(child) if depth == 1 => Some(child.local_name().as_ref().to_vec()),
                _ => None,
            };
            let closing_root = matches!(&event, Event::End(_) if depth == 1);
            if closing_root || child_name.as_deref().is_some_and(|name| name != b"fileVersion" && name != b"fileSharing") {
                if !has_sharing && !sharing_written && recommended {
                    writer.write_event(Event::Empty(element(None, format!("{prefix}fileSharing"), "readOnlyRecommended", true)?))
                        .map_err(|error| error.to_string())?;
                    sharing_written = true;
                }
                if !has_properties && !properties_written && backup {
                    writer.write_event(Event::Empty(element(None, format!("{prefix}workbookPr"), "backupFile", true)?))
                        .map_err(|error| error.to_string())?;
                    properties_written = true;
                }
            }
            let starts = matches!(&event, Event::Start(_));
            let event = match event {
                Event::Start(child) | Event::Empty(child) if matches!(child_name.as_deref(), Some(b"fileSharing" | b"workbookPr")) => {
                    let sharing = child.local_name().as_ref() == b"fileSharing";
                    let key = if sharing { "readOnlyRecommended" } else { "backupFile" };
                    let enabled = if sharing { recommended } else { backup };
                    let name = String::from_utf8(child.name().as_ref().to_vec()).map_err(|error| error.to_string())?;
                    let next = element(Some(&child), name, key, enabled)?;
                    if sharing { sharing_written = true; } else { properties_written = true; }
                    if sharing && next.attributes().next().is_none() {
                        if starts { reader.read_to_end(child.name()).map_err(|error| error.to_string())?; }
                        continue;
                    }
                    if starts { depth += 1; Event::Start(next) } else { Event::Empty(next) }
                }
                Event::Start(child) => { depth += 1; Event::Start(child.into_owned()) }
                Event::End(child) => { depth -= 1; Event::End(child.into_owned()) }
                Event::Eof => break,
                other => other.into_owned(),
            };
            writer.write_event(event).map_err(|error| error.to_string())?;
        }
        let replacement = writer.into_inner();
        let mut original = zip::ZipArchive::new(Cursor::new(package)).map_err(|error| error.to_string())?;
        let mut output = zip::ZipWriter::new(Cursor::new(Vec::new()));
        output.set_raw_comment(original.comment().into());
        for index in 0..original.len() {
            let part = original.by_index(index).map_err(|error| error.to_string())?;
            if part.name() == "xl/workbook.xml" {
                let mut options = zip::write::SimpleFileOptions::default().compression_method(part.compression())
                    .last_modified_time(part.last_modified().unwrap_or_default());
                if let Some(mode) = part.unix_mode() { options = options.unix_permissions(mode); }
                output.start_file(part.name(), options).map_err(|error| error.to_string())?;
                output.write_all(&replacement).map_err(|error| error.to_string())?;
            } else { output.raw_copy_file(part).map_err(|error| error.to_string())?; }
        }
        output.finish().map(|cursor| cursor.into_inner()).map_err(|error| error.to_string())
    }


    fn xml_elements(xml: &str, tag: &[u8], wanted_depth: usize) -> Result<Vec<BTreeMap<String, String>>, String> {
        use quick_xml::events::Event;
        let mut reader = quick_xml::Reader::from_str(xml);
        let mut depth = 0;
        let mut elements = Vec::new();
        loop {
            let event = reader.read_event().map_err(|error| error.to_string())?;
            if let Event::Start(element) | Event::Empty(element) = &event {
                if depth==wanted_depth && element.local_name().as_ref()==tag {
                    let mut attributes = BTreeMap::new();
                    for attribute in element.attributes() {
                        let attribute = attribute.map_err(|error| error.to_string())?;
                        let key = String::from_utf8(attribute.key.as_ref().to_vec()).map_err(|error| error.to_string())?;
                        let value = attribute.decode_and_unescape_value(reader.decoder()).map_err(|error| error.to_string())?;
                        attributes.insert(key, value.into_owned());
                    }
                    elements.push(attributes);
                }
            }
            match event {
                Event::Start(_) => depth += 1,
                Event::End(_) => depth=depth.saturating_sub(1),
                Event::Eof => break,
                _ => {}
            }
        }
        Ok(elements)
    }

    /// Direct children as complete XML fragments; namespace declarations
    /// are retained at the embedding site when these cross package parts.
    fn xml_children(xml: &str) -> Result<Vec<(String,String)>,String> {
        use quick_xml::events::Event;
        let mut reader=quick_xml::Reader::from_str(xml);
        let (mut depth,mut start)=(0usize,None);
        let mut children=Vec::new();
        loop {
            let before=reader.buffer_position() as usize;
            match reader.read_event().map_err(|error|error.to_string())? {
                Event::Start(e) => {
                    if depth==1 {start=Some((before,String::from_utf8_lossy(e.local_name().as_ref()).into_owned()));}
                    depth+=1;
                }
                Event::Empty(e) if depth==1 => children.push((String::from_utf8_lossy(e.local_name().as_ref()).into_owned(),
                    xml[before..reader.buffer_position() as usize].into())),
                Event::End(_) => {
                    depth=depth.saturating_sub(1);
                    if depth==1 {
                        if let Some((at,name))=start.take() {children.push((name,xml[at..reader.buffer_position() as usize].into()));}
                    }
                }
                Event::Eof => break,
                _ => {}
            }
        }
        Ok(children)
    }

    fn default_revision_font(&self) -> Result<(String,Vec<(String,String)>),String> {
        use quick_xml::events::Event;
        let package=self.package.as_deref().ok_or("Missing workbook package")?;
        let mut archive=oxidocs_common::archive::OoxmlArchive::new(package).map_err(|error|error.to_string())?;
        let styles=archive.read_part("xl/styles.xml").map_err(|error|error.to_string())?;
        let parts=Self::xml_children(&styles)?;
        let section=|name: &str|parts.iter().find(|(tag,_)|tag==name).map(|(_,xml)|xml.as_str())
            .ok_or_else(||format!("Missing {name} style section"));
        let mut xf=Self::xml_elements(section("cellXfs")?,b"xf",1)?.into_iter().next().ok_or("Missing default cell format")?;
        if xf.get("applyFont").is_some_and(|v|v=="0" || v=="false") {
            if let Some(index)=xf.get("xfId").and_then(|id|id.parse::<usize>().ok()) {
                if let Some(parent)=Self::xml_elements(section("cellStyleXfs")?,b"xf",1)?.get(index) {xf=parent.clone();}
            }
        }
        let index=xf.get("fontId").and_then(|id|id.parse::<usize>().ok()).unwrap_or(0);
        let fonts=Self::xml_children(section("fonts")?)?;
        let font=fonts.get(index).ok_or("Missing default font")?;
        let mut namespaces=String::new();
        let mut reader=quick_xml::Reader::from_str(&styles);
        loop {
            match reader.read_event().map_err(|error|error.to_string())? {
                Event::Start(e) | Event::Empty(e) => {
                    for a in e.attributes() {
                        let a=a.map_err(|error|error.to_string())?;
                        let key=String::from_utf8_lossy(a.key.as_ref());
                        if key=="xmlns" || key.starts_with("xmlns:") {
                            let value=a.decode_and_unescape_value(reader.decoder()).map_err(|error|error.to_string())?;
                            namespaces.push_str(&format!(" {key}=\"{}\"",quick_xml::escape::escape(value)));
                        }
                    }
                    break;
                }
                Event::Eof => break,
                _ => {}
            }
        }
        Ok((namespaces,Self::xml_children(&font.1)?))
    }

    fn package_is_shared(package: &[u8]) -> Result<bool, String> {
        let mut archive = oxidocs_common::archive::OoxmlArchive::new(package).map_err(|error| error.to_string())?;
        let Ok(xml) = archive.read_part("xl/_rels/workbook.xml.rels") else { return Ok(false); };
        let relationships = Self::xml_elements(&xml, b"Relationship", 1)?;
        let Some(header) = relationships.iter().find(|rel|
            rel.get("Type").is_some_and(|value| value=="http://schemas.openxmlformats.org/officeDocument/2006/relationships/revisionHeaders")
            && rel.get("TargetMode").is_none_or(|value| value!="External")) else { return Ok(false); };
        let Some(target) = header.get("Target") else { return Ok(false); };
        let absolute = target.starts_with('/');
        let mut parts = if absolute {Vec::new()} else {vec!["xl"]};
        for part in target.split('/') {
            match part {
                "" | "." => {}
                ".." => { if parts.pop().is_none() { return Ok(false); } }
                part => parts.push(part),
            }
        }
        let Ok(headers) = archive.read_part(&parts.join("/")) else { return Ok(false); };
        Ok(Self::xml_elements(&headers, b"headers", 0)?.iter().any(|header| header.contains_key("guid")))
    }

    fn append_xml_children(xml: &str, children: &[SharedXmlChild<'_>]) -> Result<Vec<u8>, String> {
        use quick_xml::events::{BytesEnd, BytesStart, Event};
        fn append(writer: &mut quick_xml::Writer<Vec<u8>>, prefix: &str, children: &[SharedXmlChild<'_>]) -> Result<(), String> {
            for child in children {
                let mut element = BytesStart::new(format!("{prefix}{}", child.name));
                for (key, value) in &child.attributes { element.push_attribute((*key, value.as_str())); }
                writer.write_event(Event::Empty(element)).map_err(|error| error.to_string())?;
            }
            Ok(())
        }
        let mut reader = quick_xml::Reader::from_str(xml);
        let mut writer = quick_xml::Writer::new(Vec::new());
        let (mut depth, mut prefix, mut written) = (0usize, String::new(), false);
        loop {
            let event = reader.read_event().map_err(|error| error.to_string())?;
            if let Event::Start(root) | Event::Empty(root) = &event {
                if depth==0 {
                    let name = String::from_utf8(root.name().as_ref().to_vec()).map_err(|error| error.to_string())?;
                    prefix = name.rsplit_once(':').map_or(String::new(), |(prefix,_)| format!("{prefix}:"));
                    if matches!(&event, Event::Empty(_)) {
                        writer.write_event(Event::Start(root.clone().into_owned())).map_err(|error| error.to_string())?;
                        append(&mut writer, &prefix, children)?;
                        writer.write_event(Event::End(BytesEnd::new(name))).map_err(|error| error.to_string())?;
                        written = true;
                        continue;
                    }
                }
            }
            if matches!(&event, Event::End(_) if depth==1) {
                append(&mut writer, &prefix, children)?;written = true;
            }
            match &event {
                Event::Start(_) => depth += 1,
                Event::End(_) => depth=depth.saturating_sub(1),
                Event::Eof => break,
                _ => {}
            }
            writer.write_event(event).map_err(|error| error.to_string())?;
        }
        if !written { return Err("XML part has no root".into()); }
        Ok(writer.into_inner())
    }

    fn replace_package_parts(package: &[u8], mut replacements: BTreeMap<String, Vec<u8>>, removed: &BTreeSet<String>) -> Result<Vec<u8>, String> {
        use std::io::{Cursor, Write};
        let mut input = zip::ZipArchive::new(Cursor::new(package)).map_err(|error| error.to_string())?;
        let mut output = zip::ZipWriter::new(Cursor::new(Vec::new()));
        output.set_raw_comment(input.comment().into());
        for index in 0..input.len() {
            let part = input.by_index(index).map_err(|error| error.to_string())?;
            let name = part.name().to_string();
            if removed.contains(&name) { continue; }
            if let Some(bytes) = replacements.remove(&name) {
                let mut options = zip::write::SimpleFileOptions::default().compression_method(part.compression())
                    .last_modified_time(part.last_modified().unwrap_or_default());
                if let Some(mode) = part.unix_mode() { options=options.unix_permissions(mode); }
                output.start_file(name, options).map_err(|error| error.to_string())?;
                output.write_all(&bytes).map_err(|error| error.to_string())?;
            } else { output.raw_copy_file(part).map_err(|error| error.to_string())?; }
        }
        for (name, bytes) in replacements {
            output.start_file(name, zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated))
                .map_err(|error| error.to_string())?;
            output.write_all(&bytes).map_err(|error| error.to_string())?;
        }
        output.finish().map(|output| output.into_inner()).map_err(|error| error.to_string())
    }

    fn remove_xml_children(xml: &str, name: &[u8], attribute: &[u8], values: &BTreeSet<String>) -> Result<Vec<u8>, String> {
        use quick_xml::events::Event;
        let mut reader = quick_xml::Reader::from_str(xml);
        let mut writer = quick_xml::Writer::new(Vec::new());
        let (mut depth, mut skipping) = (0usize, None);
        loop {
            let event = reader.read_event().map_err(|error| error.to_string())?;
            let remove = match &event {
                Event::Start(element) | Event::Empty(element) if depth==1 && element.local_name().as_ref()==name =>
                    element.try_get_attribute(attribute).map_err(|error|error.to_string())?
                        .map(|value| value.decode_and_unescape_value(reader.decoder()).map(|value|values.contains(value.as_ref())))
                        .transpose().map_err(|error|error.to_string())?.unwrap_or(false),
                _ => false,
            };
            if remove && matches!(&event, Event::Start(_)) { skipping=Some(depth); }
            match &event {
                Event::Start(_) => depth += 1,
                Event::End(_) => depth=depth.saturating_sub(1),
                Event::Eof => break,
                _ => {}
            }
            if let Some(parent) = skipping {
                if depth==parent && matches!(&event,Event::End(_)) { skipping=None; }
                continue;
            }
            if !remove { writer.write_event(event).map_err(|error|error.to_string())?; }
        }
        Ok(writer.into_inner())
    }

    fn package_with_shared_formats(package: &[u8]) -> Result<Vec<u8>, String> {
        use quick_xml::{events::Event, Reader, Writer};
        let mut archive = oxidocs_common::archive::OoxmlArchive::new(package).map_err(|e| e.to_string())?;
        let styles = archive.read_part("xl/styles.xml").map_err(|e| e.to_string())?;
        let mut reader = Reader::from_str(&styles);
        let mut writer = Writer::new(Vec::new());
        let mut depth = 0usize;
        let mut formats_depth = None;
        let mut omitted_depth = None;
        loop {
            let event = reader.read_event().map_err(|e| e.to_string())?;
            let level = match &event {
                Event::Start(_) => { let level = depth; depth += 1; level },
                Event::End(_) => { depth = depth.saturating_sub(1); depth },
                _ => depth,
            };
            if let Some(omitted) = omitted_depth {
                if matches!(event, Event::End(_)) && level == omitted { omitted_depth = None; }
                if matches!(event, Event::Eof) { return Err("Unclosed default cell format element".into()); }
                continue;
            }
            match &event {
                Event::Start(e) if e.local_name().as_ref() == b"cellXfs" => formats_depth = Some(level),
                Event::End(e) if e.local_name().as_ref() == b"cellXfs" => formats_depth = None,
                Event::Start(e) | Event::Empty(e)
                    if formats_depth.is_some_and(|parent| level == parent + 1) && e.local_name().as_ref() == b"xf" =>
                {
                    let mut tag = e.to_owned();
                    tag.clear_attributes();
                    for attribute in e.attributes() {
                        let attribute = attribute.map_err(|e| e.to_string())?;
                        tag.push_attribute(attribute);
                    }
                    for name in ["numFmtId", "fontId", "fillId", "borderId"] {
                        if e.try_get_attribute(name).map_err(|e| e.to_string())?.is_none() {
                            tag.push_attribute((name, "0"));
                        }
                    }
                    writer.write_event(if matches!(event, Event::Empty(_)) { Event::Empty(tag) } else { Event::Start(tag) })
                        .map_err(|e| e.to_string())?;
                    continue;
                }
                Event::Start(e) | Event::Empty(e)
                    if formats_depth.is_some_and(|parent| level == parent + 2)
                        && matches!(e.local_name().as_ref(), b"alignment" | b"protection") =>
                {
                    let mut tag = e.to_owned();
                    tag.clear_attributes();
                    let mut retained = 0;
                    for attribute in e.attributes() {
                        let attribute = attribute.map_err(|e| e.to_string())?;
                        let value = attribute.decode_and_unescape_value(reader.decoder()).map_err(|e| e.to_string())?;
                        let omit = matches!((e.local_name().as_ref(), attribute.key.as_ref(), value.as_ref()),
                            (b"alignment", b"vertical", "bottom")
                            | (b"alignment", b"wrapText", "0" | "false")
                            | (b"alignment", b"indent", "0")
                            | (b"protection", b"locked", "1" | "true")
                            | (b"protection", b"hidden", "0" | "false"));
                        if !omit { tag.push_attribute(attribute); retained += 1; }
                    }
                    if retained == 0 {
                        if matches!(event, Event::Start(_)) { omitted_depth = Some(level); }
                        continue;
                    }
                    writer.write_event(if matches!(event, Event::Empty(_)) { Event::Empty(tag) } else { Event::Start(tag) })
                        .map_err(|e| e.to_string())?;
                    continue;
                }
                Event::Eof => break,
                _ => {}
            }
            writer.write_event(event.into_owned()).map_err(|e| e.to_string())?;
        }
        let updated = writer.into_inner();
        if updated == styles.as_bytes() { return Ok(package.to_vec()); }
        Self::replace_package_parts(package, BTreeMap::from([("xl/styles.xml".into(), updated)]), &BTreeSet::new())
    }

    fn package_without_sharing(package: &[u8]) -> Result<Vec<u8>, String> {
        if !Self::package_is_shared(package)? { return Ok(package.to_vec()); }
        let mut archive=oxidocs_common::archive::OoxmlArchive::new(package).map_err(|error|error.to_string())?;
        let types=archive.read_part("[Content_Types].xml").map_err(|error|error.to_string())?;
        let mut removed=BTreeSet::new();
        let mut part_names=BTreeSet::new();
        for item in Self::xml_elements(&types,b"Override",1)? {
            let revision_type=item.get("ContentType").is_some_and(|value| [
                "application/vnd.openxmlformats-officedocument.spreadsheetml.revisionHeaders+xml",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.revisionLog+xml",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.userNames+xml",
            ].contains(&value.as_str()));
            if revision_type {
                let name=item.get("PartName").ok_or("Missing revision part name")?;
                part_names.insert(name.clone());
                let part=name.trim_start_matches('/').to_string();
                let sidecar=part.rsplit_once('/').map_or_else(||format!("_rels/{part}.rels"),
                    |(parent,leaf)|format!("{parent}/_rels/{leaf}.rels"));
                removed.insert(part);removed.insert(sidecar);
            }
        }
        let relationship_types=BTreeSet::from([
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/revisionHeaders".into(),
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/usernames".into(),
        ]);
        let relationships=archive.read_part("xl/_rels/workbook.xml.rels").map_err(|error|error.to_string())?;
        let replacements=BTreeMap::from([
            ("[Content_Types].xml".into(),Self::remove_xml_children(&types,b"Override",b"PartName",&part_names)?),
            ("xl/_rels/workbook.xml.rels".into(),Self::remove_xml_children(&relationships,b"Relationship",b"Type",&relationship_types)?),
        ]);
        Self::replace_package_parts(package,replacements,&removed)
    }


    fn related_part(base: &str, target: &str) -> Result<String, String> {
        let mut parts: Vec<_> = if target.starts_with('/') { Vec::new() }
            else { base.rsplit_once('/').map_or("", |(parent, _)| parent).split('/').filter(|p| !p.is_empty()).collect() };
        for part in target.split('/') {
            match part {
                "" | "." => {}
                ".." => { parts.pop().ok_or("Relationship escapes the package")?; }
                part => parts.push(part),
            }
        }
        Ok(parts.join("/"))
    }

    fn sidecar_part(part: &str) -> String {
        part.rsplit_once('/').map_or_else(||format!("_rels/{part}.rels"),
            |(parent, leaf)|format!("{parent}/_rels/{leaf}.rels"))
    }

    /// Change the header counters without disturbing imported namespaces or
    /// extension nodes, and append a complete, namespace-qualified header.
    fn append_revision_header(xml: &str, attributes: &BTreeMap<&str, String>, header: &str) -> Result<Vec<u8>, String> {
        use quick_xml::events::{BytesStart, Event};
        let mut reader=quick_xml::Reader::from_str(xml);
        let mut writer=quick_xml::Writer::new(Vec::new());
        let mut depth=0usize;
        loop {
            let mut event=reader.read_event().map_err(|e|e.to_string())?;
            if depth==0 {
                if let Event::Start(root)=&event {
                    let name=String::from_utf8(root.name().as_ref().to_vec()).map_err(|e|e.to_string())?;
                    let mut replacement=BytesStart::new(name);
                    for attr in root.attributes() {
                        let attr=attr.map_err(|e|e.to_string())?;
                        if !attributes.keys().any(|key|key.as_bytes()==attr.key.as_ref()) {
                            replacement.push_attribute(attr);
                        }
                    }
                    for (key,value) in attributes { replacement.push_attribute((*key,value.as_str())); }
                    event=Event::Start(replacement);
                }
            }
            if depth==1 && matches!(&event, Event::End(_)) {
                let mut child=quick_xml::Reader::from_str(header);
                loop {
                    let next=child.read_event().map_err(|e|e.to_string())?;
                    if matches!(&next,Event::Eof) { break; }
                    writer.write_event(next).map_err(|e|e.to_string())?;
                }
            }
            match &event {
                Event::Start(_) => depth+=1,
                Event::End(_) => depth=depth.saturating_sub(1),
                Event::Eof => break,
                _ => {}
            }
            writer.write_event(event).map_err(|e|e.to_string())?;
        }
        Ok(writer.into_inner())
    }

    /// Canonical content excludes a formula's cached result and explicit style
    /// edits. The same content written twice is a single net change at save.
    fn revision_cell_content(cell: Option<&oxicells_core::ir::Cell>) -> (Option<&'static str>, String) {
        let Some(cell)=cell else { return (None,String::new()); };
        let text = |units: &[u16]| quick_xml::escape::escape(oxicells_core::encode_ooxml_string(units)).into_owned();
        if let Some(formula)=cell.formula.as_deref() {
            let units=cell.formula_units().map(|units|units.to_vec()).unwrap_or_else(||formula.encode_utf16().collect());
            let units=units.strip_prefix(&[b'=' as u16]).unwrap_or(&units);
            return (None,format!("<f>{}</f>",text(units)));
        }
        if let Some(units)=cell.value.string_units() {
            if units.is_empty() { return (None,String::new()); }
            let preserve=units.first().into_iter().chain(units.last()).any(|unit|matches!(*unit,9|10|13|32));
            return (Some("inlineStr"),format!("<is><t{}>{}</t></is>",
                if preserve {" xml:space=\"preserve\""} else {""},text(&units)));
        }
        match &cell.value {
            oxicells_core::ir::CellValue::Number(value) => (None,format!("<v>{value}</v>")),
            oxicells_core::ir::CellValue::Boolean(value) => (Some("b"),format!("<v>{}</v>",u8::from(*value))),
            oxicells_core::ir::CellValue::Error(value) => (Some("e"),format!("<v>{}</v>",quick_xml::escape::escape(value))),
            _ => (None,String::new()),
        }
    }

    fn revision_font_color(style: &CellStyle) -> String {
        let color=style.font_reference.as_ref().and_then(|r|r.color_reference()).or_else(||style.font_color.as_deref().map(|rgb| {
            let rgb=rgb.trim_start_matches('#');
            oxicells_core::ir::CellColorReference {rgb:Some(if rgb.len()==6 {format!("FF{rgb}")} else {rgb.into()}),..Default::default()}
        })).unwrap_or(oxicells_core::ir::CellColorReference {automatic:true,..Default::default()});
        oxicells_core::editor::cell_fill_color_xml("color",Some(&color),None)
    }

    fn revision_border(style: &CellStyle, edges: &[&str], outline: bool) -> String {
        let mut xml=if outline {"<border outline=\"0\">".to_string()} else {"<border>".to_string()};
        for &tag in edges {
            let edge=match tag {"left"=>&style.border_left,"right"=>&style.border_right,
                "top"=>&style.border_top,"diagonal"=>&style.border_diagonal,_=>&style.border_bottom};
            if let Some(edge)=edge {
                xml.push_str(&format!("<{tag} style=\"{}\">",quick_xml::escape::escape(&edge.style)));
                if let Some(color)=&edge.color {
                    let color=color.trim_start_matches('#');
                    let argb=if color.len()==6 {format!("FF{color}")} else {color.into()};
                    xml.push_str(&format!("<color rgb=\"{}\"/>",argb.to_ascii_uppercase()));
                } else if tag=="diagonal" {xml.push_str("<color auto=\"1\"/>");}
                else {xml.push_str("<color indexed=\"64\"/>");}
                xml.push_str(&format!("</{tag}>"));
            } else {xml.push_str(&format!("<{tag}/>"));}
        }
        xml.push_str("</border>");xml
    }

    fn normal_style_revision_dxf(style: &CellStyle, number_format_id: u32) -> String {
        let mut font=String::from("<font>");
        for (tag,value) in [("b",style.bold),("i",style.italic),("strike",style.strikethrough),
            ("condense",false),("extend",false),("outline",false),("shadow",false)] {
            font.push_str(&if value {format!("<{tag}/>")} else {format!("<{tag} val=\"0\"/>")});
        }
        font.push_str(&format!("<u val=\"{}\"/><vertAlign val=\"{}\"/><sz val=\"{}\"/>",
            style.effective_underline_kind().map_or("none",|kind|kind.as_str()),
            style.font_position.map_or("baseline",|position|position.as_str()),style.font_size.unwrap_or(11.0)));
        font.push_str(&Self::revision_font_color(style));
        if let Some(name)=style.latin_font_name() {font.push_str(&format!("<name val=\"{}\"/>",quick_xml::escape::escape(name)));}
        if let Some(family)=style.font_family {font.push_str(&format!("<family val=\"{family}\"/>"));}
        if let Some(charset)=style.font_charset.filter(|v|*v!=0) {font.push_str(&format!("<charset val=\"{charset}\"/>"));}
        let scheme=style.font_reference.as_ref().and_then(|r|r.scheme.as_deref()).unwrap_or("none");
        font.push_str(&format!("<scheme val=\"{}\"/>",quick_xml::escape::escape(scheme)));
        font.push_str("</font>");
        let edges:&[&str]=if style.border_diagonal.is_some() || style.diagonal_up || style.diagonal_down {
            &["left","right","top","bottom","diagonal"]
        } else { &["left","right","top","bottom"] };
        let border=Self::revision_border(style,edges,true)
            .replacen("<border outline=\"0\">",&format!(
                "<border diagonalUp=\"{}\" diagonalDown=\"{}\" outline=\"0\">",
                u8::from(style.diagonal_up),u8::from(style.diagonal_down)),1);
        let code=quick_xml::escape::escape(style.number_format.as_deref().unwrap_or("General"));
        let fill=oxicells_core::editor::cell_fill_xml(&interior_fill(style),true);
        let horizontal=style.horizontal_align.as_deref().unwrap_or("general");
        let vertical=style.vertical_align.as_deref().unwrap_or("bottom");
        let wrap=u8::from(style.wrap_text);let shrink=u8::from(style.shrink_to_fit);let indent=style.indent;
        let angle=style.text_rotation.unwrap_or(0);let rotation=if style.stacked_text {255} else if angle<0 {90-angle} else {angle};let justify=u8::from(style.justify_last_line);let locked=u8::from(!style.unlocked);let hidden=u8::from(style.formula_hidden);
        let reading_order=match style.reading_order {
            oxicells_core::ir::CellReadingOrder::Contextual=>0,
            oxicells_core::ir::CellReadingOrder::LeftToRight=>1,
            oxicells_core::ir::CellReadingOrder::RightToLeft=>2,
        };
        format!("{font}<numFmt numFmtId=\"{number_format_id}\" formatCode=\"{code}\"/>{fill}<alignment horizontal=\"{horizontal}\" vertical=\"{vertical}\" textRotation=\"{rotation}\" wrapText=\"{wrap}\" indent=\"{indent}\" justifyLastLine=\"{justify}\" shrinkToFit=\"{shrink}\" readingOrder=\"{reading_order}\"/>{border}<protection locked=\"{locked}\" hidden=\"{hidden}\"/>")
    }

    fn revision_number_format(format: &str) -> Option<(u32,&str)> {
        // Revision format identities use Excel's history format table; they
        // are distinct from the built-in IDs written to styles.xml.
        match format {
            "General" => Some((0,"General")),
            MONEY_FORMAT => Some((11,"\"$\"#,##0.00_);\\(\"$\"#,##0.00\\)")),
            "m/d/yyyy" => Some((19,"m/d/yyyy")),
            "0%" => Some((13,"0%")),
            "0" | "0.0" | "0.00" | "0.000" | "#,##0" | "#,##0.00" => Some((4,format)),
            "0.00%" => Some((14,"0.00%")),
            "@" => Some((30,"@")),
            _ if format.strip_prefix("0.").is_some_and(|digits| !digits.is_empty() && digits.bytes().all(|digit|digit==b'0')) => Some((4,format)),
            _ => None,
        }
    }

    fn cell_revision(at: CellAddress, before: &(Option<&'static str>,String), after: &(Option<&'static str>,String),
        format: Option<&SharedInputFormat>, sheet_id: u32, copy_only: bool, before_exists: bool) -> SharedCellChange {
        let (index,row,column)=(at.sheet,at.row,at.column);
        let new_format=format.and_then(|held|Self::revision_number_format(&held.after));
        let mut format_attributes=String::new();
        let mut format_children=String::new();
        if let Some((id,code))=new_format {
            if let Some(held)=format.filter(|held|held.changed) {
                format_attributes=format!(" odxf=\"1\" dxf=\"1\" numFmtId=\"{id}\"");
                if !held.before_automatic {
                    if let Some((old_id,old_code))=Self::revision_number_format(&held.before) {
                        format_children.push_str(&format!("<odxf><numFmt numFmtId=\"{old_id}\" formatCode=\"{}\"/></odxf>",quick_xml::escape::escape(old_code)));
                    }
                }
                format_children.push_str(&format!("<ndxf><numFmt numFmtId=\"{id}\" formatCode=\"{}\"/></ndxf>",quick_xml::escape::escape(code)));
            } else if id!=0 { format_attributes=format!(" numFmtId=\"{id}\""); }
        }
        if let Some((before,after))=format.and_then(|held|held.style_dxf.as_ref()) {
            format_attributes=format!(" odxf=\"1\"{}{}",if after.is_empty() {""} else {" dxf=\"1\""},new_format.filter(|(id,_)|*id!=0)
                .map_or_else(String::new,|(id,_)|format!(" numFmtId=\"{id}\"")));
            format_children.clear();
            if copy_only || format.is_none_or(|held|!held.before_recorded) {
                format_children.push_str(&format!("<odxf>{before}</odxf>"));
            }
            if !after.is_empty() {format_children.push_str(&format!("<ndxf>{after}</ndxf>"));}
        }
        let mut body=String::new();
        let mut col=column+1;let mut letters=Vec::new();
        while col>0 { letters.push((b'A'+((col-1)%26) as u8) as char);col=(col-1)/26; }
        let reference=format!("{}{row}",letters.into_iter().rev().collect::<String>());
        for (tag,(kind,contents)) in [("oc",before),("nc",after)] {
            if tag=="oc" && !before_exists {continue;}
            let ty=kind.map(|value|format!(" t=\"{value}\"")).unwrap_or_default();
            if contents.is_empty() { body.push_str(&format!("<{tag} r=\"{reference}\"{ty}/>")); }
            else { body.push_str(&format!("<{tag} r=\"{reference}\"{ty}>{contents}</{tag}>")); }
        }
        body.push_str(&format_children);
        SharedCellChange {
            sequence:format.map_or(u64::MAX,|held|held.sequence),sheet:index,row,column,
            attributes:format!(" sId=\"{sheet_id}\"{format_attributes}"),body,
        }
    }

    fn package_with_revision_history(package: &[u8], baseline: &[u8], current: &oxicells_core::ir::Workbook,
        author: &[u16], now: Option<f64>, operation: SharedSaveOperation,
        pending: &Self) -> Result<Vec<u8>, String> {
        let copy_only=matches!(operation,SharedSaveOperation::Copy);
        if !Self::package_is_shared(baseline)? { return Self::package_with_sharing(package,author,now); }
        let old=oxicells_core::parse_xlsx_preserving_values(baseline).map_err(|e|e.to_string())?;
        let mut archive=oxidocs_common::archive::OoxmlArchive::new(package).map_err(|e|e.to_string())?;
        let rels=archive.read_part("xl/_rels/workbook.xml.rels").map_err(|e|e.to_string())?;
        let relationships_ns="http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        let relations=Self::xml_elements(&rels,b"Relationship",1)?;
        let header_relation=relations.iter().find(|r|r.get("Type").is_some_and(|ty|ty==&format!("{relationships_ns}/revisionHeaders"))
            && r.get("TargetMode").is_none_or(|mode|mode!="External")).ok_or("Missing revision headers relationship")?;
        let header_part=Self::related_part("xl/workbook.xml",header_relation.get("Target").ok_or("Missing revision headers target")?)?;
        let headers=archive.read_part(&header_part).map_err(|e|e.to_string())?;
        let attributes=Self::xml_elements(&headers,b"headers",0)?.into_iter().next().ok_or("Missing revision headers")?;
        let parse_counter = |name: &str, default: u32| -> Result<u32,String> {
            attributes.get(name).map(|value|value.parse().map_err(|e: std::num::ParseIntError|e.to_string())).unwrap_or(Ok(default))
        };
        let mut revision_id=parse_counter("revisionId",0)?;
        let first_revision=revision_id.checked_add(1).ok_or("Revision identity overflow")?;
        let workbook=archive.read_part("xl/workbook.xml").map_err(|e|e.to_string())?;
        let sheet_ids=Self::xml_elements(&workbook,b"sheet",2)?.into_iter().map(|sheet|
            sheet.get("sheetId").ok_or_else(||"Missing sheet identity".to_string())?.parse::<u32>().map_err(|e|e.to_string()))
            .collect::<Result<Vec<_>,_>>()?;
        let mut cell_changes=Vec::new();
        let recorded_cells: BTreeSet<_>=if copy_only {
            pending.shared_cell_edits.iter().map(|edit|(edit.at.sheet,edit.at.row,edit.at.column)).collect()
        } else {BTreeSet::new()};
        for (index,sheet) in current.sheets.iter().enumerate() {
            let previous=old.sheets.get(index);
            let old_cells: BTreeMap<_,_>=previous.into_iter().flat_map(|sheet|sheet.rows.iter())
                .flat_map(|row|row.cells.iter().map(move |cell|((row.index,cell.col),cell))).collect();
            let new_cells: BTreeMap<_,_>=sheet.rows.iter().flat_map(|row|row.cells.iter().map(move |cell|((row.index,cell.col),cell))).collect();
            let positions: BTreeSet<_>=old_cells.keys().chain(new_cells.keys()).copied().collect();
            for (row,column) in positions {
                let before=Self::revision_cell_content(old_cells.get(&(row,column)).copied());
                let after=Self::revision_cell_content(new_cells.get(&(row,column)).copied());
                if recorded_cells.contains(&(index,row,column)) {continue;}
                if before==after { continue; }
                let sheet_id=sheet_ids.get(index).ok_or("Missing revision sheet identity")?;
                let format=pending.shared_input_formats.get(&(index,row,column));
                cell_changes.push(Self::cell_revision(CellAddress {sheet:index,row,column},&before,&after,format,*sheet_id,copy_only,old_cells.contains_key(&(row,column))));
            }
        }
        if copy_only {
            for edit in &pending.shared_cell_edits {
                let Some(&sheet_id)=sheet_ids.get(edit.at.sheet) else {continue;};
                cell_changes.push(Self::cell_revision(edit.at,&edit.before,&edit.after,Some(&edit.format),sheet_id,true,edit.before_exists));
            }
        }
        cell_changes.sort_by_key(|change|(change.sequence,change.sheet,change.row,change.column));
        let mut attached=BTreeSet::new();
        let mut records=Vec::new();
        for change in cell_changes {
            revision_id=revision_id.checked_add(1).ok_or("Revision identity overflow")?;
            attached.insert((change.sequence,change.sheet,change.row,change.column));
            records.push((change.sequence,format!("<rcc rId=\"{revision_id}\"{}>{}</rcc>",change.attributes,change.body)));
        }
        for edit in &pending.shared_format_edits {
            let at=edit.range.first();
            if edit.automatic && attached.contains(&(edit.sequence,at.sheet,at.row,at.column)) {continue;}
            let Some(sheet_id)=sheet_ids.get(edit.range.sheet) else {continue;};
            let reference=edit.references.clone().unwrap_or_else(||format_range_address(edit.range,false,false));
            let span=edit.length.map(|length|format!(" start=\"{}\" length=\"{length}\"",edit.start)).unwrap_or_default();
            records.push((edit.sequence,format!("<rfmt sheetId=\"{sheet_id}\" sqref=\"{reference}\"{span}><dxf>{}</dxf></rfmt>",edit.dxf)));
        }
        records.sort_by_key(|record|record.0);
        let changes=records.into_iter().map(|(_,xml)|xml).collect::<String>();
        // SaveAs opens a session; a plain Save with no queued edits keeps it.
        let no_edits = pending.shared_cell_edits.is_empty() && pending.shared_format_edits.is_empty();
        if changes.is_empty() && (copy_only || (matches!(operation,SharedSaveOperation::Save) && no_edits)) { return Ok(package.to_vec()); }
        let sidecar=Self::sidecar_part(&header_part);
        let header_rels=archive.read_part(&sidecar).map_err(|e|e.to_string())?;
        let existing=Self::xml_elements(&header_rels,b"Relationship",1)?;
        let ids: BTreeSet<_>=existing.iter().filter_map(|rel|rel.get("Id").cloned()).collect();
        let mut index=1u32;
        while ids.contains(&format!("rId{index}")) { index=index.checked_add(1).ok_or("Relationship identity overflow")?; }
        let relation_id=format!("rId{index}");
        let mut log_index=1u32;
        let parent=header_part.rsplit_once('/').map_or("",|(parent,_)|parent);
        let log_part=loop {
            let name=format!("{parent}/revisionLog{log_index}.xml");
            if archive.try_read_part(&name).map_err(|e|e.to_string())?.is_none() { break name; }
            log_index=log_index.checked_add(1).ok_or("Revision log identity overflow")?;
        };
        let guid=Self::shared_guid()?;
        let timestamp=now.map(Self::shared_timestamp).unwrap_or_else(||current_threaded_comment_timestamp().trim_end_matches('Z').to_string());
        let author=oxicells_core::encode_ooxml_string(author).replace('\t',"_x0009_").replace('\n',"_x000A_");
        let author=quick_xml::escape::escape(&author);
        let ns="http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        let max_sheet=sheet_ids.iter().max().copied().ok_or("Missing revision sheet map")?
            .checked_add(1).ok_or("Sheet identity overflow")?;
        let sheet_map=sheet_ids.iter().map(|id|format!("<sheetId val=\"{id}\"/>")).collect::<String>();
        let bounds=if revision_id<first_revision {String::new()} else {
            format!(" minRId=\"{first_revision}\"{}",if revision_id>first_revision {format!(" maxRId=\"{revision_id}\"")} else {String::new()})
        };
        let header=format!("<header xmlns=\"{ns}\" xmlns:r=\"{relationships_ns}\" guid=\"{guid}\" dateTime=\"{timestamp}\" maxSheetId=\"{max_sheet}\" userName=\"{author}\" r:id=\"{relation_id}\"{bounds}><sheetIdMap count=\"{}\">{sheet_map}</sheetIdMap></header>",sheet_ids.len());
        let mut counters=BTreeMap::from([("guid",guid),("diskRevisions","1".into())]);
        if revision_id>0 { counters.insert("revisionId",revision_id.to_string()); }
        if !copy_only { counters.insert("version",parse_counter("version",1)?.checked_add(1).ok_or("Revision version overflow")?.to_string()); }
        let log=format!("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><revisions xmlns=\"{ns}\">{changes}</revisions>");
        let types=archive.read_part("[Content_Types].xml").map_err(|e|e.to_string())?;
        let replacements=BTreeMap::from([
            (header_part,Self::append_revision_header(&headers,&counters,&header)?),
            (sidecar,Self::append_xml_children(&header_rels,&[SharedXmlChild {name:"Relationship",attributes:vec![
                ("Id",relation_id),("Type",format!("{relationships_ns}/revisionLog")),
                ("Target",format!("revisionLog{log_index}.xml"))]}])?),
            ("[Content_Types].xml".into(),Self::append_xml_children(&types,&[SharedXmlChild {name:"Override",attributes:vec![
                ("PartName",format!("/{log_part}")),("ContentType","application/vnd.openxmlformats-officedocument.spreadsheetml.revisionLog+xml".into())]}])?),
            (log_part,log.into_bytes()),
        ]);
        Self::replace_package_parts(package,replacements,&BTreeSet::new())
    }

    fn shared_guid() -> Result<String, String> {
        let mut bytes = [0u8;16];
        #[cfg(target_arch="wasm32")]
        {
            use wasm_bindgen::JsCast;
            let crypto = js_sys::Reflect::get(&js_sys::global(), &JsValue::from_str("crypto"))
                .map_err(|_| "Cannot read the host random source".to_string())?;
            let random = js_sys::Reflect::get(&crypto, &JsValue::from_str("getRandomValues"))
                .map_err(|_| "Cannot read the host random source".to_string())?
                .dyn_into::<js_sys::Function>().map_err(|_| "Host random source is unavailable".to_string())?;
            let data = js_sys::Uint8Array::new_with_length(16);
            random.call1(&crypto, data.as_ref()).map_err(|_| "Host random source failed".to_string())?;
            data.copy_to(&mut bytes);
        }
        #[cfg(not(target_arch="wasm32"))]
        {
            use std::hash::{BuildHasher, Hasher};
            let state = std::collections::hash_map::RandomState::new();
            for (index, chunk) in bytes.chunks_mut(8).enumerate() {
                let mut hash=state.build_hasher();hash.write_usize(index);
                chunk.copy_from_slice(&hash.finish().to_le_bytes());
            }
        }
        bytes[6]=(bytes[6]&0x0f)|0x40;bytes[8]=(bytes[8]&0x3f)|0x80;
        Ok(format!("{{{}}}", uuid::Uuid::from_bytes(bytes).hyphenated().to_string().to_ascii_uppercase()))
    }

    fn shared_timestamp(serial: f64) -> String {
        let seconds=((serial-25_569.0)*86_400.0+0.000001).floor() as i64;
        let days=seconds.div_euclid(86_400);let day_seconds=seconds.rem_euclid(86_400);
        let z=days+719_468;let era=z.div_euclid(146_097);let doe=z-era*146_097;
        let yoe=(doe-doe/1_460+doe/36_524-doe/146_096)/365;let mut year=yoe+era*400;
        let doy=doe-(365*yoe+yoe/4-yoe/100);let mp=(5*doy+2)/153;
        let day=doy-(153*mp+2)/5+1;let month=mp+if mp<10 {3} else {-9};
        if month<=2 { year+=1; }
        format!("{year:04}-{month:02}-{day:02}T{:02}:{:02}:{:02}",day_seconds/3600,day_seconds%3600/60,day_seconds%60)
    }

    fn package_with_sharing(package: &[u8], author: &[u16], now: Option<f64>) -> Result<Vec<u8>, String> {
        if Self::package_is_shared(package)? { return Ok(package.to_vec()); }
        let mut archive=oxidocs_common::archive::OoxmlArchive::new(package).map_err(|error| error.to_string())?;
        let workbook=archive.read_part("xl/workbook.xml").map_err(|error| error.to_string())?;
        let ids=Self::xml_elements(&workbook,b"sheet",2)?.into_iter().map(|sheet|
            sheet.get("sheetId").ok_or_else(||"Missing sheet identity".to_string())?.parse::<u32>().map_err(|error|error.to_string()))
            .collect::<Result<Vec<_>,_>>()?;
        let max_id=ids.iter().max().copied().ok_or("Cannot share a workbook without sheets")?
            .checked_add(1).ok_or("Workbook sheet identity overflow")?;
        let guid=Self::shared_guid()?;
        let timestamp=now.map(Self::shared_timestamp).unwrap_or_else(||current_threaded_comment_timestamp().trim_end_matches('Z').to_string());
        let author=oxicells_core::encode_ooxml_string(author).replace('\t',"_x0009_").replace('\n',"_x000A_");
        let author=quick_xml::escape::escape(&author);
        let ns="http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        let relationships_ns="http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        let package_ns="http://schemas.openxmlformats.org/package/2006/relationships";
        let sheet_ids=ids.iter().map(|id|format!("<sheetId val=\"{id}\"/>")).collect::<String>();
        let headers=format!("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><headers xmlns=\"{ns}\" xmlns:r=\"{relationships_ns}\" guid=\"{guid}\"><header guid=\"{guid}\" dateTime=\"{timestamp}\" maxSheetId=\"{max_id}\" userName=\"{author}\" r:id=\"rId1\"><sheetIdMap count=\"{}\">{sheet_ids}</sheetIdMap></header></headers>",ids.len());
        let users=format!("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><users xmlns=\"{ns}\" count=\"1\"><userInfo guid=\"{guid}\" name=\"{author}\" id=\"1\" dateTime=\"{timestamp}\"/></users>");
        let log=format!("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><revisions xmlns=\"{ns}\"/>");
        let header_rels=format!("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"{package_ns}\"><Relationship Id=\"rId1\" Type=\"{relationships_ns}/revisionLog\" Target=\"revisionLog1.xml\"/></Relationships>");
        let rel_xml=archive.read_part("xl/_rels/workbook.xml.rels").map_err(|error|error.to_string())?;
        let existing=Self::xml_elements(&rel_xml,b"Relationship",1)?;
        let mut ids: BTreeSet<_>=existing.iter().filter_map(|rel|rel.get("Id").cloned()).collect();
        let mut relationships=Vec::new();
        for (kind,target) in [("revisionHeaders","revisions/revisionHeaders.xml"),("usernames","revisions/userNames.xml")] {
            let ty=format!("{relationships_ns}/{kind}");
            if existing.iter().any(|rel|rel.get("Type")==Some(&ty)) {
                return Err("Existing shared relationship has no readable revision headers".into());
            }
            let mut index=1;
            while ids.contains(&format!("rId{index}")) { index+=1; }
            let id=format!("rId{index}");ids.insert(id.clone());
            relationships.push(SharedXmlChild {name:"Relationship",attributes:vec![("Id",id),("Type",ty),("Target",target.into())]});
        }
        let type_xml=archive.read_part("[Content_Types].xml").map_err(|error|error.to_string())?;
        let existing_types=Self::xml_elements(&type_xml,b"Override",1)?;
        let mut types=Vec::new();
        for (name,kind) in [("revisionHeaders.xml","revisionHeaders"),("revisionLog1.xml","revisionLog"),("userNames.xml","userNames")] {
            let part=format!("/xl/revisions/{name}");
            if !existing_types.iter().any(|item|item.get("PartName")==Some(&part)) {
                types.push(SharedXmlChild {name:"Override",attributes:vec![("PartName",part),
                    ("ContentType",format!("application/vnd.openxmlformats-officedocument.spreadsheetml.{kind}+xml"))]});
            }
        }
        let mut replacements=BTreeMap::from([
            ("xl/revisions/revisionHeaders.xml".into(),headers.into_bytes()),
            ("xl/revisions/revisionLog1.xml".into(),log.into_bytes()),
            ("xl/revisions/userNames.xml".into(),users.into_bytes()),
            ("xl/revisions/_rels/revisionHeaders.xml.rels".into(),header_rels.into_bytes()),
        ]);
        replacements.insert("xl/_rels/workbook.xml.rels".into(),Self::append_xml_children(&rel_xml,&relationships)?);
        replacements.insert("[Content_Types].xml".into(),Self::append_xml_children(&type_xml,&types)?);
        Self::replace_package_parts(package,replacements,&BTreeSet::new())
    }

    pub(super) fn package_format(package: &[u8]) -> Result<i64, String> {
        let mut archive = oxidocs_common::archive::OoxmlArchive::new(package).map_err(|error| error.to_string())?;
        let types = archive.read_part("[Content_Types].xml").map_err(|error| error.to_string())?.to_ascii_lowercase();
        Ok(if types.contains("application/vnd.ms-excel.template.macroenabled.main+xml") { 53 }
            else if types.contains("application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml") { 54 }
            else if types.contains("application/vnd.ms-excel.sheet.macroenabled.main+xml") { 52 }
            else { 51 })
    }

    pub(super) fn set_package(&mut self, package: Vec<u8>) -> Result<(), String> {
        // Reject invalid packages before replacing a loaded host's source.
        oxicells_core::XlsxEditor::new(&package).map_err(|error| error.to_string())?;
        self.format = Some(Self::package_format(&package)? as f64);
        (self.read_only_recommended, self.create_backup) = Self::package_preferences(&package)?;
        self.shared = Self::package_is_shared(&package)?;
        if let Some(project) = crate::vba_project::retain_workbook_project(&package)? { self.retained_project = Some(project); }
        self.package = Some(package);
        self.imported_package = true;
        Ok(())
    }

    pub(super) fn configure(&mut self, context: Option<FileContext>, source: &str, file_name: &mut Option<String>) -> Result<(), String> {
        self.exports.clear();
        self.source = source.to_string();
        if file_name.is_none() {
            if let Some(full_name) = self.full_name.as_ref() {
                *file_name = Some(full_name.rsplit(['\\', '/']).next().unwrap_or(full_name).to_string());
            }
        }
        let Some(context) = context else { return Ok(()); };
        if let Some(revision) = context.edit_revision { self.edit_revision = Some(revision); }
        // A retained workbook owns its current saved package. Repeated editor
        // bootstrap bytes are a mount, not a reload of pre-SaveAs metadata.
        if self.package.is_none() {
            if let Some(package) = context.package_bytes {
                self.set_package(package)?;
            }
        }
        if let Some(full_name) = context.full_name {
            if full_name.is_empty() { return Err("invalid VBA file context: empty fullName".into()); }
            *file_name = Some(full_name.rsplit(['\\', '/']).next().unwrap_or(&full_name).to_string());
            self.full_name = Some(full_name);
        }
        if let Some(read_only) = context.read_only { self.read_only = read_only; }
        if let Some(full_name) = self.full_name.clone() {
            if let Some((parent, _)) = full_name.rsplit_once(['\\', '/']) {
                if !parent.is_empty() {self.mount_folder(&format!("{parent}\\"))?;}
            }
        }
        for folder in context.folders {
            self.mount_folder(&folder)?;
        }
        // Inputs are mounts, not repeated assignments which would roll back
        // writes made by the preceding invocation of a retained project.
        for file in context.files {
            self.mount_file(&file.path,file.bytes)?;
        }
        if let (Some(path), Some(package)) = (self.full_name.clone().or(file_name.clone()), self.package.clone()) {
            let absolute=self.files.absolute_path(path.trim());
            if !self.mounted_files.contains(&absolute.to_lowercase()) {
                // The input package is a host mount. A display filename can
                // resolve into an ambient cwd absent from a replaced closed
                // namespace, so import its parent before publishing bytes.
                if let Some((parent,_))=absolute.rsplit_once('\\') {
                    self.mount_folder(&format!("{parent}\\"))?;
                }
                self.mount_file(&path,package)?;
            }
        }
        Ok(())
    }

    pub(super) fn produced_files(&self, before: &Files) -> Vec<ProducedFile> {
        let key = |path: &str| path.trim().replace('/', "\\").to_lowercase();
        let mut files: BTreeMap<_, _> = self.files.changed_files(before).into_iter()
            .map(|(path, bytes)| (key(&path), ProducedFile { path, bytes })).collect();
        for path in &self.exports {
            if let Ok(bytes) = self.files.bytes(path) {
                files.entry(key(path)).or_insert_with(|| ProducedFile { path: path.clone(), bytes: bytes.to_vec() });
            }
        }
        files.into_values().collect()
    }
}

pub(super) fn validate_context_js(options: &JsValue) -> Result<(), JsError> {
    let context = js_sys::Reflect::get(options, &JsValue::from_str("fileContext"))
        .map_err(|_| JsError::new("invalid VBA options: cannot read fileContext"))?;
    if context.is_null() || context.is_undefined() { return Ok(()); }
    if !context.is_object() || js_sys::Array::is_array(&context) {
        return Err(JsError::new("invalid VBA fileContext: expected an object"));
    }
    for key in js_sys::Object::keys(&js_sys::Object::from(context.clone())).iter() {
        let key = key.as_string().unwrap_or_default();
        if !matches!(key.as_str(), "packageBytes" | "fullName" | "readOnly" | "folders" | "files") {
            return Err(JsError::new(&format!("invalid VBA fileContext field {key}")));
        }
    }
    let files = js_sys::Reflect::get(&context, &JsValue::from_str("files"))
        .map_err(|_| JsError::new("invalid VBA fileContext files"))?;
    if files.is_null() || files.is_undefined() { return Ok(()); }
    if !js_sys::Array::is_array(&files) { return Err(JsError::new("invalid VBA fileContext files: expected an array")); }
    for file in js_sys::Array::from(&files).iter() {
        if !file.is_object() || js_sys::Array::is_array(&file) { return Err(JsError::new("invalid VBA input file")); }
        for key in js_sys::Object::keys(&js_sys::Object::from(file)).iter() {
            if !matches!(key.as_string().as_deref(), Some("path" | "bytes")) { return Err(JsError::new("invalid VBA input file field")); }
        }
    }
    Ok(())
}

pub(super) fn validate_snapshot_js(options: &JsValue) -> Result<(), JsError> {
    let state=js_sys::Reflect::get(options,&JsValue::from_str("fileSystemState"))
        .map_err(|_|JsError::new("invalid VBA fileSystemState"))?;
    if state.is_null() || state.is_undefined() {return Ok(());}
    if !state.is_object() || js_sys::Array::is_array(&state) {return Err(JsError::new("invalid VBA fileSystemState: expected an object"));}
    for key in js_sys::Object::keys(&js_sys::Object::from(state.clone())).iter() {
        if !matches!(key.as_string().as_deref(),Some("folders"|"files"|"mountedFolders"|"mountedFiles"|"attributes"|"volumes")) {return Err(JsError::new("invalid VBA fileSystemState field"));}
    }
    let volumes=js_sys::Reflect::get(&state,&JsValue::from_str("volumes")).map_err(|_|JsError::new("invalid VBA fileSystemState volumes"))?;
    if !volumes.is_undefined() {
        if !js_sys::Array::is_array(&volumes) {return Err(JsError::new("invalid VBA fileSystemState volumes: expected an array"));}
        for entry in js_sys::Array::from(&volumes).iter() {
            if entry.is_null() || !entry.is_object() || js_sys::Array::is_array(&entry) {return Err(JsError::new("invalid VBA fileSystemState volume"));}
            for key in js_sys::Object::keys(&js_sys::Object::from(entry.clone())).iter() {
                if !matches!(key.as_string().as_deref(),Some("root"|"label")) {return Err(JsError::new("invalid VBA fileSystemState volume field"));}
            }
            for key in ["root","label"] {
                if !js_sys::Reflect::get(&entry,&JsValue::from_str(key)).map_err(|_|JsError::new("invalid VBA fileSystemState volume field"))?.is_string() {
                    return Err(JsError::new("invalid VBA fileSystemState volume root/label"));
                }
            }
        }
    }
    let attributes=js_sys::Reflect::get(&state,&JsValue::from_str("attributes")).map_err(|_|JsError::new("invalid VBA fileSystemState attributes"))?;
    if !attributes.is_undefined() {
        if !js_sys::Array::is_array(&attributes) {return Err(JsError::new("invalid VBA fileSystemState attributes: expected an array"));}
        for entry in js_sys::Array::from(&attributes).iter() {
            if entry.is_null() || !entry.is_object() || js_sys::Array::is_array(&entry) {return Err(JsError::new("invalid VBA fileSystemState attribute"));}
            for key in js_sys::Object::keys(&js_sys::Object::from(entry.clone())).iter() {
                if !matches!(key.as_string().as_deref(),Some("path"|"value")) {return Err(JsError::new("invalid VBA fileSystemState attribute field"));}
            }
            let path=js_sys::Reflect::get(&entry,&JsValue::from_str("path")).map_err(|_|JsError::new("invalid VBA fileSystemState attribute path"))?;
            let value=js_sys::Reflect::get(&entry,&JsValue::from_str("value")).map_err(|_|JsError::new("invalid VBA fileSystemState attribute value"))?;
            if !path.is_string() || !value.as_f64().is_some_and(|value|value.is_finite() && value.fract()==0.0 && value.abs()<=9_007_199_254_740_991.0) {
                return Err(JsError::new("invalid VBA fileSystemState attribute path/value"));
            }
        }
    }
    for field in ["folders","files","mountedFolders","mountedFiles"] {
        let values=js_sys::Reflect::get(&state,&JsValue::from_str(field)).map_err(|_|JsError::new("invalid VBA fileSystemState field"))?;
        if !js_sys::Array::is_array(&values) {return Err(JsError::new("invalid VBA fileSystemState: expected arrays"));}
        for value in js_sys::Array::from(&values).iter() {
            if field!="files" {
                if !value.is_string() {return Err(JsError::new("invalid VBA fileSystemState path"));}
            } else {
                if !value.is_object() || js_sys::Array::is_array(&value) {return Err(JsError::new("invalid VBA fileSystemState file"));}
                for key in js_sys::Object::keys(&js_sys::Object::from(value)).iter() {
                    if !matches!(key.as_string().as_deref(),Some("path"|"bytes")) {return Err(JsError::new("invalid VBA fileSystemState file field"));}
                }
            }
        }
    }
    Ok(())
}

impl WorkbookHost<'_> {

    pub(super) fn begin_shared_cell_input(&mut self, at: CellAddress) -> Option<SharedCellEdit> {
        if !self.file_io.shared {return None;}
        let key=(at.sheet,at.row,at.column);
        let format=self.cell_here(at.sheet,at.row,at.column)
            .and_then(|cell|cell.style.number_format).or_else(||self.template_style(at).number_format)
            .unwrap_or_else(||"General".into());
        self.file_io.shared_revision_sequence=self.file_io.shared_revision_sequence.saturating_add(1);
        let before_recorded=self.file_io.shared_input_formats.contains_key(&key);
        let previous_format=self.file_io.shared_input_formats.insert(key,SharedInputFormat {
            sequence:self.file_io.shared_revision_sequence,before:format.clone(),after:format,
            changed:false,before_automatic:self.file_io.shared_automatic_formats.contains(&key),before_recorded,style_dxf:None,style_pasted:false,
        });
        let held=self.cell_here(at.sheet,at.row,at.column);
        Some(SharedCellEdit {
            at,before:WorkbookFiles::revision_cell_content(held.as_ref()),before_exists:held.is_some(),after:(None,String::new()),
            format:self.file_io.shared_input_formats[&key].clone(),previous_format,
        })
    }

    pub(super) fn finish_shared_cell_input(&mut self, edit: Option<SharedCellEdit>) {
        let Some(mut edit)=edit else {return;};
        let held=self.cell_here(edit.at.sheet,edit.at.row,edit.at.column);
        edit.after=WorkbookFiles::revision_cell_content(held.as_ref());
        // Empty-to-empty writes dirty the book but add no cell revision.
        // Restore the prior input identity so a later real save retains the
        // ordering and format of the last write Excel actually recorded.
        if edit.before.0.is_none() && edit.before.1.is_empty() && edit.after==edit.before && edit.format.style_dxf.is_none() {
            let key=(edit.at.sheet,edit.at.row,edit.at.column);
            if let Some(previous)=edit.previous_format {
                self.file_io.shared_input_formats.insert(key,previous);
            } else {
                self.file_io.shared_input_formats.remove(&key);
            }
            return;
        }
        edit.previous_format=None;
        self.file_io.shared_cell_edits.push(edit);
    }

    pub(super) fn finish_shared_cell_paste(&mut self, edit: Option<SharedCellEdit>, before: CellStyle) -> Result<(),String> {
        let Some(mut edit)=edit else {return Ok(());};
        let default=self.workbook.default_style.clone();
        let inherit=|mut style: CellStyle| {
            style.font_size=style.font_size.or(default.font_size);
            style.font_name=style.font_name.or_else(||default.font_name.clone());
            style.font_color=style.font_color.or_else(||default.font_color.clone());
            style.font_family=style.font_family.or(default.font_family);
            style.font_charset=style.font_charset.or(default.font_charset);
            style
        };
        let before=inherit(before);
        let after=inherit(self.cell_here(edit.at.sheet,edit.at.row,edit.at.column).map(|cell|cell.style)
            .unwrap_or_else(||self.template_style(edit.at)));
        let mut old=String::new();let mut new=String::new();
        let mut old_font=String::new();let mut new_font=String::new();let mut changed=BTreeSet::new();
        for (tag,was,now) in [("b",before.bold,after.bold),("i",before.italic,after.italic),
            ("u",before.underline,after.underline),("strike",before.strikethrough,after.strikethrough)] {
            if was==now {continue;}
            changed.insert(tag);
            let off=if tag=="u" {"none"} else {"0"};
            old_font.push_str(&if was {format!("<{tag}/>")} else {format!("<{tag} val=\"{off}\"/>")});
            new_font.push_str(&if now {format!("<{tag}/>")} else {format!("<{tag} val=\"{off}\"/>")});
        }
        let color=|value: &str| {
            let hex=value.trim_start_matches('#');
            if hex.len()==6 {format!("FF{hex}").to_ascii_uppercase()} else {hex.to_ascii_uppercase()}
        };
        let font_values=|style: &CellStyle| [
            ("sz",style.font_size.map(|v|v.to_string())),
            ("color",style.font_color.clone()),("name",style.font_name.clone()),
            ("family",style.font_family.map(|v|v.to_string())),("charset",style.font_charset.map(|v|v.to_string()))];
        for (((tag,was),(_,now)),(_,normal)) in font_values(&before).into_iter().zip(font_values(&after)).zip(font_values(&default)) {
            if was==now {continue;}
            changed.insert(tag);
            // A differential font inherits unchanged Normal properties.
            // In particular an old Normal size/face/color is omitted, while
            // an explicitly nondefault old value must survive in odxf.
            let old_value=if was==normal {None} else {was};
            for (out,value,is_old) in [(&mut old_font,old_value,true),(&mut new_font,now,false)] {
                if is_old && value.is_none() {continue;}
                if tag=="color" {
                    out.push_str(&value.map_or_else(||"<color auto=\"1\"/>".into(),
                        |value|format!("<color rgb=\"{}\"/>",color(&value))));
                } else if let Some(value)=value {
                    out.push_str(&format!("<{tag} val=\"{}\"/>",quick_xml::escape::escape(&value)));
                    if tag=="name" {out.push_str("<scheme val=\"none\"/>");}
                }
            }
        }
        if !changed.is_empty() {
            let default=&self.workbook.default_style;
            let normal=after.bold==default.bold && after.italic==default.italic && after.underline==default.underline
                && after.strikethrough==default.strikethrough && font_values(&after)==font_values(default);
            let namespaces=if normal || !edit.before_exists {
                let (namespaces,children)=self.file_io.default_revision_font()?;
                for (tag,xml) in children {
                    if !changed.contains(tag.as_str()) {
                        if !edit.before_exists {old_font.push_str(&xml);}
                        new_font.push_str(&xml);
                    }
                }
                namespaces
            } else {String::new()};
            let old_namespaces=if edit.before_exists {""} else {namespaces.as_str()};
            old.push_str(&format!("<font{old_namespaces}>{old_font}</font>"));
            new.push_str(&format!("<font{namespaces}>{new_font}</font>"));
        }
        if before.number_format!=after.number_format {
            old.push_str(&self.shared_format_dxf(before.number_format.as_deref().unwrap_or("General"),false));
            new.push_str(&self.shared_format_dxf(after.number_format.as_deref().unwrap_or("General"),false));
        }
        if before.bg_color!=after.bg_color {
            for (out,value) in [(&mut old,before.bg_color.as_deref()),(&mut new,after.bg_color.as_deref())] {
                out.push_str(&value.map_or_else(||"<fill><patternFill patternType=\"none\"><bgColor indexed=\"65\"/></patternFill></fill>".into(),
                    |value|format!("<fill><patternFill patternType=\"solid\"><bgColor rgb=\"{}\"/></patternFill></fill>",color(value))));
            }
        }
        let horizontal=before.horizontal_align.as_deref().unwrap_or("general")!=after.horizontal_align.as_deref().unwrap_or("general");
        let vertical=before.vertical_align.as_deref().unwrap_or("bottom")!=after.vertical_align.as_deref().unwrap_or("bottom");
        let wrap=before.wrap_text!=after.wrap_text;
        let shrink=before.shrink_to_fit!=after.shrink_to_fit;
        if horizontal || vertical || wrap || shrink {
            for (out,style,is_old) in [(&mut old,&before,true),(&mut new,&after,false)] {
                let mut attributes=String::new();
                if horizontal {attributes.push_str(&format!(" horizontal=\"{}\"",style.horizontal_align.as_deref().unwrap_or("general")));}
                // Native Excel's pasted revision DXF writes a missing new
                // vertical alignment as top, even while its live VBA getter
                // still returns bottom. The old DXF uses bottom. Keep this
                // history encoding separate from the cell's actual alignment.
                let v=style.vertical_align.as_deref().unwrap_or(if is_old {"bottom"} else {"top"});
                attributes.push_str(&format!(" vertical=\"{v}\""));
                if wrap {attributes.push_str(&format!(" wrapText=\"{}\"",u8::from(style.wrap_text)));}
                if shrink {attributes.push_str(&format!(" shrinkToFit=\"{}\"",u8::from(style.shrink_to_fit)));}
                out.push_str(&format!("<alignment{attributes}/>"));
            }
        }
        let edges: Vec<_>=[("left",&before.border_left,&after.border_left),("right",&before.border_right,&after.border_right),
            ("top",&before.border_top,&after.border_top),("bottom",&before.border_bottom,&after.border_bottom)]
            .into_iter().filter_map(|(tag,was,now)|(was!=now).then_some(tag)).collect();
        if !edges.is_empty() {
            old.push_str(&WorkbookFiles::revision_border(&before,&edges,true));
            new.push_str(&WorkbookFiles::revision_border(&after,&edges,true));
        }
        edit.format.style_pasted=true;
        edit.format.after=after.number_format.clone().unwrap_or_else(||"General".into());
        if !old.is_empty() || !new.is_empty() {
            self.file_io.shared_format_edits.push(SharedFormatEdit {
                start:0,sequence:edit.format.sequence,range:CellRange::single(edit.at),references:None,length:Some(0),dxf:new.clone(),automatic:true,
            });
            edit.format.style_dxf=Some((old,new));
        }
        self.file_io.shared_input_formats.insert((edit.at.sheet,edit.at.row,edit.at.column),edit.format.clone());
        self.finish_shared_cell_input(Some(edit));
        Ok(())
    }

    pub(super) fn refresh_shared_cell_input(&mut self, at: CellAddress) {
        if !self.file_io.shared {return;}
        let held=self.cell_here(at.sheet,at.row,at.column);
        let after=WorkbookFiles::revision_cell_content(held.as_ref());
        if let Some(edit)=self.file_io.shared_cell_edits.iter_mut().rev().find(|edit|edit.at==at) {edit.after=after;}
    }

    fn shared_format_dxf(&mut self, format: &str, automatic: bool) -> String {
        let known=if automatic {WorkbookFiles::revision_number_format(format)} else {
            match format {
                "General" => Some((0,"General")), "0" => Some((1,"0")),
                "0.00" => Some((2,"0.00")), "#,##0" => Some((3,"#,##0")),
                "@" => Some((30,"@")),
                "#,##0.00" => Some((4,"#,##0.00")),
                MONEY_FORMAT | "m/d/yyyy" | "0%" | "0.00%" => WorkbookFiles::revision_number_format(format),
                _ => None,
            }
        };
        let (id,code)=if let Some(known)=known {known} else {
            if self.file_io.shared_format_ids.is_empty() {
                if let Some(package)=&self.file_io.package {
                    if let Ok(mut archive)=oxidocs_common::archive::OoxmlArchive::new(package) {
                        if let Ok(styles)=archive.read_part("xl/styles.xml") {
                            if let Ok(formats)=WorkbookFiles::xml_elements(&styles,b"numFmt",2) {
                                for held in formats {
                                    if let (Some(code),Some(id))=(held.get("formatCode"),held.get("numFmtId").and_then(|v|v.parse::<u32>().ok())) {
                                        self.file_io.shared_format_ids.insert(code.clone(),id);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            let next=self.file_io.shared_format_ids.values().copied().filter(|id|*id>=164).max()
                .map_or(164,|id|id.saturating_add(1));
            let id=*self.file_io.shared_format_ids.entry(format.into()).or_insert(next);
            (id,format)
        };
        format!("<numFmt numFmtId=\"{id}\" formatCode=\"{}\"/>",quick_xml::escape::escape(code))
    }

    pub(super) fn record_shared_automatic_format(&mut self, at: CellAddress, format: &str) {
        if !self.file_io.shared {return;}
        let key=(at.sheet,at.row,at.column);
        let Some(input)=self.file_io.shared_input_formats.get_mut(&key) else {return;};
        input.after=format.into();input.changed=input.before!=format;
        let (sequence,changed)=(input.sequence,input.changed);
        let format_metadata=input.clone();
        if let Some(edit)=self.file_io.shared_cell_edits.last_mut().filter(|edit|edit.at==at && edit.format.sequence==sequence) {
            edit.format=format_metadata;
        }
        if changed {
            let dxf=self.shared_format_dxf(format,true);
            self.file_io.shared_format_edits.push(SharedFormatEdit {
                start:0,sequence,range:CellRange::single(at),references:None,length:Some(0),dxf,automatic:true,
            });
            self.file_io.shared_automatic_formats.insert(key);
        }
    }

    fn shared_character_format(&self, at: CellAddress, start: u32) -> CellStyle {
        let Some(cell)=self.cell_here(at.sheet,at.row,at.column) else {return self.template_style(at);};
        let mut style=cell.style.clone();
        if let Some(dress)=dresses_of(&cell).get(start as usize) {dress.write_font(&mut style);}
        style
    }

    pub(super) fn shared_format_request(&self, receiver: &ObjectRef, name: &str, value: &Value) -> Option<SharedFormatRequest> {
        if !self.file_io.shared || self.gone(receiver) || matches!(value,Value::Null) {return None;}
        let lower=name.to_ascii_lowercase();
        if matches!(value, Value::Array(_)) && matches!(lower.as_str(),"locked"|"formulahidden") {return None;}
        if let Some((at,start,length))=self.characters_font_of(receiver) {
            let property=match lower.as_str() {
                "name"=>"charfontname","size"=>"fontsize","bold"=>"bold","italic"=>"italic",
                "underline"=>"underline","color"|"colorindex"|"themecolor"|"tintandshade"=>"fontcolor",
                "themefont"=>"themefont",_=>return None,
            };
            let text=self.cell_here(at.sheet,at.row,at.column).map(|cell|cell_text_of(&cell)).unwrap_or_default();
            let (from,to)=font_character_span(start,length,&text);
            let before=self.shared_character_format(at,from as u32);
            return Some(SharedFormatRequest {range:CellRange::single(at),property,before,mixed:false,span:Some((from as u32,(to-from) as u32))});
        }
        let (range,property)=if let Some(range)=self.range_font(receiver) {
            (range,match lower.as_str() {"bold"=>"bold","italic"=>"italic","color"|"colorindex"|"themecolor"=>"fontcolor","tintandshade"=>"fonttint",
                "size"=>"fontsize","name"=>"fontname","themefont"=>"themefont","underline"=>"underline","superscript"|"subscript"=>"fontposition","strikethrough"=>"strike",_=>return None})
        } else if let Some(range)=self.range_interior(receiver) {
            (range,match lower.as_str() {
                "pattern"=>"fillpattern","colorindex"=>"fillindex",
                "color"|"themecolor"|"tintandshade"|"patterncolor"|"patterncolorindex"|
                "patternthemecolor"|"patterntintandshade"=>"fill",_=>return None})
        } else if let Some((range,_))=self.range_borders(receiver) {
            (range,match lower.as_str() {"linestyle"|"weight"|"color"|"colorindex"=>"border",_=>return None})
        } else {
            (self.range(receiver)?,match lower.as_str() {
                "style"=>"style","numberformat"|"numberformatlocal"=>"numberformat","horizontalalignment"=>"horizontal",
                "verticalalignment"=>"vertical","wraptext"=>"wrap","shrinktofit"=>"shrink","readingorder"=>"readingorder","orientation"=>"rotation","indentlevel"=>"indent","addindent"=>"justify","locked"=>"locked","formulahidden"=>"hidden",_=>return None})
        };
        let at=range.first();
        let before=self.cell_here(at.sheet,at.row,at.column).map(|cell|cell.style).unwrap_or_else(||self.template_style(at));
        let mixed = match property {
            "hidden" => matches!(self.uniform_style(range, |style| style.formula_hidden), Ok(None)),
            "locked" => matches!(self.uniform_style(range, |style| style.unlocked), Ok(None)),
            _ => false,
        };
        Some(SharedFormatRequest {range,property,before,mixed,span:None})
    }

    pub(super) fn finish_shared_format(&mut self, request: SharedFormatRequest) {
        let SharedFormatRequest {range,property,before,mixed,span}=request;
        let at=range.first();
        let after=span.map_or_else(||self.cell_here(at.sheet,at.row,at.column).map(|cell|cell.style).unwrap_or_else(||self.template_style(at)),
            |(start,_)|self.shared_character_format(at,start));
        let (length,dxf)=match property {
            "style" => {
                // Native's Normal application records the full style even
                // when only one loaded font attribute visibly changes.
                if self.styled.contains_key(&at) {return;}
                (None,WorkbookFiles::normal_style_revision_dxf(&after,self.named_normal_style().map_or(0,|style|style.number_format_id)))
            },

            "numberformat" => {
                if before.number_format==after.number_format {return;}
                let code=after.number_format.as_deref().unwrap_or("General");
                let dxf=self.shared_format_dxf(code,false);
                self.file_io.shared_automatic_formats.retain(|&(sheet,row,column)|
                    sheet!=range.sheet || row<range.start_row || row>range.end_row || column<range.start_column || column>range.end_column);
                (None,dxf)
            }
            "underline" => {
                let old=before.effective_underline_kind();let new=after.effective_underline_kind();
                let xml=if old==new {String::new()} else {
                    match new {
                        Some(oxicells_core::ir::CellUnderlineKind::Single) => "<font><u/></font>".into(),
                        Some(kind) => format!("<font><u val=\"{}\"/></font>",kind.as_str()),
                        None => "<font><u val=\"none\"/></font>".into(),
                    }
                };
                (Some(i32::MAX as u32),xml)
            },
            "fontposition" => {
                let xml=if before.font_position==after.font_position {String::new()} else {
                    format!("<font><vertAlign val=\"{}\"/></font>",after.font_position.map_or("baseline",|position|position.as_str()))
                };
                (Some(i32::MAX as u32),xml)
            },
            "bold"|"italic"|"strike" => {
                let (old,new,tag)=match property {
                    "bold" => (before.bold,after.bold,"b"), "italic" => (before.italic,after.italic,"i"),
                    "underline" => (before.underline,after.underline,"u"),
                    _ => (before.strikethrough,after.strikethrough,"strike"),
                };
                let off=if tag=="u" {"none"} else {"0"};
                (Some(i32::MAX as u32),if old==new {String::new()} else if new {format!("<font><{tag}/></font>")}
                    else {format!("<font><{tag} val=\"{off}\"/></font>")})
            }
            "fontsize" => (Some(i32::MAX as u32),if before.font_size==after.font_size {String::new()} else {
                after.font_size.map_or_else(String::new,|value|format!("<font><sz val=\"{value}\"/></font>"))
            }),
            "fontname"|"charfontname" => (Some(i32::MAX as u32),if before.latin_font_name()==after.latin_font_name()
                && before.font_family==after.font_family && before.font_charset==after.font_charset {String::new()} else {
                let mut font=String::from("<font>");
                if let Some(name)=after.latin_font_name() {font.push_str(&format!("<name val=\"{}\"/>",quick_xml::escape::escape(name)));}
                if before.font_family!=after.font_family {
                    if let Some(family)=after.font_family {font.push_str(&format!("<family val=\"{family}\"/>"));}
                }
                if before.font_charset!=after.font_charset {
                    if let Some(charset)=after.font_charset {font.push_str(&format!("<charset val=\"{charset}\"/>"));}
                }
                let scheme=after.font_reference.as_ref().and_then(|r|r.scheme.as_deref()).unwrap_or("none");
                let old_scheme=before.font_reference.as_ref().and_then(|r|r.scheme.as_deref()).unwrap_or("none");
                if property!="charfontname" || old_scheme!=scheme {
                    font.push_str(&format!("<scheme val=\"{}\"/>",quick_xml::escape::escape(scheme)));
                }
                font.push_str("</font>");font
            }),
            "themefont" => {
                let mut font=String::from("<font>");
                if let Some(scheme)=after.font_reference.as_ref().and_then(|r|r.scheme.as_deref()).filter(|s|*s!="none") {
                    if let Some(name)=after.latin_font_name() {font.push_str(&format!("<name val=\"{}\"/>",quick_xml::escape::escape(name)));}
                    font.push_str(&format!("<scheme val=\"{}\"/>",quick_xml::escape::escape(scheme)));
                }
                font.push_str("</font>");(Some(0),font)
            },
            "fonttint" => {
                let old=WorkbookFiles::revision_font_color(&before);let new=WorkbookFiles::revision_font_color(&after);
                let family=after.font_family.map_or_else(String::new,|family|format!("<family val=\"{family}\"/>"));
                (None,if old==new {String::new()} else {format!("<font>{new}{family}</font>")})
            },
            "fontcolor" => {
                let old=WorkbookFiles::revision_font_color(&before);let new=WorkbookFiles::revision_font_color(&after);
                (Some(0),if old==new {String::new()} else {format!("<font>{new}</font>")})
            },
            "fill"|"fillpattern"|"fillindex" => {
                let old=interior_fill(&before);let new=interior_fill(&after);
                if property=="fillpattern" && old.pattern==new.pattern {return;}
                let pattern=if old.pattern!=new.pattern {
                    format!(" patternType=\"{}\"",quick_xml::escape::escape(&new.pattern))
                } else {String::new()};
                let mut colors=String::new();
                if property!="fillpattern" {
                    for (pattern_color,tag,fallback) in [(true,"fgColor",64),(false,"bgColor",65)] {
                        let old_color=interior_color_token(&old,pattern_color);
                        let new_color=interior_color_token(&new,pattern_color);
                        if old_color!=new_color {
                            colors.push_str(&oxicells_core::editor::cell_fill_color_xml(tag,new_color.as_ref(),Some(fallback)));
                        }
                    }
                }
                if pattern.is_empty() && colors.is_empty() {return;}
                (Some(0),format!("<fill><patternFill{pattern}>{colors}</patternFill></fill>"))
            },
            "border" => (Some(0),WorkbookFiles::revision_border(&after,&["left","right","top","bottom"],false)),
            "horizontal" => {
                if before.horizontal_align.as_deref().unwrap_or("general")==after.horizontal_align.as_deref().unwrap_or("general") {return;}
                (None,format!("<alignment horizontal=\"{}\"/>",after.horizontal_align.as_deref().unwrap_or("general")))
            },
            "vertical" => {
                if before.vertical_align.as_deref().unwrap_or("bottom")==after.vertical_align.as_deref().unwrap_or("bottom") {return;}
                (None,format!("<alignment vertical=\"{}\"/>",after.vertical_align.as_deref().unwrap_or("bottom")))
            },
            "wrap" => {
                if before.wrap_text==after.wrap_text {return;}
                (None,format!("<alignment wrapText=\"{}\"/>",u8::from(after.wrap_text)))
            },
            "hidden" | "locked" => {
                let (old, new, attribute) = if property == "hidden" {
                    (before.formula_hidden, after.formula_hidden, "hidden")
                } else {
                    (!before.unlocked, !after.unlocked, "locked")
                };
                if !mixed && old == new { return; }
                (None, format!("<protection {attribute}=\"{}\"/>", u8::from(new)))
            },
            "rotation" => {
                if before.text_rotation==after.text_rotation && before.stacked_text==after.stacked_text {return;}
                let angle=after.text_rotation.unwrap_or(0);
                let encoded=if after.stacked_text {255} else if angle<0 {90-angle} else {angle};
                (None,format!("<alignment textRotation=\"{encoded}\"/>"))
            },
            "justify" => {
                if before.justify_last_line==after.justify_last_line {return;}
                (None,format!("<alignment justifyLastLine=\"{}\"/>",u8::from(after.justify_last_line)))
            },
            "indent" => {
                if before.indent==after.indent && before.horizontal_align==after.horizontal_align {return;}
                let mut attributes=String::new();
                if before.horizontal_align.as_deref().unwrap_or("general")!=after.horizontal_align.as_deref().unwrap_or("general") {
                    attributes.push_str(&format!(" horizontal=\"{}\"",after.horizontal_align.as_deref().unwrap_or("general")));
                }
                if before.indent!=after.indent {attributes.push_str(&format!(" indent=\"{}\"",after.indent));}
                (None,format!("<alignment{attributes}/>"))
            },
            "readingorder" => {
                if before.reading_order==after.reading_order {return;}
                let order=match after.reading_order {
                    oxicells_core::ir::CellReadingOrder::Contextual=>0,
                    oxicells_core::ir::CellReadingOrder::LeftToRight=>1,
                    oxicells_core::ir::CellReadingOrder::RightToLeft=>2,
                };
                (None,format!("<alignment readingOrder=\"{order}\"/>"))
            },
            "shrink" => {
                if before.shrink_to_fit==after.shrink_to_fit {return;}
                (None,format!("<alignment shrinkToFit=\"{}\"/>",u8::from(after.shrink_to_fit)))
            },
            _ => return,
        };
        // A copy can retain an empty old-format record when a later
        // explicit format replaces its pasted format before the real save.
        if before!=after || mixed {
            let within=|at: CellAddress|at.sheet==range.sheet && at.row>=range.start_row && at.row<=range.end_row
                && at.column>=range.start_column && at.column<=range.end_column;
            for (&(sheet,row,column),input) in &mut self.file_io.shared_input_formats {
                if within(CellAddress {sheet,row,column}) && input.style_pasted && input.style_dxf.is_none() {
                    input.style_dxf=Some((String::new(),String::new()));
                }
            }
            for edit in &mut self.file_io.shared_cell_edits {
                if within(edit.at) && edit.format.style_pasted && edit.format.style_dxf.is_none() {
                    edit.format.style_dxf=Some((String::new(),String::new()));
                }
            }
        }
        self.file_io.shared_revision_sequence=self.file_io.shared_revision_sequence.saturating_add(1);
        self.file_io.shared_format_edits.push(SharedFormatEdit {
            start:span.map_or(0,|v|v.0),sequence:self.file_io.shared_revision_sequence,range,references:None,
            length:span.map(|v|v.1).or(length),dxf,automatic:false,
        });
    }

    pub(super) fn coalesce_shared_reading_order_edits(&mut self, from: usize, areas: &[CellRange], include_unchanged: bool) {
        let edits=&self.file_io.shared_format_edits[from..];
        let Some(first)=edits.first() else {return;};
        if (!include_unchanged && edits.len()!=areas.len()) || edits.len()>areas.len() || edits.iter().any(|edit|
            edit.dxf!=first.dxf || edit.length!=first.length || edit.start!=first.start
            || edit.automatic || edit.range.sheet!=first.range.sheet) {return;}
        let mut combined=first.clone();
        combined.references=Some(areas.iter().map(|range|format_range_address(*range,false,false))
            .collect::<Vec<_>>().join(" "));
        self.file_io.shared_format_edits.truncate(from);
        self.file_io.shared_format_edits.push(combined);
    }

    pub(super) fn record_shared_border_around(&mut self, range: CellRange, line: bool, colour: bool) {
        if !self.file_io.shared {return;}
        let at=range.first();
        let style=self.cell_here(at.sheet,at.row,at.column).map(|cell|cell.style).unwrap_or_else(||self.template_style(at));
        let dxf=WorkbookFiles::revision_border(&style,&["left","right","top","bottom"],true);
        for length in line.then_some(Some(0)).into_iter().chain(colour.then_some(None)) {
            self.file_io.shared_revision_sequence=self.file_io.shared_revision_sequence.saturating_add(1);
            self.file_io.shared_format_edits.push(SharedFormatEdit {
                start:0,sequence:self.file_io.shared_revision_sequence,range,references:None,length,dxf:dxf.clone(),automatic:false,
            });
        }
    }

    pub(super) fn finish_shared_replace_paint(&mut self, from: usize) {
        // Replace's paint record applies to the cell, without a character span.
        for edit in self.file_io.shared_format_edits.iter_mut().skip(from) {
            if edit.dxf.starts_with("<font><color ") {
                edit.start=0;edit.length=None;
            }
        }
    }

    fn sync_links_for_save(&mut self) {
        for sheet in &mut self.workbook.sheets {
            sheet.hyperlinks.clear();
        }
        for link in &self.links {
            let Some(sheet) = self.workbook.sheets.get_mut(link.anchor.sheet) else { continue };
            let internal = link.address.is_empty();
            sheet.hyperlinks.push(oxicells_core::ir::Hyperlink {
                cell: (link.anchor.start_row.saturating_sub(1), link.anchor.start_column),
                end_cell: ((link.anchor.end_row, link.anchor.end_column)
                    != (link.anchor.start_row, link.anchor.start_column))
                    .then_some((link.anchor.end_row.saturating_sub(1), link.anchor.end_column)),
                target: if internal { link.sub_address.clone() } else { link.address.clone() },
                internal,
                sub_address: (!internal && !link.sub_address.is_empty()).then(|| link.sub_address.clone()),
                tooltip: (!link.screen_tip.is_empty()).then(|| link.screen_tip.clone()),
            });
        }
    }

    fn copy_method_failure(&self) -> String {
        let description = match self.error_language {
            ErrorLanguage::Japanese => "'SaveCopyAs' メソッドは失敗しました: '_Workbook' オブジェクト",
            ErrorLanguage::English => "Method 'SaveCopyAs' of object '_Workbook' failed",
        };
        host_error_described(1004, description)
    }

    fn copy_name_failure(&self, path: &str) -> String {
        let description = match self.error_language {
            ErrorLanguage::Japanese => format!("申し訳ございません。{path}が見つかりません。名前が変更されたか、移動や削除が行われた可能性があります。"),
            ErrorLanguage::English => format!("Sorry, we couldn't find {path}. Is it possible it was moved, renamed or deleted?"),
        };
        oxivba_core::host_error_from(1004, "Microsoft Excel", description)
    }

    pub(super) fn copy_access_failure(&self, path: &str) -> String {
        let description = match self.error_language {
            ErrorLanguage::Japanese => format!("ファイル '{path}' にアクセスできません。次のいずれかの理由が考えられます。\n\n• ファイル名またはパスが存在しません。\n• ファイルが他のプログラムによって使用されています。\n• 保存しようとしているブックと同じ名前のブックが現在開かれています。"),
            ErrorLanguage::English => format!("Cannot access '{path}'. Possible reasons include a missing file or path, a file in use by another program, or an open workbook with the same name."),
        };
        oxivba_core::host_error_from(1004, "Microsoft Excel", description)
    }

    /// Save the closing workbook through the shared virtual filesystem.
    /// False means calculation requested VBA callbacks and Close must retry.
    pub(super) fn save_for_close(&mut self, filename: Option<&Value>, files: &mut Files) -> Result<bool, String> {
        self.save_in_format(filename, None, files)
    }

    pub(super) fn save_in_format(&mut self, filename: Option<&Value>, format: Option<i64>, files: &mut Files) -> Result<bool, String> {
        self.save_in_format_with_preferences(filename, format, None, files)
    }

    pub(super) fn save_in_format_with_preferences(&mut self, filename: Option<&Value>, format: Option<i64>,
        preferences: Option<(bool, bool, bool)>, files: &mut Files) -> Result<bool, String> {
        let path = match filename {
            None | Some(Value::Missing | Value::Null) => self.file_io.full_name.clone()
                .ok_or_else(|| host_error(1004, "Workbook.Close requires a filename for an unsaved workbook"))?,
            Some(Value::String(path)) => path.clone(),
            Some(Value::Utf16(units)) => String::from_utf16(units)
                .map_err(|_| self.copy_name_failure(&String::from_utf16_lossy(units)))?,
            Some(Value::Empty | Value::NullString) => String::new(),
            Some(value) => shown_text(value, None),
        };
        if path.is_empty() {
            let description = match self.error_language {
                ErrorLanguage::Japanese => "ファイルにアクセスできませんでした。次のいずれかを確認してみてください:\n\n• 指定したフォルダーが存在することを確認します。\n• ファイルが保存されているフォルダーが読み取り専用でないことを確認します。\n• ファイル名とフォルダーのパスに次の文字が含まれていないことを確認します: <  >  ?  [  ]  :  | *\n• ファイル名とフォルダーのパスが 218 文字以下になっていることを確認します。".to_string(),
                ErrorLanguage::English => "The file could not be accessed. Check that the folder exists, is writable, and the filename and path are valid.".to_string(),
            };
            return Err(oxivba_core::host_error_from(1004, "Microsoft Excel", description));
        }
        let normalized = files.absolute_path(&path);
        if self.file_io.read_only && self.file_io.full_name.as_deref()
            .is_some_and(|current| current.eq_ignore_ascii_case(&normalized)) {
            return Err(host_error(1004, "Workbook.Close requires another filename for a read-only workbook"));
        }
        self.save_copy_path(&path, format, preferences, false, files).map_err(|error| match error {
            CopyFailure::Method(error) | CopyFailure::Excel(error) => error,
        })?;
        if !self.user_functions.borrow().pending.is_empty() { return Ok(false); }
        self.file_io.full_name = Some(normalized.clone());
        self.file_name = Some(normalized.rsplit('\\').next().unwrap_or(&normalized).to_string());
        self.file_io.package = Some(files.bytes(&normalized).map_err(|error| error.description().to_string())?.to_vec());
        self.file_io.imported_package = false;
        if let Some(format) = format { self.file_io.format = Some(format as f64); }
        self.file_io.read_only = false;
        self.file_io.shared_input_formats.clear();
        self.file_io.shared_automatic_formats.clear();
        self.file_io.shared_format_edits.clear();
        self.file_io.shared_cell_edits.clear();
        self.file_io.shared_revision_sequence=0;
        self.saved = true;
        Ok(true)
    }

    pub(super) fn save_copy_as(&mut self, args: &[Value], files: &mut Files) -> Result<Option<Value>, String> {
        if args.len() > 1 { return Err(host_error(450, "SaveCopyAs accepts one Filename")); }
        let input = args.first().ok_or_else(|| self.copy_method_failure())?;
        // Resolve the complete argument before publishing files. A bad array
        // conversion rejects the whole call, including earlier valid elements.
        let inputs = self.copy_inputs(input)?;
        if !self.user_functions.borrow().pending.is_empty() { return Ok(Some(Value::Empty)); }
        let mut method_failure = None;
        let mut excel_failure = None;
        for input in inputs {
            let result = match input {
                CopyInput::Filename(path) => self.save_copy_path(&path, None, None, true, files),
                CopyInput::Error => Err(CopyFailure::Method(self.copy_method_failure())),
            };
            match result {
                Err(CopyFailure::Method(error)) => { method_failure.get_or_insert(error); }
                Err(CopyFailure::Excel(error)) => { excel_failure.get_or_insert(error); }
                Ok(()) => {}
            }
            if !self.user_functions.borrow().pending.is_empty() { return Ok(Some(Value::Empty)); }
        }
        // File-specific failures keep the first concrete diagnostic, while
        // other destinations still receive their copies. Error-valued elements
        // supply the generic method diagnostic only if no concrete one exists.
        if let Some(error) = excel_failure.or(method_failure) { return Err(error); }
        Ok(Some(Value::Empty))
    }

    fn copy_inputs(&mut self, input: &Value) -> Result<Vec<CopyInput>, String> {
        let mut pending = vec![(input.clone(), false)];
        let mut inputs = Vec::new();
        while let Some((value, in_array)) = pending.pop() {
            let path = match value {
                Value::Array(array) => {
                    if array.values.is_empty() { return Err(host_error(13, "Filename array is empty")); }
                    pending.extend(array.values.into_iter().rev().map(|value| (value, true)));
                    continue;
                }
                Value::Object(object) => {
                    let range = self.range(&object).or_else(|| self.blocks(&object).and_then(|blocks| blocks.first()).copied());
                    if let Some(range) = range {
                        self.settle(range);
                        pending.push((self.range_value(range)?, in_array));
                        continue;
                    }
                    if in_array && matches!(object.kind.as_str(), "Collection" | "Dictionary") {
                        return Err(host_error(450, "Filename object's default Item requires an index"));
                    }
                    return Err(self.copy_method_failure());
                }
                Value::Null if in_array => return Err(host_error(13, "Filename array contains Null")),
                Value::Null | Value::Missing | Value::Record(_) => return Err(self.copy_method_failure()),
                Value::Nothing => return Err(host_error(5, "Filename is a null object reference")),
                Value::Error(_) => { inputs.push(CopyInput::Error); continue; }
                Value::Empty | Value::NullString => continue,
                Value::String(path) => path,
                Value::Utf16(units) => String::from_utf16_lossy(&units),
                Value::Boolean(value) => if value { "TRUE" } else { "FALSE" }.to_string(),
                Value::Currency(value) => {
                    let cents=(value.unsigned_abs()+50)/100;
                    let whole=(cents/100).to_string();
                    let mut grouped=String::new();
                    for (index,ch) in whole.chars().enumerate() {
                        if index>0 && (whole.len()-index).is_multiple_of(3) { grouped.push(','); }
                        grouped.push(ch);
                    }
                    let currency=format!("{}{grouped}.{:02}",format_debug_value(&international(25)?),cents%100);
                    if value<0 { format!("({currency})") } else { currency }
                },
                value => format_debug_value(&value),
            };
            if !path.is_empty() { inputs.push(CopyInput::Filename(path)); }
        }
        Ok(inputs)
    }

    fn save_copy_path(&mut self, path: &str, format: Option<i64>,
        preferences: Option<(bool, bool, bool)>, copy_only: bool, files: &mut Files) -> Result<(), CopyFailure> {
        if path.encode_utf16().count() > 255 { return Err(CopyFailure::Method(self.copy_method_failure())); }
        let invalid = path.chars().enumerate().any(|(index, ch)| {
            matches!(ch, '?' | '*' | '"' | '<' | '>' | '|') || ch == '\0' || ch == ':' && index != 1
        });
        if invalid { return Err(CopyFailure::Excel(self.copy_name_failure(path))); }
        if path.ends_with(['\\', '/']) { return Err(CopyFailure::Excel(self.copy_access_failure(path))); }
        let normalized = files.absolute_path(path);
        if let Some((parent, _)) = normalized.rsplit_once('\\') {
            if !parent.is_empty() && !files.folder_exists(parent) {
                // Shared real saves first try an Excel-style temporary file.
                // A failed creation leaves the saved package and history intact.
                let inaccessible = if self.file_io.shared && !copy_only {
                    let guid=WorkbookFiles::shared_guid().map_err(|_|CopyFailure::Method(self.copy_method_failure()))?;
                    format!("{parent}\\{}", &guid[1..9])
                } else { path.to_string() };
                return Err(CopyFailure::Excel(self.copy_access_failure(&inaccessible)));
            }
        }
        let Some(package) = self.file_io.package.as_ref() else { return Err(CopyFailure::Method(self.copy_method_failure())); };
        let package = if self.file_io.imported_package {
            let mut archive = oxidocs_common::archive::OoxmlArchive::new(package)
                .map_err(|_| CopyFailure::Method(self.copy_method_failure()))?;
            if let Some(styles) = archive.try_read_part("xl/styles.xml")
                .map_err(|_| CopyFailure::Method(self.copy_method_failure()))? {
                let imported = oxicells_core::parser::import_styles_xml(&styles)
                    .map_err(|_| CopyFailure::Method(self.copy_method_failure()))?;
                WorkbookFiles::replace_package_parts(package,
                    BTreeMap::from([("xl/styles.xml".into(), imported.into_bytes())]), &BTreeSet::new())
                    .map_err(|_| CopyFailure::Method(self.copy_method_failure()))?
            } else { package.clone() }
        } else { package.clone() };
        let saved = {
            let mut operation = self.user_functions.borrow_mut();
            *operation.copy_saved.get_or_insert(self.saved)
        };
        if self.calculation != -4135 || self.application_setting("calculatebeforesave") == Some(Value::Boolean(true)) {
            self.recalculate();
            self.saved = saved;
            // Runtime calls the requested UDFs, resumes calculation and retries
            // this same member before execution advances to the next statement.
            if !self.user_functions.borrow().pending.is_empty() { return Ok(()); }
        }
        self.sync_calculation_settings();
        self.sync_band_styles_to_workbook();
        self.sync_legacy_notes_for_save();
        self.sync_links_for_save();
        let mut editor = oxicells_core::XlsxEditor::new(&package).map_err(|_| CopyFailure::Method(self.copy_method_failure()))?;
        editor.apply_workbook(&self.workbook).map_err(|_| CopyFailure::Method(self.copy_method_failure()))?;
        let mut bytes = editor.save().map_err(|_| CopyFailure::Method(self.copy_method_failure()))?;
        let failure = || CopyFailure::Method(self.copy_method_failure());
        let project = crate::vba_project::retain_workbook_project(&bytes).map_err(|_| failure())?;
        if project.is_none() && self.file_io.retained_project.is_some() {
            bytes = crate::vba_project::convert_workbook_format(&bytes, 52, self.file_io.retained_project.as_ref())
                .map_err(|_| failure())?;
        }
        if project.is_some() || self.file_io.retained_project.is_some() {
            let source_code_page = self.automation_code_page.map(|page| match page {
                AnsiCodePage::Japanese => 932, AnsiCodePage::Western => 1252,
            });
            bytes = crate::vba_project::save_workbook_project(&bytes, &self.file_io.source, source_code_page).map_err(|_| CopyFailure::Method(self.copy_method_failure()))?;
            bytes = crate::vba_project::compact_project_storage(&bytes).map_err(|_| CopyFailure::Method(self.copy_method_failure()))?;
        }
        let retained_project = crate::vba_project::retain_workbook_project(&bytes)
            .map_err(|_| CopyFailure::Method(self.copy_method_failure()))?;
        let format = format.or(self.file_io.format.map(|value| value as i64)).unwrap_or(51);
        if matches!(format, 51 | 54) || self.file_io.format != Some(format as f64) {
            bytes = crate::vba_project::convert_workbook_format(&bytes, format, retained_project.as_ref())
                .map_err(|_| CopyFailure::Method(self.copy_method_failure()))?;
        }
        let (recommended, backup, shared) = preferences.unwrap_or((
            self.file_io.read_only_recommended, self.file_io.create_backup, self.file_io.shared));
        bytes = WorkbookFiles::package_with_preferences(&bytes, recommended, backup).map_err(|_| failure())?;
        if shared {
            let author = self.application_setting("username").unwrap_or_else(||Value::String("User".into()));
            let units = author.string_units().map(|units|units.into_owned()).unwrap_or_else(||"User".encode_utf16().collect());
            let operation=if copy_only {SharedSaveOperation::Copy}
                else if preferences.is_some() {SharedSaveOperation::SaveAs} else {SharedSaveOperation::Save};
            bytes = WorkbookFiles::package_with_revision_history(&bytes, &package, &self.workbook, &units, self.now, operation, &self.file_io).map_err(|_|failure())?;
            bytes = WorkbookFiles::package_with_shared_formats(&bytes).map_err(|_|failure())?;
        } else {
            bytes = WorkbookFiles::package_without_sharing(&bytes).map_err(|_|failure())?;
        }
        // Save/SaveAs back up the previous destination, not the new bytes.
        // SaveCopyAs preserves the preference but does not create a backup.
        let backup_file = if backup && !copy_only {
            files.bytes(&normalized).ok().map(|previous| {
                let (parent, leaf) = normalized.rsplit_once('\\').unwrap_or(("", &normalized));
                let stem = leaf.rsplit_once('.').map_or(leaf, |(stem, _)| stem);
                let name = match self.error_language {
                    ErrorLanguage::Japanese => format!("{stem} のバックアップ.xlk"),
                    ErrorLanguage::English => format!("Backup of {stem}.xlk"),
                };
                (format!("{parent}\\{name}"), previous.to_vec())
            })
        } else { None };
        let result = if let Some((backup_path, previous)) = &backup_file {
            let mut next = files.clone();
            next.replace_bytes(backup_path, previous.clone())
                .and_then(|()| next.replace_bytes(&normalized, bytes))
                .map(|()| *files = next)
        } else { files.replace_bytes(&normalized, bytes) };
        result.map_err(|error| match error {
            FileError::AlreadyOpen | FileError::PathNotFound => CopyFailure::Excel(self.copy_access_failure(path)),
            _ => CopyFailure::Method(self.copy_method_failure()),
        })?;
        if let Some((path, _)) = backup_file { self.file_io.exports.insert(path); }
        self.file_io.read_only_recommended = recommended;
        self.file_io.create_backup = backup;
        self.file_io.shared = shared;
        self.file_io.retained_project = retained_project;
        self.file_io.exports.insert(normalized);
        Ok(())
    }
}
