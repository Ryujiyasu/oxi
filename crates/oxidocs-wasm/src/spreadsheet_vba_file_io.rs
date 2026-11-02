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

#[derive(Clone, Default)]
pub(super) struct WorkbookFiles {
    pub package: Option<Vec<u8>>,
    pub full_name: Option<String>,
    pub read_only: bool,
    pub format: Option<f64>,
    pub source: String,
    pub files: Files,
    pub exports: BTreeSet<String>,
    mounted_folders: BTreeSet<String>,
    mounted_files: BTreeSet<String>,
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
        font: (&str, f32),
        dimensions: (f32, f32),
    ) -> Result<Vec<u8>, String> {
        use std::io::Write;
        use oxicells_core::ir::SheetKind;
        let main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        let rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        let xlm = "http://schemas.microsoft.com/office/excel/2006/main";
        let escape = |text: &str| quick_xml::escape::escape(text).into_owned();
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
        parts.push(("xl/styles.xml".into(), format!(r#"<styleSheet xmlns="{main}"><fonts count="1"><font><sz val="{}"/><name val="{}"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>"#, font.1, escape(font.0))));
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

    pub(super) fn set_package(&mut self, package: Vec<u8>) -> Result<(), String> {
        // Reject invalid packages before replacing a loaded host's source.
        oxicells_core::XlsxEditor::new(&package).map_err(|error| error.to_string())?;
        let mut archive = oxidocs_common::archive::OoxmlArchive::new(&package).map_err(|error| error.to_string())?;
        let types = archive.read_part("[Content_Types].xml").map_err(|error| error.to_string())?.to_ascii_lowercase();
        self.format = Some(if types.contains("application/vnd.ms-excel.template.macroenabled.main+xml") { 53.0 }
            else if types.contains("application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml") { 54.0 }
            else if types.contains("application/vnd.ms-excel.sheet.macroenabled.main+xml") { 52.0 }
            else { 51.0 });
        self.package = Some(package);
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
        if let Some(package) = context.package_bytes {
            self.set_package(package)?;
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

    fn copy_access_failure(&self, path: &str) -> String {
        let description = match self.error_language {
            ErrorLanguage::Japanese => format!("ファイル '{path}' にアクセスできません。次のいずれかの理由が考えられます。\n\n• ファイル名またはパスが存在しません。\n• ファイルが他のプログラムによって使用されています。\n• 保存しようとしているブックと同じ名前のブックが現在開かれています。"),
            ErrorLanguage::English => format!("Cannot access '{path}'. Possible reasons include a missing file or path, a file in use by another program, or an open workbook with the same name."),
        };
        oxivba_core::host_error_from(1004, "Microsoft Excel", description)
    }

    /// Save the closing workbook through the shared virtual filesystem.
    /// False means calculation requested VBA callbacks and Close must retry.
    pub(super) fn save_for_close(&mut self, filename: Option<&Value>, files: &mut Files) -> Result<bool, String> {
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
        self.save_copy_path(&path, files).map_err(|error| match error {
            CopyFailure::Method(error) | CopyFailure::Excel(error) => error,
        })?;
        if !self.user_functions.borrow().pending.is_empty() { return Ok(false); }
        self.file_io.full_name = Some(normalized.clone());
        self.file_name = Some(normalized.rsplit('\\').next().unwrap_or(&normalized).to_string());
        self.file_io.package = Some(files.bytes(&normalized).map_err(|error| error.description().to_string())?.to_vec());
        self.file_io.read_only = false;
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
                CopyInput::Filename(path) => self.save_copy_path(&path, files),
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

    fn save_copy_path(&mut self, path: &str, files: &mut Files) -> Result<(), CopyFailure> {
        if path.encode_utf16().count() > 255 { return Err(CopyFailure::Method(self.copy_method_failure())); }
        let invalid = path.chars().enumerate().any(|(index, ch)| {
            matches!(ch, '?' | '*' | '"' | '<' | '>' | '|') || ch == '\0' || ch == ':' && index != 1
        });
        if invalid { return Err(CopyFailure::Excel(self.copy_name_failure(path))); }
        if path.ends_with(['\\', '/']) { return Err(CopyFailure::Excel(self.copy_access_failure(path))); }
        let normalized = files.absolute_path(path);
        if let Some((parent, _)) = normalized.rsplit_once('\\') {
            if !parent.is_empty() && !files.folder_exists(parent) { return Err(CopyFailure::Excel(self.copy_access_failure(path))); }
        }
        let Some(package) = self.file_io.package.as_ref() else { return Err(CopyFailure::Method(self.copy_method_failure())); };
        let package = package.clone();
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
        if crate::vba_project::read_project_source_with_code_page(&bytes, None).map_err(|_| CopyFailure::Method(self.copy_method_failure()))?.is_some() {
            bytes = crate::vba_project::write_project_source(&bytes, &self.file_io.source, None).map_err(|_| CopyFailure::Method(self.copy_method_failure()))?;
            bytes = crate::vba_project::compact_project_storage(&bytes).map_err(|_| CopyFailure::Method(self.copy_method_failure()))?;
        }
        files.replace_bytes(&normalized, bytes).map_err(|error| match error {
            FileError::AlreadyOpen | FileError::PathNotFound => CopyFailure::Excel(self.copy_access_failure(path)),
            _ => CopyFailure::Method(self.copy_method_failure()),
        })?;
        self.file_io.exports.insert(normalized);
        Ok(())
    }
}
