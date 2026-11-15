// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use oxicells_core::ir::Workbook;
use oxidocs_wasm::{
    read_spreadsheet_vba_project_native, NativeVbaAction, NativeVbaCall, SpreadsheetVbaSession,
};
use serde_json::{json, Value};
use std::{
    fs,
    io::{BufRead, Write},
    path::Path,
    sync::atomic::{AtomicU64, Ordering},
};

fn write_package(path: &Path, bytes: &[u8]) -> Result<(), String> {
    static NEXT: AtomicU64 = AtomicU64::new(0);
    let parent = path.parent().unwrap_or(Path::new("."));
    for _ in 0..16 {
        let staging = parent.join(format!(
            ".oxicells-{}-{}.tmp",
            std::process::id(),
            NEXT.fetch_add(1, Ordering::Relaxed)
        ));
        let mut file = match fs::OpenOptions::new()
            .write(true)
            .create_new(true)
            .open(&staging)
        {
            Ok(file) => file,
            Err(error) if error.kind() == std::io::ErrorKind::AlreadyExists => continue,
            Err(error) => return Err(error.to_string()),
        };
        let written = file.write_all(bytes).and_then(|_| file.sync_all());
        drop(file);
        let result = written.and_then(|_| fs::rename(&staging, path));
        if result.is_err() {
            // Only remove the staging file that this invocation created exclusively.
            let _ = fs::remove_file(&staging);
        }
        return result.map_err(|error| error.to_string());
    }
    Err("could not create a unique save staging file".into())
}

struct OpenBook {
    session: SpreadsheetVbaSession,
    book: Option<Workbook>,
    source: String,
    initial_options: String,
    continued_options: String,
    initial_name: String,
    active_sheet: usize,
    seed: u64,
    loaded: bool,
}

fn fields(request: &Value, allowed: &[&str]) -> Result<(), String> {
    let object = request.as_object().ok_or("request must be a JSON object")?;
    if object.keys().any(|key| !allowed.contains(&key.as_str())) {
        return Err("unknown request field".into());
    }
    Ok(())
}

fn text<'a>(request: &'a Value, key: &str) -> Result<&'a str, String> {
    request
        .get(key)
        .and_then(Value::as_str)
        .ok_or_else(|| format!("{key} must be a string"))
}

fn clock(request: &Value) -> Result<f64, String> {
    request
        .get("currentTime")
        .and_then(Value::as_f64)
        .filter(|v| v.is_finite())
        .ok_or_else(|| "currentTime must be a finite local OLE date serial".into())
}

fn package_active_sheet(bytes: &[u8], count: usize) -> Result<usize, String> {
    let mut archive =
        oxidocs_common::archive::OoxmlArchive::new(bytes).map_err(|e| e.to_string())?;
    let xml = archive
        .read_part("xl/workbook.xml")
        .map_err(|e| e.to_string())?;
    let mut reader = quick_xml::Reader::from_str(&xml);
    loop {
        match reader.read_event().map_err(|e| e.to_string())? {
            quick_xml::events::Event::Start(e) | quick_xml::events::Event::Empty(e)
                if e.local_name().as_ref() == b"workbookView" =>
            {
                let at = oxidocs_common::xml_utils::get_attr(&e, "activeTab")
                    .map(|v| v.parse::<usize>())
                    .transpose()
                    .map_err(|e| e.to_string())?
                    .unwrap_or(0);
                return if at < count {
                    Ok(at)
                } else {
                    Err("active sheet index is out of bounds".into())
                };
            }
            quick_xml::events::Event::Eof => return Ok(0),
            _ => {}
        }
    }
}

impl OpenBook {
    fn open(request: &Value) -> Result<(Self, Value), String> {
        fields(
            request,
            &["id", "op", "path", "sourceCodePage", "seed", "currentTime"],
        )?;
        clock(request)?;
        let path = Path::new(text(request, "path")?);
        let bytes = fs::read(path).map_err(|e| e.to_string())?;
        let code_page = request
            .get("sourceCodePage")
            .filter(|v| !v.is_null())
            .map(|v| {
                v.as_u64()
                    .and_then(|v| u16::try_from(v).ok())
                    .ok_or("sourceCodePage must be a 16-bit integer")
            })
            .transpose()?;
        let seed = request
            .get("seed")
            .and_then(Value::as_u64)
            .ok_or("seed must be an unsigned 64-bit integer")?;
        let mut metadata: Value =
            serde_json::from_str(&read_spreadsheet_vba_project_native(&bytes, code_page)?)
                .map_err(|e| e.to_string())?;
        let source = metadata
            .as_object_mut()
            .unwrap()
            .remove("source")
            .and_then(|v| v.as_str().map(str::to_owned))
            .unwrap_or_default();
        let initial_name = path
            .file_name()
            .and_then(|v| v.to_str())
            .ok_or("filename must be valid UTF-8")?
            .to_owned();
        let full_name = format!("C:\\work\\{initial_name}");
        let mut continued = json!({});
        if let Some(name) = metadata.get("projectName").filter(|v| !v.is_null()) {
            continued["projectName"] = name.clone();
        }
        metadata["fileContext"] = json!({"packageBytes":bytes,"fullName":full_name,"folders":["C:\\work"],"sourceCodePage":code_page});
        metadata["applicationState"] = json!({"displayAlerts":false});
        let book = oxicells_core::parse_xlsx(&bytes).map_err(|e| e.to_string())?;
        if book.sheets.is_empty() {
            return Err("workbook contains no worksheets".into());
        }
        let active_sheet = package_active_sheet(&bytes, book.sheets.len())?;
        let response = json!({"protocol":1,"version":env!("CARGO_PKG_VERSION"),"hasMacros":!source.is_empty(),
            "projectName":metadata["projectName"],"sheets":book.sheets.iter().map(|sheet|&sheet.name).collect::<Vec<_>>()});
        Ok((
            Self {
                session: SpreadsheetVbaSession::default(),
                book: Some(book),
                source,
                initial_options: metadata.to_string(),
                continued_options: continued.to_string(),
                initial_name,
                active_sheet,
                seed,
                loaded: false,
            },
            response,
        ))
    }

    fn operation(&mut self, request: &Value) -> Result<Value, String> {
        fields(
            request,
            &[
                "id",
                "op",
                "procedure",
                "args",
                "reference",
                "member",
                "kind",
                "references",
                "path",
                "currentTime",
            ],
        )?;
        let current_time = clock(request)?;
        let operation = text(request, "op")?;
        let args_json = request
            .get("args")
            .cloned()
            .unwrap_or_else(|| json!([]))
            .to_string();
        let args = self.session.decode_native_arguments(&args_json)?;
        if !matches!(operation, "run" | "invoke") && !args.is_empty() {
            return Err("only run and invoke accept arguments".into());
        }
        let mut call = NativeVbaCall::new(&self.source, "", args, current_time, self.seed);
        let mut destination = None;
        call.action = match operation {
            "run" => NativeVbaAction::Run {
                procedure: text(request, "procedure")?,
            },
            "compile" => NativeVbaAction::CompileProject,
            "invoke" => {
                let reference = text(request, "reference")?;
                self.session.decode_native_arguments(
                    &json!([{"type":"object","value":{"reference":reference}}]).to_string(),
                )?;
                NativeVbaAction::Invoke {
                    reference: reference.into(),
                    member: text(request, "member")?.into(),
                    kind: request
                        .get("kind")
                        .and_then(Value::as_i64)
                        .filter(|kind| matches!(kind, 1 | 2 | 4 | 8))
                        .ok_or("kind must be an integer")?,
                }
            }
            "release" => {
                let references = request
                    .get("references")
                    .and_then(Value::as_array)
                    .ok_or("references must be an array")?
                    .iter()
                    .map(|value| {
                        value
                            .as_str()
                            .map(str::to_owned)
                            .ok_or("reference must be a string")
                    })
                    .collect::<Result<Vec<_>, _>>()?;
                for reference in &references {
                    self.session.decode_native_arguments(
                        &json!([{"type":"object","value":{"reference":reference}}]).to_string(),
                    )?;
                }
                NativeVbaAction::Release { references }
            }
            "close" if self.loaded => NativeVbaAction::ReleaseAll,
            "close" => return Ok(json!({"closed":true})),
            "save" => {
                let path = Path::new(text(request, "path")?);
                let format = match path
                    .extension()
                    .and_then(|v| v.to_str())
                    .map(str::to_ascii_lowercase)
                    .as_deref()
                {
                    Some("xlsx") => 51,
                    Some("xlsm") => 52,
                    Some("xltm") => 53,
                    Some("xltx") => 54,
                    _ => return Err("save requires .xlsx, .xlsm, .xltx or .xltm".into()),
                };
                if !path.parent().unwrap_or(Path::new(".")).is_dir() {
                    return Err("save destination folder does not exist".into());
                }
                let name = path
                    .file_name()
                    .and_then(|v| v.to_str())
                    .ok_or("filename must be valid UTF-8")?;
                destination = Some(path.to_owned());
                NativeVbaAction::SaveAs {
                    file_name: format!("C:\\work\\{name}"),
                    format,
                }
            }
            _ => return Err("unknown operation".into()),
        };
        call.active_sheet = self.active_sheet;
        call.options = if self.loaded {
            &self.continued_options
        } else {
            &self.initial_options
        };
        if !self.loaded {
            call.file_name = Some(self.initial_name.clone());
        }
        let book = self.book.take().ok_or("workbook session is unavailable")?;
        let outcome = self.session.run_native(book, call)?;
        let mut answer: Value =
            serde_json::from_str(&outcome.to_typed_json()?).map_err(|e| e.to_string())?;
        answer["sheets"] = json!(outcome
            .workbook()
            .sheets
            .iter()
            .map(|sheet| &sheet.name)
            .collect::<Vec<_>>());
        self.active_sheet = answer
            .get("active_sheet")
            .and_then(Value::as_u64)
            .unwrap_or(self.active_sheet as u64) as usize;
        self.loaded = true;
        let write_result = if let Some(path) = destination {
            if answer.get("error").is_some_and(|v| !v.is_null()) {
                Ok(())
            } else {
                outcome
                    .saved_package_bytes()
                    .ok_or_else(|| "save did not produce a package".to_string())
                    .and_then(|bytes| write_package(&path, bytes))
            }
        } else {
            Ok(())
        };
        self.book = Some(outcome.into_workbook());
        let object = answer.as_object_mut().ok_or("invalid native outcome")?;
        object.remove("workbook");
        object.remove("file_system_state");
        if let Some(save) = object.get_mut("save").and_then(Value::as_object_mut) {
            save.remove("package_bytes");
        }
        if operation == "save" {
            object.remove("files");
        }
        if let Some(books) = object
            .get_mut("other_workbooks")
            .and_then(Value::as_array_mut)
        {
            for book in books {
                if let Some(book) = book.as_object_mut() {
                    book.remove("workbook");
                }
            }
        }
        if let Some(selected) = object
            .get_mut("workbook_selection")
            .and_then(Value::as_object_mut)
        {
            selected.remove("package_bytes");
            selected.remove("file_system_state");
            if let Some(book) = selected.get_mut("selected").and_then(Value::as_object_mut) {
                book.remove("workbook");
            }
        }
        write_result?;
        Ok(answer)
    }
}

pub(super) fn serve(input: impl BufRead, mut output: impl Write) -> Result<(), String> {
    let mut book: Option<OpenBook> = None;
    for line in input.lines() {
        let line = line.map_err(|e| e.to_string())?;
        let parsed = serde_json::from_str::<Value>(&line);
        let id = parsed
            .as_ref()
            .ok()
            .and_then(|request| request.get("id"))
            .cloned()
            .unwrap_or(Value::Null);
        let close = parsed
            .as_ref()
            .ok()
            .and_then(|request| request.get("op"))
            .and_then(Value::as_str)
            == Some("close");
        let result = parsed.map_err(|e| e.to_string()).and_then(|request| {
            if text(&request, "op")? == "open" {
                if book.is_some() {
                    return Err("this process already owns a workbook".into());
                }
                let (opened, response) = OpenBook::open(&request)?;
                book = Some(opened);
                Ok(response)
            } else {
                book.as_mut()
                    .ok_or("no workbook is open")?
                    .operation(&request)
            }
        });
        let response = match result {
            Ok(outcome) => json!({"id":id,"ok":true,"outcome":outcome}),
            Err(error) => json!({"id":id,"ok":false,"error":error}),
        };
        serde_json::to_writer(&mut output, &response).map_err(|e| e.to_string())?;
        output
            .write_all(b"\n")
            .and_then(|_| output.flush())
            .map_err(|e| e.to_string())?;
        if close {
            break;
        }
    }
    Ok(())
}
