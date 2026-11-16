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

fn execution_environment(request: &Value) -> Result<Value, String> {
    let mut options = json!({});
    if let Some(language) = request.get("errorLanguage").filter(|v| !v.is_null()) {
        let name = language.as_str().ok_or("errorLanguage must be a string")?;
        if !matches!(name.to_ascii_lowercase().as_str(), "en" | "en-us" | "en-gb" | "ja" | "ja-jp") {
            return Err("unsupported VBA errorLanguage".into());
        }
        options["errorLanguage"] = language.clone();
    }
    if let Some(help) = request.get("vbaHelpFile").filter(|v| !v.is_null()) {
        if !help.is_string() { return Err("vbaHelpFile must be a string".into()); }
        options["vbaHelpFile"] = help.clone();
    }
    if let Some(page) = request.get("hostAnsiCodePage").filter(|v| !v.is_null()) {
        if !matches!(page.as_u64(), Some(932 | 1252)) {
            return Err("hostAnsiCodePage must be the integer 932 or 1252".into());
        }
        options["hostAnsiCodePage"] = page.clone();
    }
    Ok(options)
}

fn capability_name(value: oxivba_core::Capability) -> &'static str {
    use oxivba_core::Capability;
    match value {
        Capability::RunsOnOpen => "auto_run",
        Capability::WritesItsOwnCode => "writes_code",
        Capability::StartsAProgram => "starts_program",
        Capability::ReachesTheNetwork => "network",
        Capability::CallsIntoADll => "dll",
        Capability::TouchesTheRegistry => "registry",
        Capability::TouchesFiles => "files",
        Capability::ReachesOutsideExcel => "outside_excel",
        Capability::HidesWhatItReaches => "dynamic_target",
        Capability::QuietensExcel => "quiets_excel",
    }
}

fn parse_error_message(message: &str, source: &str) -> String {
    if source.is_ascii() {
        message.into()
    } else {
        format!("{message}; non-ASCII VBA source may use a different code page; reopen with source_code_page=... if the recorded encoding is incorrect")
    }
}

fn procedure_metadata(procedure: &oxivba_core::Procedure) -> Value {
    use oxivba_core::ast::{ParamMode, ProcKind, Visibility};
    let kind = match procedure.kind {
        ProcKind::Sub => "Sub",
        ProcKind::Function => "Function",
        ProcKind::PropertyGet => "Property Get",
        ProcKind::PropertyLet => "Property Let",
        ProcKind::PropertySet => "Property Set",
    };
    let visibility = match procedure.visibility {
        Visibility::Private => "Private",
        Visibility::Friend => "Friend",
        Visibility::Default | Visibility::Public | Visibility::Global => "Public",
    };
    let returns_value = matches!(procedure.kind, ProcKind::Function | ProcKind::PropertyGet);
    let return_type = returns_value.then(|| {
        procedure
            .return_type
            .as_ref()
            .map_or("Variant", |kind| kind.name.as_str())
    });
    let parameters: Vec<_> = procedure
        .params
        .iter()
        .map(|parameter| {
            let mode = match parameter.mode {
                ParamMode::ByRef => "ByRef",
                ParamMode::ByVal => "ByVal",
                ParamMode::ParamArray => "ParamArray",
            };
            json!({"name":parameter.name,"mode":mode,"type":parameter.type_name.name,
            "optional":parameter.optional || parameter.mode == ParamMode::ParamArray,
            "is_array":parameter.is_array || parameter.type_name.is_array,
            "has_default":parameter.default.is_some()})
        })
        .collect();
    let required = procedure
        .params
        .iter()
        .filter(|parameter| !parameter.optional && parameter.mode != ParamMode::ParamArray)
        .count();
    json!({"name":procedure.name,"line":procedure.span.line,"kind":kind,
        "visibility":visibility,"parameters":parameters,"required_arguments":required,
        "return_type":return_type,
        "returns_array":procedure.return_type.as_ref().is_some_and(|kind|kind.is_array)})
}

/// Read source only. This does not initialize a runtime or certify compatibility.
fn inspect_source(source: &str) -> Value {
    let mut modules = Vec::new();
    let mut signals = Vec::new();
    let mut auto_run = Vec::new();
    let mut unresolved = Vec::new();
    let mut unparsed = 0;
    let mut complete = true;
    for (index, block) in oxivba_core::source_modules(source).into_iter().enumerate() {
        let name = block
            .named
            .unwrap_or_else(|| format!("Module{}", index + 1));
        let parsed = match oxivba_core::parse_module(&block.text) {
            Ok(module) => module,
            Err(error) => {
                let line = match &error {
                    oxivba_core::LexError::UnterminatedString { line }
                    | oxivba_core::LexError::UnterminatedDate { line }
                    | oxivba_core::LexError::UnterminatedBracketExpr { line }
                    | oxivba_core::LexError::UnexpectedChar { line, .. } => *line,
                };
                complete = false;
                modules.push(
                    json!({"name":name,"is_class":block.class,"source_line":block.line,
                    "procedures":[],"api_names":{},"external_declares":[],
                    "unparsed_lines":0,"syntax_error_lines":[],
                    "parse_error":{"message":parse_error_message(&error.to_string(), &block.text),
                        "line":line,"source_line":block.line + line as usize - 1}}),
                );
                continue;
            }
        };
        let analysis = oxivba_core::analyse(&parsed);
        let safety = oxivba_core::assess(&parsed, &analysis);
        complete &= analysis.metrics.unparsed == 0 && parsed.syntax_errors.is_empty();
        unparsed += analysis.metrics.unparsed;
        for signal in safety.signals {
            signals.push(
                json!({"module":name,"capability":capability_name(signal.capability),
                "what":signal.what,"reason":signal.reason,"line":signal.line,
                "source_line":(signal.line>0).then(||block.line + signal.line as usize - 1)}),
            );
        }
        auto_run.extend(
            safety
                .runs_without_asking
                .into_iter()
                .map(|entry| format!("{name}.{entry}")),
        );
        unresolved.extend(safety.unresolved_late_binding.into_iter().map(
            |line| json!({"module":name,"line":line,"source_line":block.line + line as usize - 1}),
        ));
        modules.push(json!({"name":name,"is_class":block.class,"source_line":block.line,
            "unparsed_lines":analysis.metrics.unparsed,
            "syntax_error_lines":parsed.syntax_errors.iter().map(|span|span.line).collect::<Vec<_>>(),
            "procedures":parsed.items.iter().filter_map(|item|match item {
                oxivba_core::ModuleItem::Procedure(procedure)=>Some(procedure_metadata(procedure)),
                _=>None,
            }).collect::<Vec<_>>(),"parse_error":null,
            "api_names":analysis.api_names,"external_declares":analysis.external_declares}));
    }
    json!({"has_macros":!source.is_empty(),"modules":modules,"signals":signals,"auto_run":auto_run,
        "unresolved_calls":unresolved,"unparsed_lines":unparsed,"analysis_complete":complete,
        "compatibility_verified":false,"execution_support":"not_assessed"})
}

fn package_active_sheet(bytes: &[u8], count: usize) -> Result<usize, String> {
    use quick_xml::events::Event;
    let mut archive=oxidocs_common::archive::OoxmlArchive::new(bytes).map_err(|e|e.to_string())?;
    let xml=archive.read_part("xl/workbook.xml").map_err(|e|e.to_string())?;
    let mut reader=quick_xml::NsReader::from_str(&xml);
    let mut path=Vec::<(Vec<u8>,bool)>::new();
    let mut views=Vec::new();
    loop {
        let event=reader.read_event().map_err(|e|e.to_string())?;
        match &event {
            Event::Start(element)|Event::Empty(element)=>{
                let standard=matches!(reader.resolve_element(element.name()).0,
                    quick_xml::name::ResolveResult::Bound(ns)if matches!(ns.as_ref(),
                        b"http://schemas.openxmlformats.org/spreadsheetml/2006/main"|
                        b"http://purl.oclc.org/ooxml/spreadsheetml/main"));
                let local=element.local_name().as_ref().to_vec();
                if standard && path.iter().all(|(_,valid)|*valid) && path.len()==2
                    && path[0].0==b"workbook" && path[1].0==b"bookViews" && local==b"workbookView" {
                    let at=oxidocs_common::xml_utils::get_attr(element,"activeTab")
                        .map(|v|v.parse::<usize>()).transpose().map_err(|e|e.to_string())?.unwrap_or(0);
                    if at>=count{return Err("active sheet index is out of bounds".into());}
                    views.push(at);
                }
                if matches!(event,Event::Start(_)){path.push((local,standard));}
            }
            Event::End(_)=>{path.pop();}
            Event::Eof=>break,
            _=>{}
        }
    }
    Ok(views.get(views.len().saturating_sub(2)).copied().unwrap_or(0))
}

impl OpenBook {
    fn open(request: &Value) -> Result<(Self, Value), String> {
        fields(
            request,
            &["id", "op", "path", "sourceCodePage", "seed", "currentTime",
              "errorLanguage", "vbaHelpFile", "hostAnsiCodePage"],
        )?;
        clock(request)?;
        let environment = execution_environment(request)?;
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
        let mut continued = environment.clone();
        for (name, value) in environment.as_object().expect("environment object") {
            metadata[name] = value.clone();
        }
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
        if operation == "inspect" {
            fields(request, &["id", "op", "currentTime"])?;
            return Ok(inspect_source(&self.source));
        }
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
        let outcome = self.session.run_native(book, call).map_err(|error| {
            if !self.source.is_ascii() && oxivba_core::parse_project(&self.source).is_err() {
                parse_error_message(&error, &self.source)
            } else {
                error
            }
        })?;
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

#[cfg(test)]
mod tests {
    #[test]
    fn cold_open_chooses_saved_active_window_sheet() {
        let fixture:Value=serde_json::from_str(include_str!("../../../tests/fixtures/vba/saved_window_views_native.json")).unwrap();
        for case in fixture["atomic_vba"]["cases"].as_array().unwrap() {
            let hex=case["carrier_bytes_hex"].as_str().unwrap();
            let bytes:Vec<u8>=(0..hex.len()).step_by(2).map(|at|u8::from_str_radix(&hex[at..at+2],16).unwrap()).collect();
            let book=oxicells_core::parse_xlsx(&bytes).unwrap();
            let active=super::package_active_sheet(&bytes,book.sheets.len()).unwrap();
            assert_eq!(book.sheets[active].name,case["Native_runs"][3]["values"][5].as_str().unwrap(),"saved windows {}",case["windows"]);
        }
    }

    use super::*;

    #[test]
    fn inspection_keeps_unread_modules_visible_and_preserves_other_findings() {
        let source = concat!(
            "Attribute VB_Name = \"Clean\"\nPublic Sub Auto_Open()\n Call Shell(\"never.exe\", 1)\nEnd Sub\n",
            "Attribute VB_Name = \"BadLex\"\nPublic Sub Broken()\n Dim name As String\n name = \"unterminated\nEnd Sub\n",
            "Attribute VB_Name = \"BadSyntax\"\nPublic Sub Broken()\n Else\nEnd Sub\n"
        );
        let report = inspect_source(source);
        assert_eq!(report["modules"].as_array().unwrap().len(), 3);
        assert_eq!(report["analysis_complete"], false);
        assert_eq!(report["compatibility_verified"], false);
        assert_eq!(report["modules"][1]["parse_error"]["line"], 4);
        assert_eq!(report["modules"][1]["procedures"], json!([]));
        assert_eq!(report["modules"][1]["api_names"], json!({}));
        assert_eq!(report["modules"][1]["external_declares"], json!([]));
        assert!(report["modules"][0]["parse_error"].is_null());
        assert!(report["unparsed_lines"].as_u64().unwrap() > 0);
        assert_eq!(report["auto_run"], json!(["Clean.Auto_Open"]));
        let shell = report["signals"]
            .as_array()
            .unwrap()
            .iter()
            .find(|signal| signal["capability"] == "starts_program" && signal["module"] == "Clean")
            .unwrap();
        // API names are aggregated per module by the existing analyzer.
        // Preserve its unknown position rather than inventing a location.
        assert_eq!(shell["line"], 0);
        assert!(shell["source_line"].is_null());
    }

    #[test]
    fn inspection_does_not_treat_comments_strings_or_inactive_branches_as_calls() {
        let source = concat!(
            "VERSION 1.0 CLASS\nAttribute VB_Name = \"Box\"\n",
            "' Shell(\"comment.exe\")\n#If Mac Then\n",
            "Public Sub Auto_Open()\n Shell \"dead.exe\"\nEnd Sub\n#End If\n",
            "Public Function Text() As String\n Text = \"Shell CreateObject\"\nEnd Function\n"
        );
        let report = inspect_source(source);
        assert_eq!(report["analysis_complete"], true);
        assert_eq!(report["signals"], json!([]));
        assert_eq!(report["auto_run"], json!([]));
        assert_eq!(report["unparsed_lines"], 0);
    }

    #[test]
    fn inspection_reports_procedure_signatures_and_effective_defaults() {
        let report = inspect_source(concat!(
            "Attribute VB_Name = \"Metadata\"\nSub Defaults(item)\nEnd Sub\n",
            "Private Function Compute(ByVal count As Long, Optional ByRef label As String = \"x\") As Long\nEnd Function\n",
            "Public Sub Collect(ParamArray items() As Variant)\nEnd Sub\n",
            "Public Function Values() As Long()\nEnd Function\n",
            "VERSION 1.0 CLASS\nAttribute VB_Name = \"Box\"\n",
            "Friend Property Get Title() As String\nEnd Property\n",
            "Public Property Let Title(ByVal value As String)\nEnd Property\n",
            "Private Property Set Item(ByVal value As Object)\nEnd Property\n"
        ));
        assert_eq!(report["analysis_complete"], true, "{report}");
        let procedures = &report["modules"][0]["procedures"];
        assert_eq!(procedures[0]["kind"], "Sub");
        assert_eq!(procedures[0]["visibility"], "Public");
        assert_eq!(procedures[0]["required_arguments"], 1);
        assert_eq!(
            procedures[0]["parameters"][0],
            json!({"name":"item","mode":"ByRef",
            "type":"Variant","optional":false,"is_array":false,"has_default":false})
        );
        assert!(procedures[0]["return_type"].is_null());
        assert_eq!(procedures[1]["kind"], "Function");
        assert_eq!(procedures[1]["visibility"], "Private");
        assert_eq!(procedures[1]["return_type"], "Long");
        assert_eq!(procedures[1]["required_arguments"], 1);
        assert_eq!(procedures[1]["parameters"][0]["mode"], "ByVal");
        assert_eq!(
            procedures[1]["parameters"][1],
            json!({"name":"label","mode":"ByRef",
            "type":"String","optional":true,"is_array":false,"has_default":true})
        );
        assert_eq!(procedures[2]["required_arguments"], 0);
        assert_eq!(procedures[2]["parameters"][0]["mode"], "ParamArray");
        assert_eq!(procedures[2]["parameters"][0]["optional"], true);
        assert_eq!(procedures[2]["parameters"][0]["is_array"], true);
        assert_eq!(procedures[3]["return_type"], "Long");
        assert_eq!(procedures[3]["returns_array"], true);
        let properties = &report["modules"][1]["procedures"];
        assert_eq!(properties[0]["kind"], "Property Get");
        assert_eq!(properties[0]["visibility"], "Friend");
        assert_eq!(properties[1]["kind"], "Property Let");
        assert_eq!(properties[2]["kind"], "Property Set");
        assert_eq!(properties[2]["visibility"], "Private");
        assert!(properties[2]["return_type"].is_null());
    }

    #[test]
    fn non_ascii_parse_failures_keep_the_error_and_suggest_the_code_page_option() {
        for (source, hint) in [
            ("Sub A()\n ' 日本語\n }\nEnd Sub", true),
            ("Sub A()\n }\nEnd Sub", false),
        ] {
            let report = inspect_source(source);
            let message = report["modules"][0]["parse_error"]["message"]
                .as_str()
                .unwrap();
            assert!(message.contains("unexpected character"), "{message}");
            assert_eq!(message.contains("source_code_page="), hint, "{message}");
            assert_eq!(report["modules"][0]["procedures"], json!([]));
            assert_eq!(report["analysis_complete"], false);
        }
    }
}
