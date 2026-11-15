// SPDX-License-Identifier: MIT OR Apache-2.0

use super::*;

/// A frame's identity survives changes to its collection position and number.
/// The loaded frame uses WorkbookHost's working fields; other frames own them
/// here. Keeping closed frames also keeps already issued references distinct
/// from subsequently created windows.
pub(super) struct WindowFrames {
    pub(super) active: usize,
    loaded: usize,
    frames: Vec<WindowFrame>,
}

struct WindowFrame {
    open: bool,
    state: i64,
    view: Option<WindowView>,
}

#[derive(Clone)]
struct WindowView {
    active_sheet: usize,
    selection: CellRange,
    active_cell: CellAddress,
    cell_selection_areas: std::collections::HashMap<usize, Vec<CellRange>>,
    selections: std::collections::HashMap<usize, (CellRange, CellAddress)>,
    views: std::collections::HashMap<usize, View>,
    selected_group: Option<usize>,
    shape_selection: Vec<u64>,
    caption: Option<Value>,
    panes: Vec<(u32, u32)>,
}

impl Default for WindowFrames {
    fn default() -> Self {
        Self {
            active: 0,
            loaded: 0,
            frames: vec![WindowFrame { open: true, state: -4137, view: None }],
        }
    }
}

impl WindowFrames {
    pub(super) fn open_ids(&self) -> Vec<usize> {
        self.frames.iter().enumerate().rev()
            .filter_map(|(id, frame)| frame.open.then_some(id)).collect()
    }

    fn number(&self, id: usize) -> Option<usize> {
        self.frames.get(id)?;
        Some(self.frames[..=id].iter().filter(|frame| frame.open).count())
    }
}

impl WindowView {
    fn take_from(host: &WorkbookHost<'_>) -> Self {
        Self {
            active_sheet: host.active_sheet,
            selection: host.selection,
            active_cell: host.active_cell,
            cell_selection_areas: host.cell_selection_areas.clone(),
            selections: host.selections.clone(),
            views: host.views.clone(),
            selected_group: host.selected_group,
            shape_selection: host.shape_selection.clone(),
            caption: host.window_caption.clone(),
            panes: host.workbook.sheets.iter().map(|sheet| (sheet.frozen_rows, sheet.frozen_cols)).collect(),
        }
    }

    fn install(self, host: &mut WorkbookHost<'_>) {
        host.active_sheet = self.active_sheet;
        host.selection = self.selection;
        host.active_cell = self.active_cell;
        host.cell_selection_areas = self.cell_selection_areas;
        host.selections = self.selections;
        host.views = self.views;
        host.selected_group = self.selected_group;
        host.shape_selection = self.shape_selection;
        host.window_caption = self.caption;
        for (sheet, (rows, columns)) in host.workbook.sheets.iter_mut().zip(self.panes) {
            sheet.frozen_rows = rows;
            sheet.frozen_cols = columns;
        }
    }

    fn renumber(&mut self, moved: &dyn Fn(usize) -> Option<usize>, sheet_count: usize) {
        self.active_sheet = moved(self.active_sheet).unwrap_or(0);
        let home = CellRange::single(CellAddress { sheet: self.active_sheet, row: 1, column: 0 });
        self.selection = moved(self.selection.sheet).map(|sheet| CellRange { sheet, ..self.selection }).unwrap_or(home);
        self.active_cell = moved(self.active_cell.sheet).map(|sheet| CellAddress { sheet, ..self.active_cell }).unwrap_or(home.first());
        self.cell_selection_areas = std::mem::take(&mut self.cell_selection_areas).into_iter()
            .filter_map(|(sheet, areas)| moved(sheet).map(|sheet| (sheet, areas.into_iter().map(|range| CellRange { sheet, ..range }).collect())))
            .collect();
        self.selections = std::mem::take(&mut self.selections).into_iter()
            .filter_map(|(sheet, (range, cell))| moved(sheet).map(|sheet| (sheet, (CellRange { sheet, ..range }, CellAddress { sheet, ..cell })))).collect();
        self.views = std::mem::take(&mut self.views).into_iter()
            .filter_map(|(sheet, view)| moved(sheet).map(|sheet| (sheet, view))).collect();
        let mut panes = vec![(0, 0); sheet_count];
        for (sheet, split) in std::mem::take(&mut self.panes).into_iter().enumerate() {
            if let Some(target) = moved(sheet).and_then(|sheet| panes.get_mut(sheet)) { *target = split; }
        }
        self.panes = panes;
    }
}

impl WorkbookHost<'_> {
    pub(super) fn window_id(&self, object: &ObjectRef) -> Option<usize> {
        match self.objects.get(host_object_index(object.handle)) {
            Some(HostObject::Window(id)) => Some(*id),
            _ => None,
        }
    }

    pub(super) fn windows_scope(&self, object: &ObjectRef) -> Option<bool> {
        match self.objects.get(host_object_index(object.handle)) {
            Some(HostObject::Windows(application)) => Some(*application),
            _ => None,
        }
    }

    pub(super) fn window_is_open(&self, id: usize) -> bool {
        self.window_frames.frames.get(id).is_some_and(|frame| frame.open)
    }

    /// Loading a held frame for a property operation does not activate it.
    pub(super) fn load_window(&mut self, id: usize) -> usize {
        let previous = self.window_frames.loaded;
        if previous != id {
            let view = self.window_frames.frames[id].view.take().expect("parked frame has a view");
            self.window_frames.frames[previous].view = Some(WindowView::take_from(self));
            view.install(self);
            self.window_frames.loaded = id;
        }
        previous
    }

    pub(super) fn renumber_window_views(&mut self, moved: &dyn Fn(usize) -> Option<usize>) {
        for frame in &mut self.window_frames.frames {
            if let Some(view) = &mut frame.view { view.renumber(moved, self.workbook.sheets.len()); }
        }
    }

    fn activate_window_frame(&mut self, id: usize) {
        self.load_window(id);
        let maximized = self.window_frames.frames.iter().any(|frame| frame.open && frame.state == -4137);
        if maximized {
            for frame in &mut self.window_frames.frames {
                if frame.open && frame.state == -4137 { frame.state = -4143; }
            }
            if self.window_frames.frames[id].state != -4140 {
                self.window_frames.frames[id].state = -4137;
            }
        }
        self.window_frames.active = id;
    }

    fn new_window(&mut self, source: Option<usize>) -> Value {
        if let Some(id) = source { self.load_window(id); }
        let sheet = self.active_sheet;
        if source.is_none() {
            let primary = *self.window_frames.open_ids().last().expect("open workbook has a frame");
            self.load_window(primary);
            if self.active_sheet != sheet {
                self.activate_sheet(sheet);
            }
        }
        let mut view = WindowView::take_from(self);
        let home = CellRange::single(CellAddress { sheet, row: 1, column: 0 });
        view.selection = home;
        view.active_cell = home.first();
        view.cell_selection_areas.clear();
        view.selections.clear();
        view.views.clear();
        view.selected_group = None;
        view.shape_selection.clear();
        view.panes.fill((0, 0));
        view.caption = None;
        let id = self.window_frames.frames.len();
        for frame in &mut self.window_frames.frames {
            if frame.open && frame.state == -4137 { frame.state = -4143; }
        }
        self.window_frames.frames.push(WindowFrame { open: true, state: -4137, view: Some(view) });
        self.load_window(id);
        self.window_frames.active = id;
        self.object(HostObject::Window(id))
    }

    fn close_window_frame(&mut self, id: usize) {
        self.window_frames.frames[id].open = false;
        if self.window_frames.active == id {
            if let Some(next) = self.window_frames.open_ids().first().copied() {
                self.load_window(next);
                self.window_frames.active = next;
            }
        }
    }

    pub(super) fn validate_window_state_argument(&self, value: &Value) -> Result<(), String> {
        let mut pending = vec![(value, false)];
        let mut object_operand = false;
        while let Some((operand, array_element)) = pending.pop() {
            match operand {
                Value::Array(array) => {
                    let Some(first) = array.values.first() else {
                        return Err(host_error(13, "WindowState cannot convert an empty array"));
                    };
                    let nested = matches!(first, Value::Array(_));
                    if array.values.iter().any(|item| matches!(item, Value::Array(_)) != nested) {
                        return Err(host_error(13, "WindowState requires a rectangular array"));
                    }
                    if let Value::Array(first) = first {
                        if array.values.iter().any(|item| match item {
                            Value::Array(item) => !first.dimensions.iter().map(|dimension| dimension.length)
                                .eq(item.dimensions.iter().map(|dimension| dimension.length)),
                            _ => true,
                        }) {
                            return Err(host_error(13, "WindowState requires matching array dimensions"));
                        }
                    }
                    pending.extend(array.values.iter().rev().map(|item| (item, true)));
                }
                Value::Null if array_element => return Err(host_error(13, "WindowState cannot convert an array Null")),
                Value::Nothing => return Err(host_error(5, "WindowState requires an Excel scalar")),
                Value::Date(serial) if *serial < 0.0 => return Err(host_error(5, "WindowState cannot convert a negative date")),
                Value::Error(code) if spreadsheet_error_text_exact(*code).is_none() =>
                    return Err(host_error(13, "WindowState cannot convert this error value")),
                Value::Object(_) | Value::Record(_) => object_operand = true,
                _ => {}
            }
        }
        if object_operand { return Err(self.scalar_set_error("Window", "WindowState")); }
        let scalar = Self::protection_scalar(value)?;
        let number = if let Value::Currency(units) = scalar {
            let magnitude = units.unsigned_abs();
            let cents = magnitude / 100 + u64::from(magnitude % 100 >= 50);
            Some((cents as f64 / 100.0) * if *units < 0 { -1.0 } else { 1.0 })
        } else if scalar.is_string() {
            scalar.padded_string_units().and_then(|units| {
                let end = units.iter().position(|unit| *unit == 0).unwrap_or(units.len());
                String::from_utf16(&units[..end]).ok()
            }).and_then(|text| text.trim_matches(' ').parse::<f64>().ok())
        } else {
            Self::legacy_assignment_number(value)
                .map_err(|number| if number == 1004 {
                    self.scalar_set_error("Window", "WindowState")
                } else { host_error(number, "WindowState cannot convert this operand") })?
        };
        let number = number.filter(|number| number.is_finite())
            .ok_or_else(|| self.scalar_set_error("Window", "WindowState"))?;
        let enum_operand = matches!(value, Value::Int16(_) | Value::Integer(_) | Value::LongLong(_)
            | Value::Byte(_) | Value::Single(_) | Value::Double(_) | Value::Decimal(_));
        let internal = if enum_operand {
            match number { -4140.0 => 1.0, -4137.0 => 2.0, -4143.0 => 3.0, _ => number }
        } else { number };
        // Compare the endpoints directly: adding the margin before floor
        // would collapse adjacent Double operands at the upper boundary.
        let margin = 2049.0 / 8589934592.0;
        if ((1.0 - margin)..(4.0 - margin)).contains(&internal) {
            Ok(())
        } else {
            Err(self.scalar_set_error("Window", "WindowState"))
        }
    }

    pub(super) fn frame_caption(&self) -> Value {
        self.window_caption.clone().unwrap_or_else(|| {
            let mut caption = self.file_name.clone().unwrap_or_else(|| "Book1".into());
            if self.window_frames.open_ids().len() > 1 {
                if let Some(number) = self.window_frames.number(self.window_frames.loaded) {
                    caption.push(':');
                    caption.push_str(&number.to_string());
                }
            }
            Value::String(caption)
        })
    }
}

impl WorkbookApplication<'_> {
    pub(super) fn window_application_visible(&self) -> bool {
        self.caption_settings().and_then(|settings| settings.get("visible")) != Some(&Value::Boolean(false))
    }

    pub(super) fn window_collection(&self, owner: Option<u64>) -> Vec<(u64, usize)> {
        if !self.window_enumeration_available && !self.window_application_visible() { return Vec::new(); }
        self.window_z_order.iter().rev().copied().filter(|(held, id)| {
            owner.is_none_or(|owner| owner == *held) && !self.closed.contains(held)
                && self.hosts.get(held).is_some_and(|host| host.window_is_open(*id))
        }).collect()
    }

    fn window_item(&mut self, owner: Option<u64>, index: &Value) -> Result<Value, String> {
        let windows = self.window_collection(owner);
        let target = if let Some(name) = index.string_units() {
            windows.into_iter().find(|(owner, id)| {
                let host = self.hosts.get_mut(owner).expect("window owner");
                let previous = host.load_window(*id);
                let caption = host.frame_caption();
                host.load_window(previous);
                caption.string_units().is_some_and(|caption| crate::vba_automation::workbook_name_matches(
                    &String::from_utf16_lossy(&caption), &name))
            })
        } else {
            let number = match index {
                Value::Empty => Some(0.0),
                Value::Boolean(value) => Some(if *value { -1.0 } else { 0.0 }),
                Value::Decimal(number) => Some(number.to_f64()),
                value => any_number(value),
            }.filter(|number| number.is_finite()).ok_or_else(|| host_error(13, "Window index must be a caption or number"))?
                .round_ties_even();
            if !(f64::from(i32::MIN)..=f64::from(i32::MAX)).contains(&number) {
                return Err(host_error(6, "Window index overflows Long"));
            }
            usize::try_from(number as i64).ok().and_then(|index| index.checked_sub(1))
                .and_then(|index| windows.get(index).copied())
        };
        let (owner, id) = target.ok_or_else(|| host_error(9, "Window index is outside the collection"))?;
        Ok(self.hosts.get_mut(&owner).expect("window owner").object(HostObject::Window(id)))
    }
}

impl ApplicationHost<'_> {
    pub(super) fn windows_call(&mut self, receiver: Option<&ObjectRef>, name: &str,
        args: &[Value], names: &[Option<String>], files: Option<&mut oxivba_core::files::Files>)
        -> Option<Result<Option<Value>, String>> {
        let mut application = self.application.borrow_mut();
        if let Some(error) = receiver.and_then(|object| application.closed_object_error(object)) {
            return Some(Err(error));
        }
        let native_owner = receiver.and_then(|object| application.native_owner(object));
        let scope = native_owner.and_then(|owner| application.hosts[&owner].windows_scope(receiver?));
        let book = native_owner.filter(|owner| receiver.is_some_and(|object| application.hosts[owner].is_workbook(object)));
        let global = receiver.is_none_or(|object| application.is_application(object));
        let window = native_owner.and_then(|owner| application.hosts[&owner].window_id(receiver?).map(|id| (owner, id)));
        if let Some((owner, id)) = window {
            if !application.hosts[&owner].window_is_open(id) {
                return Some(Err(host_error(424, "The Window object is no longer available")));
            }
        }
        if global && name.eq_ignore_ascii_case("Visible") && args.is_empty() {
            return Some(Ok(Some(Value::Boolean(application.window_application_visible()))));
        }
        let property = name.eq_ignore_ascii_case("Windows") && (global || book.is_some());
        if property || scope.is_some() {
            let owner = if property { book } else { scope.filter(|all| !all).and(native_owner) };
            if property && args.is_empty() {
                let owner = book.unwrap_or(self.code_owner);
                return Some(Ok(Some(application.hosts.get_mut(&owner).expect("Windows caller").object(HostObject::Windows(global)))));
            }
            if scope.is_some() && name.eq_ignore_ascii_case("Count") && args.is_empty() {
                return Some(Ok(Some(Value::Integer(application.window_collection(owner).len() as i64))));
            }
            if property || name.is_empty() || name.eq_ignore_ascii_case("Item") || name.eq_ignore_ascii_case("_Default") {
                let args = if names.is_empty() { Ok(args.to_vec()) } else {
                    normalize_named_arguments(args, names, &["Index"], "Windows.Item")
                };
                return Some(args.and_then(|args| match args.as_slice() {
                    [index] => application.window_item(owner, index).map(Some),
                    _ => Err(host_error(450, "Windows.Item requires one index")),
                }));
            }
        }
        if global && name.eq_ignore_ascii_case("ActiveWindow") && args.is_empty() {
            let owner = application.active;
            if application.closed.contains(&owner) { return Some(Ok(Some(Value::Nothing))); }
            let host = application.hosts.get_mut(&owner).expect("active window owner");
            return Some(Ok(Some(host.object(HostObject::Window(host.window_frames.active)))));
        }
        let new_window = book.map(|owner| (owner, None))
            .or_else(|| window.map(|(owner, id)| (owner, Some(id))));
        if let Some((owner, source)) = new_window.filter(|_| name.eq_ignore_ascii_case("NewWindow")) {
            if !args.is_empty() {
                return Some(Err(host_error(450, if source.is_some() {
                    "Window.NewWindow takes no arguments"
                } else { "Workbook.NewWindow takes no arguments" })));
            }
            application.ensure_application_owner(owner);
            let result = application.hosts.get_mut(&owner).expect("NewWindow owner").new_window(source);
            let id = application.hosts[&owner].window_frames.active;
            application.window_z_order.push((owner, id));
            application.window_enumeration_available = true;
            application.activate(owner);
            return Some(Ok(Some(result)));
        }
        if let Some(owner) = book.filter(|_| name.eq_ignore_ascii_case("Activate")) {
            if !args.is_empty() { return Some(Err(host_error(450, "Workbook.Activate takes no arguments"))); }
            application.activate(owner);
            if application.active == owner {
                let host = application.hosts.get_mut(&owner).expect("Workbook.Activate owner");
                if let Some(primary) = host.window_frames.open_ids().last().copied() {
                    host.activate_window_frame(primary);
                }
                application.activate(owner);
            }
            return Some(Ok(Some(Value::Empty)));
        }
        if let Some((owner, id)) = window {
            if name.eq_ignore_ascii_case("WindowState") && args.is_empty() {
                return Some(Ok(Some(Value::Integer(application.hosts[&owner].window_frames.frames[id].state))));
            }
            if name.eq_ignore_ascii_case("Index") && args.is_empty() {
                let position = application.window_collection(None).iter().position(|held| *held == (owner, id));
                return Some(position.map(|position| Some(Value::Integer((position + 1) as i64)))
                    .ok_or_else(|| {
                        let description = match application.hosts[&owner].error_language {
                            ErrorLanguage::Japanese => "Window クラスの Index プロパティを取得できません。",
                            ErrorLanguage::English => "Unable to get the Index property of the Window class",
                        };
                        oxivba_core::host_error_from(1004, "Microsoft Excel", description)
                    }));
            }
            if name.eq_ignore_ascii_case("WindowNumber") && args.is_empty() {
                let number = application.hosts[&owner].window_frames.number(id).unwrap_or(0);
                return Some(Ok(Some(Value::Double(number as f64))));
            }
            if name.eq_ignore_ascii_case("Activate") {
                if !args.is_empty() { return Some(Err(host_error(450, "Window.Activate takes no arguments"))); }
                application.activate(owner);
                if application.active == owner {
                    let host = application.hosts.get_mut(&owner).expect("Window owner");
                    host.activate_window_frame(id);
                    application.activate(owner);
                }
                return Some(Ok(Some(Value::Empty)));
            }
            if name.eq_ignore_ascii_case("Close") {
                let args = if names.is_empty() { Ok(args.to_vec()) } else {
                    normalize_named_arguments(args, names, &["SaveChanges", "Filename", "RouteWorkbook"], "Window.Close")
                };
                return Some(args.and_then(|args| {
                    let hidden = !application.window_application_visible();
                    let active_owner = application.active == owner;
                    let host = application.hosts.get_mut(&owner).expect("Window owner");
                    if !host.window_frames.frames[id].open { return Err(host_error(1004, "Window is closed")); }
                    if host.window_frames.open_ids().len() == 1 { return application.close_workbook(owner, &args, files).map(Some); }
                    let closing_active = host.window_frames.active == id;
                    host.close_window_frame(id);
                    if hidden && active_owner && closing_active { application.window_enumeration_available = false; }
                    Ok(Some(Value::Empty))
                }));
            }
        }
        None
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn native_group(index: usize) {
        native_group_fields(index, None);
    }

    fn native_group_fields(index: usize, compared_fields: Option<usize>) {
        let fixture: serde_json::Value = serde_json::from_str(include_str!(
            "../../../tests/fixtures/vba/window_collections_native.json")).unwrap();
        let group = &fixture["groups"][index];
        let hex = group["carrier_bytes_hex"].as_str().unwrap();
        let carrier: Vec<u8> = (0..hex.len()).step_by(2)
            .map(|at| u8::from_str_radix(&hex[at..at + 2], 16).unwrap()).collect();
        let file_name = group["file_name"].as_str().unwrap();
        let startup: Vec<u16> = serde_json::from_value(group["startup_profile_utf16"].clone()).unwrap();
        let fields = group["fields"].as_array().unwrap();
        let mut failures = Vec::new();
        for case in group["cases"].as_array().unwrap() {
            let mode = case["mode"].as_i64().unwrap();
            let binding = case["binding"].as_i64().unwrap();
            let context = VbaRunContext {
                environment: ErrorEnvironment { language: ErrorLanguage::Japanese,
                    help_file: group["help_file"].as_str().unwrap().into() },
                host_ansi_code_page: None, project_name: "VBAProject".into(),
                file_name: Some(file_name.into()), random_seed: 1, current_time: 46000.0,
                application_state: Some(VbaApplicationState {
                    visible: Some(false),
                    enable_events: false, display_alerts: false,
                    standard_font_startup: Some(VbaStandardFontStartup { profile_utf16: Some(startup.clone()) }),
                    ..Default::default()
                }),
                source_edits: None, document_modules: None, file_system_state: None,
                file_context: Some(file_io::FileContext {
                    package_bytes: Some(carrier.clone()), full_name: Some(format!(r"C:\work\{file_name}")),
                    read_only: Some(true),
                    folders: vec![r"C:\work".into()], ..Default::default()
                }),
            };
            let mut args = vec![Value::Integer(mode), Value::Integer(binding)];
            if let Some(bits) = case["input_operand_bits"].as_str() {
                args.push(Value::Double(f64::from_bits(u64::from_str_radix(bits, 16).unwrap())));
            }
            let outcome = execute_workbook_action_in_session(
                oxicells_core::parse_xlsx_preserving_values(&carrier).unwrap(),
                group["program"].as_str().unwrap(), Some(group["procedure"].as_str().unwrap_or("Probe")),
                args, 0, context, &mut None).unwrap();
            if let Some(error) = outcome.error {
                failures.push(format!("mode={mode}, binding={binding}: {}", serde_json::to_value(error).unwrap()));
                continue;
            }
            if !group["mutates_caller_inputs"].as_bool().unwrap_or(false) {
                assert!(!outcome.workbook_changed, "{} mode={mode}, binding={binding}: caller unchanged", group["id"]);
            }
            assert!(outcome.files.is_empty() && outcome.deleted_files.is_empty(),
                "{} mode={mode}, binding={binding}: probe writes no files", group["id"]);
            let result = serde_json::to_value(outcome.result).unwrap();
            let actual = result["value"]["values"].as_array().expect("probe returns an array");
            let expected = case["expected"].as_array().unwrap();
            assert_eq!(actual.len(), fields.len());
            assert_eq!(expected.len(), fields.len());
            let compared_fields = if binding >= 8 { fields.len() }
                else { compared_fields.unwrap_or(fields.len()) };
            assert!(compared_fields <= fields.len());
            for (at, (actual, expected)) in actual.iter().zip(expected).enumerate()
                .filter(|(at, _)| *at < compared_fields || *at > 7) {
                if at == 8 {
                    if let Some(bits) = case["expected_operand_bits"].as_str() {
                        let actual = actual["value"].as_f64().expect("numeric operand").to_bits();
                        let expected = u64::from_str_radix(bits, 16).unwrap();
                        if actual != expected {
                            failures.push(format!("mode={mode}, binding={binding}, operand bits: {actual:016x} != {expected:016x}"));
                        }
                        continue;
                    }
                }
                let mut value = actual["value"].clone();
                if group["id"] == "property_types" && at < 39 && at % 3 == 0
                    && case["expected"][at + 1] == 10 && actual["type"] == "error" {
                    // COM exposes VT_ERROR as a signed SCODE. The public engine
                    // API exposes its VBA error code; preserve the complete
                    // native SCODE in the fixture and compare the exact mapping.
                    let code = value.as_u64().expect("VBA error code");
                    assert!(code <= u64::from(u16::MAX));
                    value = serde_json::Value::from(i64::from((0x800a_0000_u32 | code as u32) as i32));
                }
                if &value != expected {
                    failures.push(format!("mode={mode}, binding={binding}, {}: {} != {}",
                        fields[at], value, expected));
                }
            }
        }
        assert!(failures.is_empty(), "{} Native differences:\n{}", group["id"], failures.join("\n"));
    }

    #[test]
    fn window_collections_match_native_binding_and_coercion() {
        native_group(0);
    }

    #[test]
    fn window_lifecycle_matches_native_frame_observations() {
        native_group(1);
    }

    #[test]
    fn new_window_state_trace_matches_native() {
        native_group(2);
    }

    #[test]
    fn window_dispatch_identity_matches_native() {
        native_group(3);
    }

    #[test]
    fn window_close_reference_state_matches_native() {
        native_group(4);
    }

    #[test]
    fn window_property_types_and_errors_match_native() {
        native_group(5);
    }

    #[test]
    fn explicit_window_new_window_matches_native() {
        native_group(6);
    }

    #[test]
    fn window_state_assignment_errors_match_native() {
        native_group_fields(7, Some(7));
    }

    #[test]
    fn window_state_assignment_error_boundaries_match_native() {
        native_group_fields(8, Some(7));
    }

    #[test]
    fn window_state_assignment_contexts_match_native() {
        native_group_fields(9, Some(7));
    }

    #[test]
    fn window_state_conversion_controls_match_native() {
        native_group_fields(10, Some(7));
    }

    #[test]
    fn window_state_numeric_rounding_match_native() {
        native_group_fields(11, Some(7));
    }

    #[test]
    fn window_state_numeric_thresholds_match_native() {
        native_group_fields(12, Some(7));
    }

    #[test]
    fn window_state_numeric_boundaries_match_native() {
        native_group_fields(13, Some(7));
    }

    #[test]
    fn window_state_enum_boundaries_match_native() {
        native_group_fields(14, Some(7));
    }

    #[test]
    fn window_state_operand_precision_match_native() {
        native_group_fields(15, Some(7));
    }

    #[test]
    fn window_state_binary_boundaries_match_native() {
        native_group_fields(16, Some(7));
    }
}
