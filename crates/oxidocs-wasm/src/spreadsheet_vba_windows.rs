// SPDX-License-Identifier: MIT OR Apache-2.0

use super::*;

/// A frame change posts a notification; the VBA activation is a separate
/// queued action. Messages posted by a host turn are not consumed recursively.
#[derive(Clone, Copy)]
pub(super) enum WindowUiMessage {
    FrameChanged(u64, usize),
    ActivateForeground(u64, usize),
}

/// A frame's identity survives changes to its collection position and number.
/// The loaded frame uses WorkbookHost's working fields; other frames own them
/// here. Keeping closed frames also keeps already issued references distinct
/// from subsequently created windows.
pub(super) struct WindowFrames {
    pub(super) content_changed: bool,
    pub(super) active: usize,
    loaded: usize,
    frames: Vec<WindowFrame>,
}

struct WindowFrame {
    creation_caption: Option<String>,
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
            content_changed: false,
            active: 0,
            loaded: 0,
            frames: vec![WindowFrame { creation_caption: None, open: true, state: -4137, view: None }],
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
        self.window_frames.active = id;
    }

    pub(super) fn loaded_window_state(&self) -> i64 {
        let frame = &self.window_frames.frames[self.window_frames.loaded];
        if frame.state == -4137 && self.window_frames.loaded != self.window_frames.active {
            -4143
        } else {
            frame.state
        }
    }

    fn new_window(&mut self, source: Option<usize>, state: i64) -> Value {
        // Saved=True acknowledges the current flag; it does not erase content
        // changes that creating another view exposes again.
        if self.window_frames.content_changed { self.saved = false; }
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
        self.window_frames.frames.push(WindowFrame { creation_caption: None, open: true, state, view: Some(view) });
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

    fn requested_window_state(&self, value: &Value) -> Result<i64, String> {
        self.requested_display_state(value, || self.scalar_set_error("Window", "WindowState"))
    }

    fn requested_application_window_state(&self, value: &Value, typed: bool) -> Result<i64, String> {
        let state = self.requested_display_state(value, || {
            if typed {
                let description = match self.error_language {
                    ErrorLanguage::Japanese => "'WindowState' メソッドは失敗しました: '_Application' オブジェクト",
                    ErrorLanguage::English => "Method 'WindowState' of object '_Application' failed",
                };
                host_error_described(1004, description)
            } else { self.scalar_set_error("Application", "WindowState") }
        })?;
        let numeric_enum = matches!(value, Value::Int16(_) | Value::Integer(_) | Value::LongLong(_)
            | Value::Byte(_) | Value::Single(_) | Value::Double(_) | Value::Decimal(_))
            && Self::legacy_assignment_number(value).ok().flatten().is_some_and(|number|
                number == -4140.0 || number == -4137.0 || number == -4143.0);
        Ok(if numeric_enum { state } else {
            match state { -4140 => -4143, -4137 => -4140, _ => -4137 }
        })
    }

    fn requested_display_state(&self, value: &Value, invalid: impl Fn() -> String) -> Result<i64, String> {
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
        if object_operand { return Err(invalid()); }
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
                    invalid()
                } else { host_error(number, "WindowState cannot convert this operand") })?
        };
        let number = number.filter(|number| number.is_finite())
            .ok_or_else(&invalid)?;
        let enum_operand = matches!(value, Value::Int16(_) | Value::Integer(_) | Value::LongLong(_)
            | Value::Byte(_) | Value::Single(_) | Value::Double(_) | Value::Decimal(_));
        let internal = if enum_operand {
            match number { -4140.0 => 1.0, -4137.0 => 2.0, -4143.0 => 3.0, _ => number }
        } else { number };
        // Compare the endpoints directly: adding the margin before floor
        // would collapse adjacent Double operands at the upper boundary.
        let margin = 2049.0 / 8589934592.0;
        if !((1.0 - margin)..(4.0 - margin)).contains(&internal) {
            return Err(invalid());
        }
        // Apply the near-integer tolerance when selecting the requested enum.
        // Direct comparisons retain adjacent Double operands at the boundaries.
        Ok(if internal < 2.0 - margin { -4140 }
            else if internal < 3.0 - margin { -4137 }
            else { -4143 })
    }

    pub(super) fn set_window_state(&mut self, value: &Value) -> Result<(), String> {
        let requested = self.requested_window_state(value)?;
        let id = self.window_frames.loaded;
        if self.window_frames.frames[id].state == requested { return Ok(()); }
        self.window_frames.frames[id].state = requested;
        if requested != -4140 {
            self.activate_window_frame(id);
        } else if self.window_frames.active == id {
            if let Some(next) = self.window_frames.open_ids().into_iter()
                .find(|next| *next != id && self.window_frames.frames[*next].state != -4140) {
                self.activate_window_frame(next);
            }
        }
        Ok(())
    }

    pub(super) fn frame_caption(&self) -> Value {
        if let Some(caption) = &self.window_frames.frames[self.window_frames.loaded].creation_caption {
            return Value::String(caption.clone());
        }
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
    fn logical_frame(&self) -> Option<(u64, usize)> {
        let owner = self.active;
        let host = self.hosts.get(&owner)?;
        (!self.closed.contains(&owner) && host.window_is_open(host.window_frames.active))
            .then_some((owner, host.window_frames.active))
    }

    fn foreground_frame(&self) -> Option<(u64, usize)> {
        if let Some((owner, id)) = self.foreground_window {
            if !self.closed.contains(&owner) && self.hosts.get(&owner).is_some_and(|host| host.window_is_open(id)) {
                return Some((owner, id));
            }
        }
        self.logical_frame()
    }

    pub(super) fn pump_window_messages(&mut self) {
        let messages = std::mem::take(&mut self.window_ui_messages);
        for message in messages {
            let (owner, id) = match message {
                WindowUiMessage::FrameChanged(owner, id) | WindowUiMessage::ActivateForeground(owner, id) => (owner, id),
            };
            if self.foreground_window != Some((owner, id)) { continue; }
            if self.closed.contains(&owner) || !self.hosts.get(&owner).is_some_and(|host| host.window_is_open(id)) {
                self.foreground_window = None;
                continue;
            }
            match message {
                WindowUiMessage::FrameChanged(..) => self.window_ui_messages
                    .push_back(WindowUiMessage::ActivateForeground(owner, id)),
                WindowUiMessage::ActivateForeground(..) => self.activate_frame(owner, id),
            }
        }
    }

    pub(super) fn finish_window_messages(&mut self) {
        while !self.window_ui_messages.is_empty() { self.pump_window_messages(); }
    }

    fn queue_window_activation_events(&mut self, previous: Option<(u64, usize)>, current: Option<(u64, usize)>) -> Option<u64> {
        if previous == current || self.application_owner.and_then(|owner| self.hosts.get(&owner))
            .is_some_and(|host| !host.enable_events) { return None; }
        let pending_start = self.pending_events.len();
        let previous_owner = previous.map(|(owner, _)| owner);
        let current_owner = current.map(|(owner, _)| owner);
        let context = |owner| WorkbookEventContext { frame: None, owner, activation_owner: None, saving: false };
        if let Some((owner, id)) = previous {
            let window = self.hosts.get_mut(&owner).expect("previous window owner").object(HostObject::Window(id));
            self.queue_workbook_event(context(owner), "WindowDeactivate", vec![window], vec![]);
            if previous_owner != current_owner {
                self.queue_workbook_event(context(owner), "Deactivate", vec![], vec![]);
            }
        }
        if let Some((owner, id)) = current {
            if previous_owner != current_owner {
                self.queue_workbook_event(context(owner), "Activate", vec![], vec![]);
            }
            let window = self.hosts.get_mut(&owner).expect("active window owner").object(HostObject::Window(id));
            self.queue_workbook_event(context(owner), "WindowActivate", vec![window], vec![]);
        }
        // Only the head can run. Publishing its tail after the answer prevents
        // getters inside a handler from delivering later notifications early.
        let mut calls = self.pending_events.split_off(pending_start);
        let completion = calls.back().map(|event| event.key);
        if let Some(first) = calls.pop_front() {
            if !calls.is_empty() { self.window_event_tails.insert(first.key, calls); }
            if self.running_events.is_empty() {
                self.pending_events.push_back(first);
            } else {
                self.pending_events.push_front(first);
            }
        }
        completion
    }

    pub(super) fn finish_window_creation(&mut self, owner: u64, id: usize) {
        if let Some(frame) = self.hosts.get_mut(&owner).and_then(|host| host.window_frames.frames.get_mut(id)) {
            frame.creation_caption = None;
        }
    }

    fn active_window_display_state(&self) -> Option<i64> {
        let (owner, id) = self.foreground_frame()?;
        Some(self.hosts[&owner].window_frames.frames[id].state)
    }

    fn next_display_window(&self, excluded: (u64, usize)) -> Option<(u64, usize)> {
        self.window_z_order.iter().rev().copied().find(|(owner, id)| {
            (*owner, *id) != excluded && !self.closed.contains(owner)
                && self.hosts.get(owner).is_some_and(|host|
                    host.window_is_open(*id) && host.window_frames.frames[*id].state != -4140)
        })
    }

    fn reported_window_state(&self, owner: u64, id: usize) -> i64 {
        let host = &self.hosts[&owner];
        let state = host.window_frames.frames[id].state;
        if state == -4137 && (self.active != owner || host.window_frames.active != id) {
            -4143
        } else {
            state
        }
    }

    fn activate_frame(&mut self, owner: u64, id: usize) {
        let previous = self.logical_frame();
        self.activate(owner);
        if self.active == owner {
            self.hosts.get_mut(&owner).expect("window owner").activate_window_frame(id);
            self.activate(owner);
        }
        self.queue_window_activation_events(previous, self.logical_frame());
    }

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
    pub(super) fn windows_set(&mut self, receiver: &ObjectRef, name: &str, value: &Value, typed: bool)
        -> Option<Result<bool, String>> {
        if !name.eq_ignore_ascii_case("WindowState") { return None; }
        let mut application = self.application.borrow_mut();
        if application.is_application(receiver) {
            let (owner, id) = application.foreground_frame()?;
            application.ensure_application_owner(owner);
            application.last_owner = owner;
            let requested = match application.hosts[&owner].requested_application_window_state(value, typed) {
                Ok(state) => state,
                Err(error) => return Some(Err(error)),
            };
            if application.hosts[&owner].window_frames.frames[id].state != requested {
                application.hosts.get_mut(&owner).expect("foreground owner").window_frames.frames[id].state = requested;
                if requested == -4140 {
                    application.foreground_window = application.next_display_window((owner, id));
                    if let Some((owner, id)) = application.foreground_window {
                        application.window_ui_messages.push_back(WindowUiMessage::FrameChanged(owner, id));
                    }
                }
            }
            return Some(Ok(true));
        }
        let owner = application.native_owner(receiver)?;
        let id = application.hosts.get(&owner)?.window_id(receiver)?;
        if let Some(error) = application.closed_object_error(receiver) { return Some(Err(error)); }
        if !application.hosts[&owner].window_is_open(id) {
            return Some(Err(host_error(424, "The Window object is no longer available")));
        }
        application.ensure_application_owner(owner);
        application.last_owner = owner;
        let requested = match application.hosts[&owner].requested_window_state(value) {
            Ok(state) => state,
            Err(error) => return Some(Err(error)),
        };
        let host = &application.hosts[&owner];
        if host.window_frames.frames[id].state == requested { return Some(Ok(true)); }
        let was_active = application.active == owner && host.window_frames.active == id;
        application.hosts.get_mut(&owner).expect("window owner").window_frames.frames[id].state = requested;
        if requested != -4140 {
            application.activate_frame(owner, id);
        } else if was_active {
            let next = application.next_display_window((owner, id));
            if let Some((next_owner, next_id)) = next { application.activate_frame(next_owner, next_id); }
        }
        Some(Ok(true))
    }

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
        if global && name.eq_ignore_ascii_case("WindowState") && args.is_empty() {
            if let Some(state) = application.active_window_display_state() {
                return Some(Ok(Some(Value::Integer(state))));
            }
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
            let previous = application.logical_frame();
            let primary = source.is_none().then(|| {
                *application.hosts[&owner].window_frames.open_ids().last()
                    .expect("NewWindow has a source frame")
            }).filter(|id| application.hosts[&owner].window_frames.frames[*id].state != -4140);
            let state = primary.map(|id| application.hosts[&owner].window_frames.frames[id].state)
                .or_else(|| application.active_window_display_state()).unwrap_or(-4137);
            // Workbook.NewWindow uses the primary frame for display-state
            // inheritance and return focus. Its sheet/view copy remains
            // independent of that display history. Window.NewWindow keeps
            // the previous foreground frame as the return target.
            if application.window_application_visible() {
                if let Some(id) = primary {
                    application.window_z_order.retain(|held| *held != (owner, id));
                    application.window_z_order.push((owner, id));
                }
            }
            let result = application.hosts.get_mut(&owner).expect("NewWindow owner").new_window(source, state);
            let id = application.hosts[&owner].window_frames.active;
            application.window_z_order.push((owner, id));
            application.window_enumeration_available = true;
            application.activate(owner);
            let event_previous = match (previous, primary) {
                (Some((previous_owner, _)), Some(primary)) if previous_owner == owner => Some((owner, primary)),
                _ => previous,
            };
            let current = application.logical_frame();
            if let Some(completion) = application.queue_window_activation_events(event_previous, current) {
                // Notifications precede attachment of the workbook's window
                // caption. The temporary blank-sheet frame has its own title.
                let host = application.hosts.get_mut(&owner).expect("created window owner");
                let title = host.file_io.full_name.as_deref()
                    .map(|path| path.rsplit(['\\', '/']).next().unwrap_or(path))
                    .unwrap_or("Sheet1");
                host.window_frames.frames[id].creation_caption = Some(format!("{title}  -  1"));
                application.window_creation_completions.insert(completion, (owner, id));
            }
            return Some(Ok(Some(result)));
        }
        if let Some(owner) = book.filter(|_| name.eq_ignore_ascii_case("Activate")) {
            if !args.is_empty() { return Some(Err(host_error(450, "Workbook.Activate takes no arguments"))); }
            let previous = application.logical_frame();
            application.activate(owner);
            if application.active == owner {
                let host = application.hosts.get_mut(&owner).expect("Workbook.Activate owner");
                if let Some(primary) = host.window_frames.open_ids().last().copied() {
                    host.activate_window_frame(primary);
                }
                application.activate(owner);
            }
            let current = application.logical_frame();
            application.queue_window_activation_events(previous, current);
            return Some(Ok(Some(Value::Empty)));
        }
        if let Some((owner, id)) = window {
            if name.eq_ignore_ascii_case("WindowState") && args.is_empty() {
                return Some(Ok(Some(Value::Integer(application.reported_window_state(owner, id)))));
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
                application.activate_frame(owner, id);
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
                    if active_owner && closing_active {
                        if let Some((next_owner, next_id)) = application.next_display_window((owner, id)) {
                            application.activate_frame(next_owner, next_id);
                        }
                    }
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
                    visible: Some(group["initial_application_visible"].as_bool().unwrap_or(false)),
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
            for (at, (actual, expected)) in actual.iter().zip(expected).enumerate() {
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
    fn window_frame_display_transitions_match_native() {
        native_group(17);
    }

    #[test]
    fn window_cross_book_display_transitions_match_native() {
        native_group(18);
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
        native_group(7);
    }

    #[test]
    fn window_state_assignment_error_boundaries_match_native() {
        native_group(8);
    }

    #[test]
    fn window_state_assignment_contexts_match_native() {
        native_group(9);
    }

    #[test]
    fn window_state_conversion_controls_match_native() {
        native_group(10);
    }

    #[test]
    fn window_state_numeric_rounding_match_native() {
        native_group(11);
    }

    #[test]
    fn window_state_numeric_thresholds_match_native() {
        native_group(12);
    }

    #[test]
    fn window_state_numeric_boundaries_match_native() {
        native_group(13);
    }

    #[test]
    fn window_state_enum_boundaries_match_native() {
        native_group(14);
    }

    #[test]
    fn window_state_operand_precision_match_native() {
        native_group(15);
    }

    #[test]
    fn window_state_binary_boundaries_match_native() {
        native_group(16);
    }
}
