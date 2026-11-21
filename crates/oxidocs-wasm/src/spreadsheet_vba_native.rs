// SPDX-License-Identifier: MIT OR Apache-2.0

//! The workbook's object model for native callers: what a macro reaches as
//! `Application`, handed to code outside VBA -- a COM add-in host exposes it
//! through IDispatch -- one member at a time.
//!
//! The session owns its workbook and keeps the host's state (object handles,
//! selection, settings) between calls, so an object handed out by one call
//! still answers the next.

use super::*;

pub struct NativeAutomation {
    workbook: Box<Workbook>,
    state: Option<WorkbookHostState>,
    original: Vec<u8>,
}

impl NativeAutomation {
    /// A session on the workbook in these .xlsx/.xlsm bytes.
    pub fn open(bytes: &[u8]) -> Result<Self, String> {
        let workbook = oxicells_core::parse_xlsx(bytes).map_err(|e| e.to_string())?;
        Ok(Self { workbook: Box::new(workbook), state: None, original: bytes.to_vec() })
    }

    fn with_host<R>(&mut self, f: impl FnOnce(&mut WorkbookHost<'_>) -> Result<R, String>) -> Result<R, String> {
        let mut host = match self.state.take() {
            Some(state) => WorkbookHost::resume_storage(WorkbookStorage::Borrowed(&mut self.workbook), state),
            None => WorkbookHost::new(&mut self.workbook, 0)?,
        };
        let answer = f(&mut host);
        let (_, state) = host.into_parts();
        self.state = Some(state);
        answer
    }

    /// The Application object.
    pub fn application(&mut self) -> Result<ObjectRef, String> {
        let module = oxivba_core::parse_module("Public Function OxiApplication() As Object\n    Set OxiApplication = Application\nEnd Function\n")
            .map_err(|e| format!("{e:?}"))?;
        self.with_host(|host| match oxivba_core::execute_with_host(&module, "OxiApplication", vec![], host).map_err(|e| format!("{e:?}"))? {
            Value::Object(object) => Ok(object),
            other => Err(format!("Application answered {other:?}")),
        })
    }

    /// A property read (`obj.Name`), `None` when the object has no such member.
    pub fn get(&mut self, object: &ObjectRef, name: &str) -> Result<Option<Value>, String> {
        self.with_host(|host| host.get(object, name))
    }

    /// A property write (`obj.Name = value`).
    pub fn set(&mut self, object: &ObjectRef, name: &str, value: Value) -> Result<bool, String> {
        self.with_host(|host| host.set(object, name, value))
    }

    /// A method call or an indexed property (`obj.Name(args)`).
    pub fn call(&mut self, object: &ObjectRef, name: &str, args: &[Value]) -> Result<Option<Value>, String> {
        self.with_host(|host| host.call(Some(object), name, args))
    }

    /// The workbook as it now stands, saved over the bytes it was opened from.
    pub fn save(&self) -> Result<Vec<u8>, String> {
        let mut editor = oxicells_core::XlsxEditor::new(&self.original).map_err(|e| e.to_string())?;
        editor.apply_workbook(&self.workbook).map_err(|e| e.to_string())?;
        editor.save().map_err(|e| e.to_string())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn a_native_caller_walks_the_object_model() {
        let original = include_bytes!("../../../tests/fixtures/basic_test.xlsx");
        let mut session = NativeAutomation::open(original).unwrap();
        let app = session.application().unwrap();
        let Some(Value::Object(sheet)) = session.get(&app, "ActiveSheet").unwrap() else { panic!("no ActiveSheet") };
        let Some(Value::Object(a1)) = session.call(&sheet, "Range", &[Value::String("A1".into())]).unwrap() else { panic!("no Range") };
        assert!(session.set(&a1, "Value", Value::String("hello".into())).unwrap());
        let Some(Value::Object(b1)) = session.call(&sheet, "Range", &[Value::String("B1".into())]).unwrap() else { panic!() };
        session.set(&b1, "Formula", Value::String("=1+2".into())).unwrap();
        let value = session.get(&b1, "Value").unwrap();
        assert!(matches!(value, Some(Value::Double(v)) if v == 3.0) || matches!(value, Some(Value::Integer(3))), "{value:?}");
        let saved = session.save().unwrap();
        let back = oxicells_core::parse_xlsx(&saved).unwrap();
        let a1 = back.sheets[0].rows.iter().find(|r| r.index == 1).and_then(|r| r.cells.iter().find(|c| c.col == 0)).unwrap();
        assert!(matches!(&a1.value, CellValue::String(s) if s == "hello"));
    }
}
