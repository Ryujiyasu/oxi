// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Runs an Excel COM add-in outside Excel.
//!
//! A COM add-in is a class implementing IDTExtensibility2; Excel creates it
//! and calls `OnConnection` with its Application object, and the add-in works
//! the workbooks through that object's IDispatch. Here the Application is
//! Oxi's: every object the workbook's object model hands out (the one VBA
//! macros run against, `NativeAutomation`) is wrapped in a generic IDispatch
//! -- a name gets a DISPID, `Invoke` becomes a property read, a property
//! write or a method call on the object -- and after the add-in's
//! connection, startup and shutdown calls the workbook is saved.
//!
//!   oxi-com-addin-host <ProgID> <in.xlsx> <out.xlsx>

#[cfg(windows)]
mod host {
    use oxidocs_wasm::NativeAutomation;
    use oxivba_core::{ObjectRef, Value};
    use std::cell::RefCell;
    use std::collections::HashMap;
    use std::ffi::c_void;
    use windows::core::*;
    use windows::Win32::Foundation::*;
    use windows::Win32::System::Com::*;
    use windows::Win32::System::Ole::*;
    use windows::Win32::System::Variant::*;

    /// Office's add-in interface (msaddndr.dll), dual.
    #[windows::core::interface("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")]
    pub unsafe trait IDTExtensibility2: IDispatch {
        fn OnConnection(&self, application: *mut c_void, mode: i32, addin: *mut c_void, custom: *mut *mut SAFEARRAY) -> HRESULT;
        fn OnDisconnection(&self, mode: i32, custom: *mut *mut SAFEARRAY) -> HRESULT;
        fn OnAddInsUpdate(&self, custom: *mut *mut SAFEARRAY) -> HRESULT;
        fn OnStartupComplete(&self, custom: *mut *mut SAFEARRAY) -> HRESULT;
        fn OnBeginShutdown(&self, custom: *mut *mut SAFEARRAY) -> HRESULT;
    }

    thread_local! {
        static SESSION: RefCell<Option<NativeAutomation>> = const { RefCell::new(None) };
        /// Member names by DISPID - 1: one numbering for every object.
        static NAMES: RefCell<Vec<String>> = const { RefCell::new(Vec::new()) };
        /// Our wrappers by their interface pointer, to know one handed back.
        static WRAPPED: RefCell<HashMap<usize, ObjectRef>> = RefCell::new(HashMap::new());
    }

    fn raw(v: &VARIANT) -> (u16, u64) {
        unsafe {
            let p = v as *const VARIANT as *const u8;
            (*(p as *const u16), *(p.add(8) as *const u64))
        }
    }

    fn make(vt: u16, data: u64) -> VARIANT {
        let mut b = [0u64; 3];
        unsafe { *(b.as_mut_ptr() as *mut u16) = vt };
        b[1] = data;
        unsafe { std::mem::transmute::<[u64; 3], VARIANT>(b) }
    }

    fn wrap(object: ObjectRef) -> IDispatch {
        let dispatch: IDispatch = Disp { object: object.clone() }.into();
        WRAPPED.with(|w| w.borrow_mut().insert(dispatch.as_raw() as usize, object));
        dispatch
    }

    fn to_variant(value: Value) -> VARIANT {
        match value {
            Value::Empty | Value::Missing | Value::Nothing => VARIANT::default(),
            Value::Null => make(VT_NULL.0, 0),
            Value::Boolean(b) => VARIANT::from(b),
            Value::Int16(n) => VARIANT::from(i32::from(n)),
            Value::Integer(n) | Value::LongLong(n) => match i32::try_from(n) {
                Ok(n) => VARIANT::from(n),
                Err(_) => VARIANT::from(n as f64),
            },
            Value::Byte(n) => VARIANT::from(i32::from(n)),
            Value::Single(n) => VARIANT::from(f64::from(n)),
            Value::Double(n) => VARIANT::from(n),
            Value::Currency(n) => VARIANT::from(n as f64 / 10_000.0),
            Value::Date(d) => make(VT_DATE.0, d.to_bits()),
            Value::String(s) => VARIANT::from(BSTR::from(s.as_str())),
            Value::Error(code) => make(VT_ERROR.0, u64::from(0x800A_0000u32.wrapping_add(code as u32))),
            Value::Object(object) => {
                let dispatch = wrap(object);
                make(VT_DISPATCH.0, dispatch.into_raw() as u64)
            }
            other => VARIANT::from(BSTR::from(format!("{other:?}"))),
        }
    }

    fn to_value(v: &VARIANT) -> Value {
        let (vt, data) = raw(v);
        match VARENUM(vt) {
            VT_EMPTY => Value::Empty,
            VT_NULL => Value::Null,
            VT_BOOL => Value::Boolean(data as u16 != 0),
            VT_BSTR => Value::String(BSTR::try_from(v).map(|s| s.to_string()).unwrap_or_default()),
            VT_DATE => Value::Date(f64::from_bits(data)),
            VT_ERROR if data as u32 == 0x8002_0004 => Value::Missing,
            VT_ERROR => Value::Error(i64::from((data as u32).wrapping_sub(0x800A_0000))),
            VT_I2 | VT_I4 | VT_UI1 | VT_I8 => Value::Integer(i64::try_from(v).unwrap_or(0)),
            VT_DISPATCH => match WRAPPED.with(|w| w.borrow().get(&(data as usize)).cloned()) {
                Some(object) => Value::Object(object),
                None => Value::Nothing,
            },
            _ => {
                let mut out = VARIANT::default();
                match unsafe { VariantChangeType(&mut out, v, VAR_CHANGE_FLAGS(0), VT_R8) } {
                    Ok(()) => Value::Double(f64::try_from(&out).unwrap_or(0.0)),
                    Err(_) => Value::Empty,
                }
            }
        }
    }

    /// One of the workbook's objects, answering IDispatch.
    #[implement(IDispatch)]
    struct Disp {
        object: ObjectRef,
    }

    impl IDispatch_Impl for Disp_Impl {
        fn GetTypeInfoCount(&self) -> Result<u32> {
            Ok(0)
        }
        fn GetTypeInfo(&self, _: u32, _: u32) -> Result<ITypeInfo> {
            Err(E_NOTIMPL.into())
        }
        fn GetIDsOfNames(&self, _: *const GUID, names: *const PCWSTR, count: u32, _: u32, ids: *mut i32) -> Result<()> {
            NAMES.with(|known| {
                let mut known = known.borrow_mut();
                for i in 0..count as usize {
                    let name = unsafe { (*names.add(i)).to_string().unwrap_or_default() };
                    let at = match known.iter().position(|n| n.eq_ignore_ascii_case(&name)) {
                        Some(at) => at,
                        None => {
                            known.push(name);
                            known.len() - 1
                        }
                    };
                    unsafe { *ids.add(i) = at as i32 + 1 };
                }
            });
            Ok(())
        }
        fn Invoke(&self, id: i32, _: *const GUID, _: u32, flags: DISPATCH_FLAGS, params: *const DISPPARAMS, result: *mut VARIANT, _: *mut EXCEPINFO, _: *mut u32) -> Result<()> {
            let Some(name) = NAMES.with(|n| n.borrow().get((id - 1).max(0) as usize).cloned()) else {
                return Err(DISP_E_MEMBERNOTFOUND.into());
            };
            // Positional arguments come last first; a property write's value
            // is the named DISPID_PROPERTYPUT argument, first in the array.
            let args: Vec<Value> = unsafe {
                let p = &*params;
                (0..p.cArgs as usize).rev().map(|i| to_value(&*p.rgvarg.add(i))).collect()
            };
            let object = self.object.clone();
            let answer: std::result::Result<Option<Value>, String> = SESSION.with(|s| {
                let mut s = s.borrow_mut();
                let session = s.as_mut().ok_or_else(|| "no session".to_string())?;
                if flags.0 & (DISPATCH_PROPERTYPUT.0 | DISPATCH_PROPERTYPUTREF.0) != 0 {
                    let value = args.last().cloned().unwrap_or(Value::Empty);
                    let target = if args.len() > 1 {
                        // obj.Name(index) = value: the indexed member's default.
                        match session.call(&object, &name, &args[..args.len() - 1])? {
                            Some(Value::Object(inner)) => return session.set(&inner, "_Default", value).map(|_| None),
                            _ => return Err(format!("{name} cannot be set with indices")),
                        }
                    } else {
                        object
                    };
                    return session.set(&target, &name, value).map(|done| if done { None } else { Some(Value::Missing) });
                }
                if args.is_empty() && flags.0 & DISPATCH_PROPERTYGET.0 != 0 {
                    if let Some(value) = session.get(&object, &name)? {
                        return Ok(Some(value));
                    }
                }
                session.call(&object, &name, &args)
            });
            match answer {
                Ok(Some(Value::Missing)) => Err(DISP_E_MEMBERNOTFOUND.into()),
                Ok(value) => {
                    if !result.is_null() {
                        unsafe { std::ptr::write(result, to_variant(value.unwrap_or(Value::Empty))) };
                    }
                    Ok(())
                }
                Err(_) => Err(DISP_E_EXCEPTION.into()),
            }
        }
    }

    pub fn run(prog_id: &str, input: &str, output: &str) -> std::result::Result<serde_json::Value, String> {
        unsafe {
            let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
        }
        let bytes = std::fs::read(input).map_err(|e| format!("{input}: {e}"))?;
        let mut session = NativeAutomation::open(&bytes)?;
        let app = session.application()?;
        SESSION.with(|s| *s.borrow_mut() = Some(session));
        let application = wrap(app);
        let addin: IDTExtensibility2 = unsafe {
            let class = CLSIDFromProgID(&HSTRING::from(prog_id)).map_err(|e| format!("no class {prog_id}: {e}"))?;
            let unknown: IUnknown = CoCreateInstance(&class, None, CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER).map_err(|e| format!("cannot create {prog_id}: {e}"))?;
            unknown.cast().map_err(|e| format!("{prog_id} is not a COM add-in (no IDTExtensibility2): {e}"))?
        };
        let mut steps = serde_json::Map::new();
        unsafe {
            let mut custom = SafeArrayCreateVector(VT_VARIANT, 0, 0);
            // ext_cm_AfterStartup (0): connected once the host is up, as
            // Excel's COMAddIns(...).Connect = True does.
            let app_raw = application.as_raw();
            steps.insert("OnConnection".into(), serde_json::json!(format!("{:?}", addin.OnConnection(app_raw, 0, app_raw, &mut custom))));
            steps.insert("OnStartupComplete".into(), serde_json::json!(format!("{:?}", addin.OnStartupComplete(&mut custom))));
            steps.insert("OnBeginShutdown".into(), serde_json::json!(format!("{:?}", addin.OnBeginShutdown(&mut custom))));
            steps.insert("OnDisconnection".into(), serde_json::json!(format!("{:?}", addin.OnDisconnection(1, &mut custom))));
            let _ = SafeArrayDestroy(custom);
        }
        let saved = SESSION.with(|s| s.borrow().as_ref().map(NativeAutomation::save)).ok_or("no session")??;
        std::fs::write(output, saved).map_err(|e| format!("{output}: {e}"))?;
        Ok(serde_json::Value::Object(steps))
    }
}

#[cfg(windows)]
fn main() {
    let args: Vec<String> = std::env::args().collect();
    if args.len() < 4 {
        eprintln!("usage: oxi-com-addin-host <ProgID> <in.xlsx> <out.xlsx>");
        std::process::exit(2);
    }
    match host::run(&args[1], &args[2], &args[3]) {
        Ok(steps) => println!("{steps}"),
        Err(error) => {
            eprintln!("{error}");
            std::process::exit(1);
        }
    }
}

#[cfg(not(windows))]
fn main() {
    eprintln!("COM add-ins are Windows components; this host runs on Windows only.");
    std::process::exit(1);
}
