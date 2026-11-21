// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Calls an Excel COM Automation add-in outside Excel.
//!
//! An Automation add-in is a COM class (registered by ProgID) whose IDispatch
//! methods a worksheet formula can call once Excel has it as an add-in. Here
//! the class is created by its ProgID and each call goes through
//! `GetIDsOfNames` and `Invoke`, the way Excel makes it: arguments as
//! VARIANTs, last first; a range of values as a two-dimensional SAFEARRAY of
//! VARIANTs with bounds from 1 (what `Range.Value` is); the answer read back
//! from its VARIANT, an array included.
//!
//!   oxi-com-host <ProgID> call <Method> '<json array of arguments>'
//!   oxi-com-host --native-messaging   (or a chrome-extension:// origin)

use std::ffi::c_void;

const VT_ARRAY_: u16 = 0x2000;
const VT_BYREF_: u16 = 0x4000;

#[cfg(windows)]
mod com {
    use super::*;
    use windows::core::*;
    use windows::Win32::System::Com::*;
    use windows::Win32::System::Ole::*;
    use windows::Win32::System::Variant::*;

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

    /// Excel's error values as VT_ERROR scodes (0x800A0000 + 2000 + n).
    const ERRORS: [(u32, &str); 7] =
        [(2000, "#NULL!"), (2007, "#DIV/0!"), (2015, "#VALUE!"), (2023, "#REF!"), (2029, "#NAME?"), (2036, "#NUM!"), (2042, "#N/A")];

    fn to_variant(v: &serde_json::Value) -> VARIANT {
        match v {
            serde_json::Value::Number(n) => VARIANT::from(n.as_f64().unwrap_or(0.0)),
            serde_json::Value::Bool(b) => VARIANT::from(*b),
            serde_json::Value::String(s) => match ERRORS.iter().find(|(_, t)| t == s) {
                Some((code, _)) => make(VT_ERROR.0, u64::from(0x800A_0000u32 + code)),
                None => VARIANT::from(BSTR::from(s.as_str())),
            },
            serde_json::Value::Array(outer) => {
                let rows: Vec<Vec<serde_json::Value>> = if outer.iter().all(serde_json::Value::is_array) {
                    outer.iter().map(|r| r.as_array().cloned().unwrap_or_default()).collect()
                } else {
                    vec![outer.clone()]
                };
                let cols = rows.iter().map(Vec::len).max().unwrap_or(0) as u32;
                let bounds = [SAFEARRAYBOUND { cElements: rows.len() as u32, lLbound: 1 }, SAFEARRAYBOUND { cElements: cols, lLbound: 1 }];
                unsafe {
                    let array = SafeArrayCreate(VT_VARIANT, 2, bounds.as_ptr());
                    for (r, row) in rows.iter().enumerate() {
                        for c in 0..cols as usize {
                            let cell = row.get(c).map(to_variant).unwrap_or_default();
                            let index = [r as i32 + 1, c as i32 + 1];
                            let _ = SafeArrayPutElement(array, index.as_ptr(), &cell as *const VARIANT as *const c_void);
                        }
                    }
                    make(VT_ARRAY_ | VT_VARIANT.0, array as u64)
                }
            }
            // An omitted argument: VT_ERROR DISP_E_PARAMNOTFOUND, as Excel passes it.
            serde_json::Value::Null => make(VT_ERROR.0, 0x8002_0004),
            serde_json::Value::Object(_) => VARIANT::default(),
        }
    }

    fn to_json(v: &VARIANT) -> serde_json::Value {
        let (vt, data) = raw(v);
        if vt & VT_BYREF_ != 0 && vt & !VT_BYREF_ == VT_VARIANT.0 {
            return to_json(unsafe { &*(data as *const VARIANT) });
        }
        if vt & VT_ARRAY_ != 0 {
            unsafe {
                let array = data as *const SAFEARRAY;
                let dims = SafeArrayGetDim(array);
                let bound = |d: u32| (SafeArrayGetLBound(array, d).unwrap_or(0), SafeArrayGetUBound(array, d).unwrap_or(-1));
                let (r0, r1) = bound(1);
                let (c0, c1) = if dims >= 2 { bound(2) } else { (0, 0) };
                let base = vt & !VT_ARRAY_;
                let element = |index: &[i32]| -> serde_json::Value {
                    if base == VT_VARIANT.0 {
                        let mut cell = VARIANT::default();
                        let _ = SafeArrayGetElement(array, index.as_ptr(), &mut cell as *mut VARIANT as *mut c_void);
                        to_json(&cell)
                    } else if base == VT_R8.0 {
                        let mut x = 0f64;
                        let _ = SafeArrayGetElement(array, index.as_ptr(), &mut x as *mut f64 as *mut c_void);
                        serde_json::json!(x)
                    } else {
                        serde_json::Value::Null
                    }
                };
                if dims < 2 {
                    return serde_json::Value::Array((r0..=r1).map(|r| element(&[r])).collect());
                }
                return serde_json::Value::Array(
                    (r0..=r1).map(|r| serde_json::Value::Array((c0..=c1).map(|c| element(&[r, c])).collect())).collect(),
                );
            }
        }
        match VARENUM(vt) {
            VT_EMPTY | VT_NULL => serde_json::Value::Null,
            VT_BOOL => serde_json::json!(data as u16 != 0),
            VT_BSTR => serde_json::json!(BSTR::try_from(v).map(|s| s.to_string()).unwrap_or_default()),
            VT_ERROR => {
                let code = (data as u32).wrapping_sub(0x800A_0000);
                ERRORS.iter().find(|(c, _)| *c == code).map_or(serde_json::json!("#VALUE!"), |(_, t)| serde_json::json!(t))
            }
            _ => {
                let mut out = VARIANT::default();
                match unsafe { VariantChangeType(&mut out, v, VAR_CHANGE_FLAGS(0), VT_R8) } {
                    Ok(()) => serde_json::json!(f64::try_from(&out).unwrap_or(0.0)),
                    Err(_) => serde_json::Value::Null,
                }
            }
        }
    }

    /// A created add-in class.
    pub struct AddIn(IDispatch);

    impl AddIn {
        pub fn create(prog_id: &str) -> std::result::Result<AddIn, String> {
            unsafe {
                let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
                let class = CLSIDFromProgID(&HSTRING::from(prog_id)).map_err(|e| format!("no class {prog_id}: {e}"))?;
                let object: IDispatch = CoCreateInstance(&class, None, CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER)
                    .map_err(|e| format!("cannot create {prog_id}: {e}"))?;
                Ok(AddIn(object))
            }
        }

        /// `{"value": ...}`, or `{"error": "#NAME?"}` for a method it has not.
        pub fn call(&self, method: &str, given: &[serde_json::Value]) -> serde_json::Value {
            unsafe {
                let name = HSTRING::from(method);
                let names = [PCWSTR(name.as_ptr())];
                let mut id = 0i32;
                if self.0.GetIDsOfNames(&GUID::zeroed(), names.as_ptr(), 1, 0, &mut id).is_err() {
                    return serde_json::json!({"error": "#NAME?"});
                }
                // Last first, as Invoke takes them.
                let mut args: Vec<VARIANT> = given.iter().rev().map(to_variant).collect();
                let params = DISPPARAMS { rgvarg: args.as_mut_ptr(), rgdispidNamedArgs: std::ptr::null_mut(), cArgs: args.len() as u32, cNamedArgs: 0 };
                let mut result = VARIANT::default();
                match self.0.Invoke(id, &GUID::zeroed(), 0, DISPATCH_METHOD, &params, Some(&mut result), None, None) {
                    Ok(()) => serde_json::json!({"value": to_json(&result)}),
                    // Excel shows #VALUE! for a call the add-in fails.
                    Err(_) => serde_json::json!({"value": "#VALUE!"}),
                }
            }
        }
    }
}

fn read_message(input: &mut impl std::io::Read) -> Option<serde_json::Value> {
    let mut size = [0u8; 4];
    input.read_exact(&mut size).ok()?;
    let mut body = vec![0u8; u32::from_le_bytes(size) as usize];
    input.read_exact(&mut body).ok()?;
    serde_json::from_slice(&body).ok()
}

fn write_message(output: &mut impl std::io::Write, value: &serde_json::Value) -> std::io::Result<()> {
    let body = serde_json::to_vec(value)?;
    output.write_all(&(body.len() as u32).to_le_bytes())?;
    output.write_all(&body)?;
    output.flush()
}

/// `{"id", "progId", "function", "args"}` in, `{"id", "value" | "error"}` out;
/// each class created once and kept.
#[cfg(windows)]
fn serve_native_messaging() {
    let mut made: std::collections::HashMap<String, com::AddIn> = std::collections::HashMap::new();
    let (stdin, stdout) = (std::io::stdin(), std::io::stdout());
    let (mut input, mut output) = (stdin.lock(), stdout.lock());
    while let Some(request) = read_message(&mut input) {
        let id = request.get("id").cloned().unwrap_or(serde_json::Value::Null);
        let prog_id = request.get("progId").and_then(|v| v.as_str()).unwrap_or_default().to_string();
        let mut answer = if !made.contains_key(&prog_id) {
            match com::AddIn::create(&prog_id) {
                Ok(addin) => {
                    made.insert(prog_id.clone(), addin);
                    serde_json::json!({})
                }
                Err(error) => serde_json::json!({"error": error}),
            }
        } else {
            serde_json::json!({})
        };
        if answer.get("error").is_none() {
            let name = request.get("function").and_then(|v| v.as_str()).unwrap_or_default();
            let args = request.get("args").and_then(|v| v.as_array()).cloned().unwrap_or_default();
            answer = made[&prog_id].call(name, &args);
        }
        answer["id"] = id;
        if write_message(&mut output, &answer).is_err() {
            break;
        }
    }
}

#[cfg(windows)]
fn main() {
    let args: Vec<String> = std::env::args().collect();
    if args.get(1).is_some_and(|a| a.starts_with("chrome-extension://") || a == "--native-messaging") {
        serve_native_messaging();
        return;
    }
    if args.len() < 4 || args[2] != "call" {
        eprintln!("usage: oxi-com-host <ProgID> call <Method> '<json args>'");
        std::process::exit(2);
    }
    let addin = match com::AddIn::create(&args[1]) {
        Ok(addin) => addin,
        Err(error) => {
            eprintln!("{error}");
            std::process::exit(1);
        }
    };
    let given: Vec<serde_json::Value> = serde_json::from_str(args.get(4).map(String::as_str).unwrap_or("[]")).unwrap_or_default();
    println!("{}", addin.call(&args[3], &given));
}

#[cfg(not(windows))]
fn main() {
    eprintln!("COM add-ins are Windows components; this host runs on Windows only.");
    std::process::exit(1);
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn native_messages_are_length_prefixed_json() {
        let mut wire = Vec::new();
        write_message(&mut wire, &serde_json::json!({"id": 1, "progId": "X.Y"})).unwrap();
        assert_eq!(u32::from_le_bytes(wire[..4].try_into().unwrap()) as usize, wire.len() - 4);
        assert_eq!(read_message(&mut std::io::Cursor::new(wire)).unwrap(), serde_json::json!({"id": 1, "progId": "X.Y"}));
    }
}
