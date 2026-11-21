// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Runs an Excel XLL add-in outside Excel.
//!
//! An XLL is a DLL built against the Excel C API. Excel loads it, calls its
//! `xlAutoOpen`, and the add-in calls back -- `Excel12v(xlfRegister, ...)` --
//! to declare each worksheet function: the exported procedure, a type string
//! (`BBB` is a double of two doubles, `QQ` an XLOPER12 of an XLOPER12, ...)
//! and the name a formula calls it by. The SDK's xlcall.cpp finds the
//! callback by asking the running executable for its `MdCallBack12` export,
//! so this binary exports one (build.rs) and answers the calls an add-in
//! makes on opening: `xlfRegister`, `xlGetName`, `xlFree`, `xlCoerce`.
//!
//! A registered function is then called with arguments marshalled by its type
//! string. Win64 passes argument i in integer register i or vector register i
//! by its type; both are loaded with the same bits, so the callee reads
//! whichever its own signature names, and arguments past the fourth go on the
//! stack eight bytes apiece (`call_win64`).
//!
//!   oxi-xll-host <add-in.xll> list
//!   oxi-xll-host <add-in.xll> call <Function> '<json array of arguments>'

use std::cell::RefCell;

#[repr(C)]
#[derive(Clone, Copy, Debug)]
pub struct Xloper12 {
    pub val: [u64; 3],
    pub xltype: u32,
}

const XLTYPE_NUM: u32 = 0x0001;
const XLTYPE_STR: u32 = 0x0002;
const XLTYPE_BOOL: u32 = 0x0004;
const XLTYPE_ERR: u32 = 0x0010;
const XLTYPE_MULTI: u32 = 0x0040;
const XLTYPE_MISSING: u32 = 0x0080;
const XLTYPE_NIL: u32 = 0x0100;
const XLTYPE_INT: u32 = 0x0800;
const XLBIT_XL_FREE: u32 = 0x1000;
const XLBIT_DLL_FREE: u32 = 0x4000;

const XLF_REGISTER: i32 = 149;
const XL_FREE: i32 = 0x4000;
const XL_COERCE: i32 = 0x4002;
const XL_GET_NAME: i32 = 0x4009;
const XLRET_SUCCESS: i32 = 0;
const XLRET_INV_XLFN: i32 = 1;

#[derive(Debug, Clone)]
struct Registration {
    procedure: String,
    types: String,
    function: String,
}

thread_local! {
    static REGISTERED: RefCell<Vec<Registration>> = const { RefCell::new(Vec::new()) };
    static MODULE: RefCell<String> = const { RefCell::new(String::new()) };
    /// Strings handed to the add-in with xlbitXLFree, kept until xlFree.
    static HANDED: RefCell<Vec<Box<[u16]>>> = const { RefCell::new(Vec::new()) };
}

unsafe fn read_str(x: &Xloper12) -> Option<String> {
    if x.xltype & 0xFFF != XLTYPE_STR {
        return None;
    }
    let p = x.val[0] as *const u16;
    if p.is_null() {
        return Some(String::new());
    }
    let len = *p as usize;
    Some(String::from_utf16_lossy(std::slice::from_raw_parts(p.add(1), len)))
}

fn counted(s: &str) -> Box<[u16]> {
    std::iter::once(s.encode_utf16().count().min(32767) as u16).chain(s.encode_utf16()).collect()
}

/// The Excel C API, as much of it as an add-in uses while it opens.
#[no_mangle]
pub unsafe extern "system" fn MdCallBack12(xlfn: i32, count: i32, opers: *mut *mut Xloper12, result: *mut Xloper12) -> i32 {
    let args: Vec<Xloper12> = (0..count.max(0) as usize).map(|i| **opers.add(i)).collect();
    let set = |x: Xloper12| {
        if !result.is_null() {
            *result = x;
        }
    };
    match xlfn {
        XLF_REGISTER => {
            let text = |i: usize| args.get(i).and_then(|a| read_str(a)).unwrap_or_default();
            let registration = Registration { procedure: text(1), types: text(2), function: text(3) };
            REGISTERED.with(|r| r.borrow_mut().push(registration));
            set(Xloper12 { val: [1f64.to_bits(), 0, 0], xltype: XLTYPE_NUM });
            XLRET_SUCCESS
        }
        XL_GET_NAME => {
            let name = MODULE.with(|m| m.borrow().clone());
            let units = counted(&name);
            let x = Xloper12 { val: [units.as_ptr() as u64, 0, 0], xltype: XLTYPE_STR | XLBIT_XL_FREE };
            HANDED.with(|h| h.borrow_mut().push(units));
            set(x);
            XLRET_SUCCESS
        }
        XL_FREE => XLRET_SUCCESS,
        XL_COERCE => {
            // Values only: a number stays a number, anything else passes through.
            if let Some(a) = args.first() {
                set(*a);
            }
            XLRET_SUCCESS
        }
        _ => XLRET_INV_XLFN,
    }
}

/// Call `f` with Win64 argument passing, each argument given as its 64 bits.
/// Returns (RAX, XMM0 as f64): the callee's integer and floating results.
#[cfg(all(windows, target_arch = "x86_64"))]
unsafe fn call_win64(f: usize, args: &[u64]) -> (u64, f64) {
    // The frame as the callee will see it: four shadow slots (holding the
    // first four arguments too, so the registers can be loaded from it), then
    // the fifth argument onward; a multiple of 16 bytes.
    let mut frame = vec![0u64; (args.len().max(4) + 1) & !1];
    frame[..args.len()].copy_from_slice(args);
    let rax: u64;
    let xmm: f64;
    std::arch::asm!(
        "mov r12, rsp",
        "and rsp, -16",
        "sub rsp, rdx",
        "mov rdi, rsp",
        "rep movsq",
        "mov rcx, [rsp]",
        "mov rdx, [rsp + 8]",
        "mov r8, [rsp + 16]",
        "mov r9, [rsp + 24]",
        "movq xmm0, rcx",
        "movq xmm1, rdx",
        "movq xmm2, r8",
        "movq xmm3, r9",
        "call r10",
        "mov rsp, r12",
        // Every register the block or the callee changes is declared as
        // changed: rep movsq walks rcx/rsi/rdi, the callee may use them all.
        inout("rdx") frame.len() * 8 => _,
        inout("rcx") frame.len() => _,
        inout("rsi") frame.as_ptr() => _,
        inout("r10") f => _,
        out("rdi") _,
        out("r12") _,
        out("rax") rax,
        out("xmm0") xmm,
        clobber_abi("C"),
    );
    (rax, xmm)
}

/// One argument as the add-in's type character asks for it, and what has to
/// stay alive while the call runs.
// The fields are never read: they keep the argument's memory alive while the
// add-in is called.
#[allow(dead_code)]
enum Held {
    None,
    Oper(Box<Xloper12>, Vec<Box<[u16]>>, Option<Box<[Xloper12]>>),
    Wide(Vec<u16>),
    Fp(Vec<u64>),
}

/// A JSON value as a value XLOPER12: a 2-D array (rows of values) or a 1-D
/// one (one row) is an xltypeMulti, as Excel passes a range to a `Q`
/// argument; what has to stay alive with it comes back too.
fn json_to_value_oper(v: &serde_json::Value) -> (Xloper12, Vec<Box<[u16]>>, Option<Box<[Xloper12]>>) {
    let serde_json::Value::Array(outer) = v else {
        let (x, keep) = json_to_oper(v);
        return (x, keep.into_iter().collect(), None);
    };
    let rows: Vec<Vec<serde_json::Value>> = if outer.iter().all(serde_json::Value::is_array) {
        outer.iter().map(|r| r.as_array().cloned().unwrap_or_default()).collect()
    } else {
        vec![outer.clone()]
    };
    let cols = rows.iter().map(Vec::len).max().unwrap_or(0);
    let mut keep = Vec::new();
    let mut cells = Vec::with_capacity(rows.len() * cols);
    for row in &rows {
        for c in 0..cols {
            let (x, k) = json_to_oper(row.get(c).unwrap_or(&serde_json::Value::Null));
            // A blank cell of a range is nil, not missing.
            let x = if x.xltype == XLTYPE_MISSING { Xloper12 { val: [0; 3], xltype: XLTYPE_NIL } } else { x };
            cells.push(x);
            keep.extend(k);
        }
    }
    let cells = cells.into_boxed_slice();
    let x = Xloper12 { val: [cells.as_ptr() as u64, (cols as u64) << 32 | rows.len() as u64, 0], xltype: XLTYPE_MULTI };
    (x, keep, Some(cells))
}

/// An FP12 (`K%`) from numbers: { rows, columns, doubles row by row }, or
/// `None` when something is not a number (Excel answers #VALUE!).
fn json_to_fp12(v: &serde_json::Value) -> Option<Vec<u64>> {
    let rows: Vec<Vec<serde_json::Value>> = match v {
        serde_json::Value::Array(outer) if outer.iter().all(serde_json::Value::is_array) => {
            outer.iter().map(|r| r.as_array().cloned().unwrap_or_default()).collect()
        }
        serde_json::Value::Array(outer) => vec![outer.clone()],
        other => vec![vec![other.clone()]],
    };
    let cols = rows.iter().map(Vec::len).max().unwrap_or(0);
    let mut out = vec![(rows.len() as u64) | (cols as u64) << 32];
    for row in &rows {
        for c in 0..cols {
            out.push(row.get(c)?.as_f64()?.to_bits());
        }
    }
    Some(out)
}

fn json_to_oper(v: &serde_json::Value) -> (Xloper12, Option<Box<[u16]>>) {
    match v {
        serde_json::Value::Number(n) => (Xloper12 { val: [n.as_f64().unwrap_or(0.0).to_bits(), 0, 0], xltype: XLTYPE_NUM }, None),
        serde_json::Value::Bool(b) => (Xloper12 { val: [u64::from(*b), 0, 0], xltype: XLTYPE_BOOL }, None),
        serde_json::Value::String(s) => {
            let units = counted(s);
            (Xloper12 { val: [units.as_ptr() as u64, 0, 0], xltype: XLTYPE_STR }, Some(units))
        }
        serde_json::Value::Null => (Xloper12 { val: [0; 3], xltype: XLTYPE_MISSING }, None),
        _ => (Xloper12 { val: [0; 3], xltype: XLTYPE_NIL }, None),
    }
}

unsafe fn oper_to_json(x: &Xloper12) -> serde_json::Value {
    match x.xltype & 0xFFF {
        XLTYPE_NUM => serde_json::json!(f64::from_bits(x.val[0])),
        XLTYPE_INT => serde_json::json!(x.val[0] as i32),
        XLTYPE_BOOL => serde_json::json!(x.val[0] as u32 != 0),
        XLTYPE_STR => serde_json::json!(read_str(x).unwrap_or_default()),
        XLTYPE_ERR => serde_json::json!(match x.val[0] as u32 {
            0 => "#NULL!",
            7 => "#DIV/0!",
            15 => "#VALUE!",
            23 => "#REF!",
            29 => "#NAME?",
            36 => "#NUM!",
            42 => "#N/A",
            other => return serde_json::json!(format!("#ERR{other}")),
        }),
        XLTYPE_MULTI => {
            let p = x.val[0] as *const Xloper12;
            let (rows, cols) = (x.val[1] as u32 as usize, (x.val[1] >> 32) as u32 as usize);
            let rows_v: Vec<serde_json::Value> = (0..rows)
                .map(|r| serde_json::Value::Array((0..cols).map(|c| oper_to_json(&*p.add(r * cols + c))).collect()))
                .collect();
            serde_json::Value::Array(rows_v)
        }
        _ => serde_json::Value::Null,
    }
}

/// The type string split into its codes: one or two characters each
/// (`C%`, `D%`, `Q`, `B`, ...), the return first; `!`, `$` and `#` dropped.
fn type_codes(types: &str) -> Vec<String> {
    let mut out = Vec::new();
    let chars: Vec<char> = types.chars().filter(|c| !matches!(c, '!' | '$' | '#')).collect();
    let mut i = 0;
    while i < chars.len() {
        if i + 1 < chars.len() && chars[i + 1] == '%' {
            out.push(format!("{}%", chars[i]));
            i += 2;
        } else {
            out.push(chars[i].to_string());
            i += 1;
        }
    }
    out
}

/// An add-in loaded and opened: its module and what it registered.
#[cfg(windows)]
struct Opened {
    module: windows::Win32::Foundation::HMODULE,
    registered: Vec<Registration>,
}

#[cfg(windows)]
impl Opened {
    fn proc(&self, name: &str) -> Option<usize> {
        use windows::core::PCSTR;
        use windows::Win32::System::LibraryLoader::GetProcAddress;
        let c = std::ffi::CString::new(name).ok()?;
        unsafe { GetProcAddress(self.module, PCSTR(c.as_ptr() as *const u8)).map(|f| f as usize) }
    }

    fn list(&self) -> serde_json::Value {
        serde_json::Value::Array(
            self.registered
                .iter()
                .map(|r| serde_json::json!({"function": r.function, "procedure": r.procedure, "types": r.types}))
                .collect(),
        )
    }

    /// Call a registered function with arguments marshalled by its type
    /// string; `{"value": ...}` or `{"error": "#NAME?"}`.
    fn call(&self, name: &str, given: &[serde_json::Value]) -> serde_json::Value {
        let Some(reg) = self.registered.iter().find(|r| r.function.eq_ignore_ascii_case(name)) else {
            return serde_json::json!({"error": "#NAME?"});
        };
        let Some(f) = self.proc(&reg.procedure) else {
            return serde_json::json!({"error": "#NAME?"});
        };
        let codes = type_codes(&reg.types);
        let (ret, params) = codes.split_first().map(|(r, p)| (r.clone(), p.to_vec())).unwrap_or_default();
        let mut held: Vec<Held> = Vec::new();
        let mut bits: Vec<u64> = Vec::new();
        for (i, code) in params.iter().enumerate() {
            let v = given.get(i).cloned().unwrap_or(serde_json::Value::Null);
            match code.as_str() {
                "B" => bits.push(v.as_f64().unwrap_or(0.0).to_bits()),
                "J" => bits.push(v.as_f64().unwrap_or(0.0) as i32 as i64 as u64),
                "A" => bits.push(u64::from(v.as_bool().unwrap_or(false))),
                "Q" | "U" | "P" | "R" => {
                    let (x, keep, cells) = json_to_value_oper(&v);
                    let boxed = Box::new(x);
                    bits.push(&*boxed as *const Xloper12 as u64);
                    held.push(Held::Oper(boxed, keep, cells));
                }
                "K%" | "K" => {
                    let Some(fp) = json_to_fp12(&v) else {
                        return serde_json::json!({"value": "#VALUE!"});
                    };
                    bits.push(fp.as_ptr() as u64);
                    held.push(Held::Fp(fp));
                }
                "C%" | "D%" => {
                    let s = v.as_str().map(str::to_string).unwrap_or_else(|| v.to_string());
                    let units: Vec<u16> = if code == "D%" {
                        counted(&s).into_vec()
                    } else {
                        s.encode_utf16().chain(std::iter::once(0)).collect()
                    };
                    bits.push(units.as_ptr() as u64);
                    held.push(Held::Wide(units));
                }
                _ => {
                    held.push(Held::None);
                    bits.push(0);
                }
            }
        }
        let (rax, xmm) = unsafe { call_win64(f, &bits) };
        // The arguments stay alive until the answer has been read: an add-in
        // may hand back what points into them.
        let value = match ret.as_str() {
            "K%" | "K" => {
                let p = rax as *const i32;
                if p.is_null() {
                    serde_json::Value::Null
                } else {
                    unsafe {
                        let (rows, cols) = (*p as usize, *p.add(1) as usize);
                        let data = p.add(2) as *const f64;
                        serde_json::Value::Array(
                            (0..rows)
                                .map(|r| serde_json::Value::Array((0..cols).map(|c| serde_json::json!(*data.add(r * cols + c))).collect()))
                                .collect(),
                        )
                    }
                }
            }
            "B" => serde_json::json!(xmm),
            "J" => serde_json::json!(rax as i32),
            "A" => serde_json::json!(rax as u16 != 0),
            "Q" | "U" | "P" | "R" => {
                let p = rax as *mut Xloper12;
                if p.is_null() {
                    serde_json::Value::Null
                } else {
                    let out = unsafe { oper_to_json(&*p) };
                    if unsafe { (*p).xltype } & XLBIT_DLL_FREE != 0 {
                        if let Some(free) = self.proc("xlAutoFree12") {
                            unsafe {
                                let free: unsafe extern "system" fn(*mut Xloper12) = std::mem::transmute(free);
                                free(p);
                            }
                        }
                    }
                    out
                }
            }
            "C%" | "D%" => {
                let p = rax as *const u16;
                if p.is_null() {
                    serde_json::Value::Null
                } else if ret == "D%" {
                    unsafe { serde_json::json!(String::from_utf16_lossy(std::slice::from_raw_parts(p.add(1), *p as usize))) }
                } else {
                    let mut n = 0;
                    unsafe {
                        while *p.add(n) != 0 {
                            n += 1;
                        }
                        serde_json::json!(String::from_utf16_lossy(std::slice::from_raw_parts(p, n)))
                    }
                }
            }
            _ => serde_json::Value::Null,
        };
        drop(held);
        serde_json::json!({"value": value})
    }
}

/// Load an XLL and run its `xlAutoOpen`, keeping what it registers.
#[cfg(windows)]
fn open_xll(path: &str) -> Result<Opened, String> {
    use windows::core::HSTRING;
    use windows::Win32::System::LibraryLoader::LoadLibraryW;
    let path = std::path::absolute(path).unwrap_or_else(|_| path.into());
    MODULE.with(|m| *m.borrow_mut() = path.to_string_lossy().into_owned());
    let module = unsafe { LoadLibraryW(&HSTRING::from(path.as_os_str())) }.map_err(|e| format!("cannot load {}: {e}", path.display()))?;
    let before = REGISTERED.with(|r| r.borrow().len());
    let mut opened = Opened { module, registered: Vec::new() };
    let Some(open) = opened.proc("xlAutoOpen") else {
        return Err(format!("{} has no xlAutoOpen: not an XLL", path.display()));
    };
    unsafe {
        let open: unsafe extern "system" fn() -> i32 = std::mem::transmute(open);
        open();
    }
    opened.registered = REGISTERED.with(|r| r.borrow()[before..].to_vec());
    Ok(opened)
}

/// A native-messaging message: a little-endian u32 length, then that many
/// bytes of UTF-8 JSON (Chrome's and Edge's framing, both ways).
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

/// The host as the browser extension starts it: requests
/// `{"id", "xll", "op": "list" | "call", "function", "args"}` in, answers
/// `{"id", ...}` out, each XLL opened once and kept.
#[cfg(windows)]
fn serve_native_messaging() {
    let mut opened: std::collections::HashMap<String, Opened> = std::collections::HashMap::new();
    let (stdin, stdout) = (std::io::stdin(), std::io::stdout());
    let (mut input, mut output) = (stdin.lock(), stdout.lock());
    while let Some(request) = read_message(&mut input) {
        let id = request.get("id").cloned().unwrap_or(serde_json::Value::Null);
        let path = request.get("xll").and_then(|v| v.as_str()).unwrap_or_default().to_string();
        let mut answer = if !opened.contains_key(&path) {
            match open_xll(&path) {
                Ok(xll) => {
                    opened.insert(path.clone(), xll);
                    serde_json::json!({})
                }
                Err(error) => serde_json::json!({"error": error}),
            }
        } else {
            serde_json::json!({})
        };
        if answer.get("error").is_none() {
            let xll = &opened[&path];
            answer = match request.get("op").and_then(|v| v.as_str()) {
                Some("list") => serde_json::json!({"functions": xll.list()}),
                Some("call") => {
                    let name = request.get("function").and_then(|v| v.as_str()).unwrap_or_default();
                    let args = request.get("args").and_then(|v| v.as_array()).cloned().unwrap_or_default();
                    xll.call(name, &args)
                }
                _ => serde_json::json!({"error": "unknown op"}),
            };
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
    // Chrome and Edge start a native-messaging host with the calling
    // extension's origin as its first argument.
    if args.get(1).is_some_and(|a| a.starts_with("chrome-extension://") || a == "--native-messaging") {
        serve_native_messaging();
        return;
    }
    if args.len() < 3 {
        eprintln!("usage: oxi-xll-host <add-in.xll> list | call <Function> '<json args>'");
        std::process::exit(2);
    }
    let xll = match open_xll(&args[1]) {
        Ok(xll) => xll,
        Err(error) => {
            eprintln!("{error}");
            std::process::exit(1);
        }
    };
    match args[2].as_str() {
        "list" => println!("{}", xll.list()),
        "call" => {
            let name = args.get(3).cloned().unwrap_or_default();
            let given: Vec<serde_json::Value> = serde_json::from_str(args.get(4).map(String::as_str).unwrap_or("[]")).unwrap_or_default();
            println!("{}", xll.call(&name, &given));
        }
        other => {
            eprintln!("unknown command {other}");
            std::process::exit(2);
        }
    }
}

#[cfg(not(windows))]
fn main() {
    eprintln!("XLL add-ins are Windows DLLs; this host runs on Windows only.");
    std::process::exit(1);
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn native_messages_are_length_prefixed_json() {
        let mut wire = Vec::new();
        write_message(&mut wire, &serde_json::json!({"id": 7, "op": "list"})).unwrap();
        assert_eq!(u32::from_le_bytes(wire[..4].try_into().unwrap()) as usize, wire.len() - 4);
        let back = read_message(&mut std::io::Cursor::new(wire)).unwrap();
        assert_eq!(back, serde_json::json!({"id": 7, "op": "list"}));
        assert!(read_message(&mut std::io::Cursor::new(vec![1u8, 0])).is_none());
    }

    #[test]
    fn type_strings_split_into_codes() {
        assert_eq!(type_codes("BBB"), vec!["B", "B", "B"]);
        assert_eq!(type_codes("QC%J!"), vec!["Q", "C%", "J"]);
        assert_eq!(type_codes("D%D%$"), vec!["D%", "D%"]);
    }

    #[cfg(all(windows, target_arch = "x86_64"))]
    #[test]
    fn win64_calls_reach_mixed_and_stacked_arguments() {
        extern "system" fn mixed(n: i32, x: f64, p: *const f64, y: f64) -> f64 {
            f64::from(n) * x + unsafe { *p } - y
        }
        extern "system" fn six(a: f64, b: f64, c: f64, d: f64, e: f64, f: f64) -> f64 {
            a + 2.0 * b + 3.0 * c + 4.0 * d + 5.0 * e + 6.0 * f
        }
        extern "system" fn seven_ints(a: i64, b: i64, c: i64, d: i64, e: i64, f: i64, g: i64) -> i64 {
            a + b * 10 + c * 100 + d * 1000 + e * 10000 + f * 100000 + g * 1000000
        }
        let k = 7.5f64;
        let (_, r) = unsafe { call_win64(mixed as usize, &[3, 2.0f64.to_bits(), &k as *const f64 as u64, 1.0f64.to_bits()]) };
        assert_eq!(r, 3.0 * 2.0 + 7.5 - 1.0);
        let six_args: Vec<u64> = (1..=6).map(|v| (v as f64).to_bits()).collect();
        let (_, r) = unsafe { call_win64(six as usize, &six_args) };
        assert_eq!(r, 1.0 + 4.0 + 9.0 + 16.0 + 25.0 + 36.0);
        let (r, _) = unsafe { call_win64(seven_ints as usize, &[1, 2, 3, 4, 5, 6, 7]) };
        assert_eq!(r as i64, 7654321);
    }
}
