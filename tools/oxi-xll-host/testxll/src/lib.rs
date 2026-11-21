// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! A test XLL built the way the Excel SDK builds one: `xlAutoOpen` registers
//! its functions through `Excel12v`, which finds the host's `MdCallBack12`
//! export exactly as the SDK's xlcall.cpp does. Excel loads it with
//! RegisterXLL and so does oxi-xll-host; the two are held against each other.

#![allow(non_snake_case, clippy::missing_safety_doc)]

use std::sync::OnceLock;

#[repr(C)]
#[derive(Clone, Copy)]
pub struct Xloper12 {
    pub val: [u64; 3],
    pub xltype: u32,
}

const XLTYPE_NUM: u32 = 0x0001;
const XLTYPE_STR: u32 = 0x0002;
const XLTYPE_ERR: u32 = 0x0010;
const XLTYPE_INT: u32 = 0x0800;
const XLTYPE_MULTI: u32 = 0x0040;
const XLBIT_DLL_FREE: u32 = 0x4000;
const XLF_REGISTER: i32 = 149;
const XLF_GET_NAME: i32 = 0x4009;
const XLERR_VALUE: u32 = 15;

type MdCallBack12 = unsafe extern "system" fn(i32, i32, *mut *mut Xloper12, *mut Xloper12) -> i32;

fn callback() -> Option<MdCallBack12> {
    static FOUND: OnceLock<Option<usize>> = OnceLock::new();
    let address = *FOUND.get_or_init(|| unsafe {
        use windows::core::{s, PCWSTR};
        use windows::Win32::System::LibraryLoader::{GetModuleHandleW, GetProcAddress};
        let module = GetModuleHandleW(PCWSTR::null()).ok()?;
        GetProcAddress(module, s!("MdCallBack12")).map(|f| f as usize)
    });
    address.map(|a| unsafe { std::mem::transmute::<usize, MdCallBack12>(a) })
}

unsafe fn excel12(xlfn: i32, result: *mut Xloper12, args: &mut [Xloper12]) -> i32 {
    let Some(f) = callback() else { return 32 };
    let mut pointers: Vec<*mut Xloper12> = args.iter_mut().map(|a| a as *mut Xloper12).collect();
    f(xlfn, pointers.len() as i32, pointers.as_mut_ptr(), result)
}

/// A counted wide string kept alive for the length of a call.
fn text(s: &str) -> (Vec<u16>, Xloper12) {
    let mut units: Vec<u16> = std::iter::once(s.encode_utf16().count() as u16).chain(s.encode_utf16()).collect();
    let x = Xloper12 { val: [units.as_mut_ptr() as u64, 0, 0], xltype: XLTYPE_STR };
    (units, x)
}

#[no_mangle]
pub unsafe extern "system" fn xlAutoOpen() -> i32 {
    // The module name, as SDK samples ask Excel for it.
    let mut name = Xloper12 { val: [0; 3], xltype: 0 };
    if excel12(XLF_GET_NAME, &mut name, &mut []) != 0 {
        return 0;
    }
    let registrations = [
        ("OxiAdd", "BBB", "OxiAdd"),
        ("OxiScale", "BJB", "OxiScale"),
        ("OxiUpper", "QQ", "OxiUpper"),
        ("OxiHypot4", "BBBBB", "OxiHypot4"),
        ("OxiSum6", "BBBBBBB", "OxiSum6"),
        ("OxiHalf", "QQ", "OxiHalf"),
        ("OxiRangeSum", "BQ", "OxiRangeSum"),
        ("OxiShape", "QQ", "OxiShape"),
        ("OxiFpSum", "BK%", "OxiFpSum"),
        ("OxiTranspose", "QQ", "OxiTranspose"),
        ("OxiFpTwice", "K%K%", "OxiFpTwice"),
    ];
    for (procedure, types, function) in registrations {
        let (_a, p) = text(procedure);
        let (_b, t) = text(types);
        let (_c, f) = text(function);
        let mut args = [name, p, t, f];
        let mut result = Xloper12 { val: [0; 3], xltype: 0 };
        excel12(XLF_REGISTER, &mut result, &mut args);
    }
    1
}

#[no_mangle]
pub extern "system" fn xlAutoClose() -> i32 {
    1
}

#[no_mangle]
pub extern "system" fn OxiAdd(a: f64, b: f64) -> f64 {
    a + b
}

#[no_mangle]
pub extern "system" fn OxiScale(n: i32, x: f64) -> f64 {
    f64::from(n) * x
}

#[no_mangle]
pub extern "system" fn OxiHypot4(a: f64, b: f64, c: f64, d: f64) -> f64 {
    (a * a + b * b + c * c + d * d).sqrt()
}

#[no_mangle]
pub extern "system" fn OxiSum6(a: f64, b: f64, c: f64, d: f64, e: f64, f: f64) -> f64 {
    a + 2.0 * b + 3.0 * c + 4.0 * d + 5.0 * e + 6.0 * f
}

/// Returned values are allocated here and handed back to xlAutoFree12.
fn boxed(x: Xloper12) -> *mut Xloper12 {
    Box::into_raw(Box::new(Xloper12 { xltype: x.xltype | XLBIT_DLL_FREE, ..x }))
}

#[no_mangle]
pub unsafe extern "system" fn OxiUpper(arg: *const Xloper12) -> *mut Xloper12 {
    let arg = &*arg;
    if arg.xltype & 0xFFF != XLTYPE_STR {
        return boxed(Xloper12 { val: [u64::from(XLERR_VALUE), 0, 0], xltype: XLTYPE_ERR });
    }
    let p = arg.val[0] as *const u16;
    let len = *p as usize;
    let s = String::from_utf16_lossy(std::slice::from_raw_parts(p.add(1), len)).to_uppercase();
    let units: Vec<u16> = std::iter::once(s.encode_utf16().count() as u16).chain(s.encode_utf16()).collect();
    let leaked = Box::leak(units.into_boxed_slice());
    boxed(Xloper12 { val: [leaked.as_ptr() as u64, 0, 0], xltype: XLTYPE_STR })
}

#[no_mangle]
pub unsafe extern "system" fn OxiHalf(arg: *const Xloper12) -> *mut Xloper12 {
    let arg = &*arg;
    match arg.xltype & 0xFFF {
        XLTYPE_NUM => boxed(Xloper12 { val: [(f64::from_bits(arg.val[0]) / 2.0).to_bits(), 0, 0], xltype: XLTYPE_NUM }),
        XLTYPE_INT => boxed(Xloper12 { val: [(f64::from(arg.val[0] as i32) / 2.0).to_bits(), 0, 0], xltype: XLTYPE_NUM }),
        _ => boxed(Xloper12 { val: [u64::from(XLERR_VALUE), 0, 0], xltype: XLTYPE_ERR }),
    }
}

/// The elements of a value argument: an array's, or the value alone.
unsafe fn elements(arg: &Xloper12) -> (usize, usize, Vec<Xloper12>) {
    if arg.xltype & 0xFFF == XLTYPE_MULTI {
        let (rows, cols) = (arg.val[1] as u32 as usize, (arg.val[1] >> 32) as u32 as usize);
        let p = arg.val[0] as *const Xloper12;
        (rows, cols, (0..rows * cols).map(|i| *p.add(i)).collect())
    } else {
        (1, 1, vec![*arg])
    }
}

/// The sum of the numbers in a range or array (text and blanks skipped).
#[no_mangle]
pub unsafe extern "system" fn OxiRangeSum(arg: *const Xloper12) -> f64 {
    let (_, _, all) = elements(&*arg);
    all.iter()
        .map(|x| match x.xltype & 0xFFF {
            XLTYPE_NUM => f64::from_bits(x.val[0]),
            XLTYPE_INT => f64::from(x.val[0] as i32),
            _ => 0.0,
        })
        .sum()
}

/// "rows x columns" of what arrived.
#[no_mangle]
pub unsafe extern "system" fn OxiShape(arg: *const Xloper12) -> *mut Xloper12 {
    let (rows, cols, _) = elements(&*arg);
    let s = format!("{rows}x{cols}");
    let units: Vec<u16> = std::iter::once(s.encode_utf16().count() as u16).chain(s.encode_utf16()).collect();
    let leaked = Box::leak(units.into_boxed_slice());
    boxed(Xloper12 { val: [leaked.as_ptr() as u64, 0, 0], xltype: XLTYPE_STR })
}

/// FP12: { rows, columns, doubles row by row }.
#[repr(C)]
pub struct Fp12 {
    pub rows: i32,
    pub columns: i32,
    pub array: [f64; 1],
}

#[no_mangle]
pub unsafe extern "system" fn OxiFpSum(fp: *const Fp12) -> f64 {
    let n = ((*fp).rows * (*fp).columns).max(0) as usize;
    std::slice::from_raw_parts((*fp).array.as_ptr(), n).iter().sum()
}

/// The array turned on its side; numbers only (anything else becomes 0).
#[no_mangle]
pub unsafe extern "system" fn OxiTranspose(arg: *const Xloper12) -> *mut Xloper12 {
    let (rows, cols, all) = elements(&*arg);
    let mut out: Vec<Xloper12> = Vec::with_capacity(rows * cols);
    for c in 0..cols {
        for r in 0..rows {
            let x = all[r * cols + c];
            let v = match x.xltype & 0xFFF {
                XLTYPE_NUM => f64::from_bits(x.val[0]),
                XLTYPE_INT => f64::from(x.val[0] as i32),
                _ => 0.0,
            };
            out.push(Xloper12 { val: [v.to_bits(), 0, 0], xltype: XLTYPE_NUM });
        }
    }
    let p = Box::leak(out.into_boxed_slice()).as_mut_ptr();
    boxed(Xloper12 { val: [p as u64, (rows as u64) << 32 | cols as u64, 0], xltype: XLTYPE_MULTI })
}

/// Each number doubled, in a buffer the add-in keeps (as XLL samples do).
#[no_mangle]
pub unsafe extern "system" fn OxiFpTwice(fp: *const Fp12) -> *mut Fp12 {
    static mut BUFFER: [f64; 4098] = [0.0; 4098];
    let (rows, cols) = ((*fp).rows, (*fp).columns);
    let n = (rows * cols).clamp(0, 4096) as usize;
    let out = std::ptr::addr_of_mut!(BUFFER) as *mut Fp12;
    (*out).rows = rows;
    (*out).columns = cols;
    let src = std::slice::from_raw_parts((*fp).array.as_ptr(), n);
    let dst = std::slice::from_raw_parts_mut((*out).array.as_mut_ptr(), n);
    for (d, s) in dst.iter_mut().zip(src) {
        *d = s * 2.0;
    }
    out
}

#[no_mangle]
pub unsafe extern "system" fn xlAutoFree12(x: *mut Xloper12) {
    if !x.is_null() {
        if (*x).xltype & 0xFFF == XLTYPE_MULTI {
            let n = ((*x).val[1] as u32 as usize) * (((*x).val[1] >> 32) as u32 as usize);
            drop(Box::from_raw(std::ptr::slice_from_raw_parts_mut((*x).val[0] as *mut Xloper12, n)));
        }
        drop(Box::from_raw(x));
    }
}
