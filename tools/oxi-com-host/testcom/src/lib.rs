// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! A COM Automation add-in: one class (`Oxi.TestCom`) whose IDispatch
//! methods Excel can call from a worksheet once the class is added as an
//! Automation add-in, and oxi-com-host can call outside Excel.
//!
//! Add(a, b), Upper(s), Sum(x), Shape(x), Twice(x): x is a number, an
//! array, or a Range object (Excel hands a Range for a reference; its Value
//! is read through IDispatch).

#![allow(non_snake_case, clippy::missing_safety_doc)]

use std::ffi::c_void;
use windows::core::*;
use windows::Win32::Foundation::*;
use windows::Win32::System::Com::*;
use windows::Win32::System::Ole::*;
use windows::Win32::System::Variant::*;

/// What the class was asked, appended to %TEMP%\oxi_testcom.log: which
/// interfaces and names a caller (Excel) wanted, for finding why a call fails.
fn trace(what: String) {
    use std::io::Write;
    let path = std::env::temp_dir().join("oxi_testcom.log");
    if let Ok(mut f) = std::fs::OpenOptions::new().create(true).append(true).open(path) {
        let _ = writeln!(f, "{what}");
    }
}

pub const CLSID_TEST: GUID = GUID::from_u128(0x6f1c2b7e_3a55_4e2b_9c1d_0a7e5b4c3d21);
const NAMES: [&str; 5] = ["Add", "Upper", "Sum", "Shape", "Twice"];

const VT_BYREF_: u16 = 0x4000;
const VT_ARRAY_: u16 = 0x2000;

/// A VARIANT's type and its first eight bytes of value (x64 layout).
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

/// A VARIANT as rows of values (a scalar is one row of one), each element
/// itself a VARIANT; a Range (or any object) is read through its Value.
unsafe fn grid(v: &VARIANT) -> Vec<Vec<VARIANT>> {
    let (vt, data) = raw(v);
    if vt & VT_BYREF_ != 0 && vt & !VT_BYREF_ == VT_VARIANT.0 {
        return grid(&*(data as *const VARIANT));
    }
    if vt == VT_DISPATCH.0 && data != 0 {
        let object: &IDispatch = std::mem::transmute(&data);
        let name = HSTRING::from("Value");
        let names = [PCWSTR(name.as_ptr())];
        let mut id = 0i32;
        if object.GetIDsOfNames(&GUID::zeroed(), names.as_ptr(), 1, 0, &mut id).is_err() {
            return vec![vec![VARIANT::default()]];
        }
        let params = DISPPARAMS::default();
        let mut value = VARIANT::default();
        if object.Invoke(id, &GUID::zeroed(), 0, DISPATCH_PROPERTYGET, &params, Some(&mut value), None, None).is_err() {
            return vec![vec![VARIANT::default()]];
        }
        return grid(&value);
    }
    if vt & VT_ARRAY_ != 0 {
        let array = data as *const SAFEARRAY;
        let dims = SafeArrayGetDim(array);
        let bound = |d: u32| (SafeArrayGetLBound(array, d).unwrap_or(0), SafeArrayGetUBound(array, d).unwrap_or(-1));
        let (r0, r1) = bound(1);
        let (c0, c1) = if dims >= 2 { bound(2) } else { (0, 0) };
        let mut out = Vec::new();
        for r in r0..=r1 {
            let mut row = Vec::new();
            for c in c0..=c1 {
                let index = if dims >= 2 { vec![r, c] } else { vec![r] };
                let mut cell = VARIANT::default();
                let _ = SafeArrayGetElement(array, index.as_ptr(), &mut cell as *mut VARIANT as *mut c_void);
                row.push(cell);
            }
            out.push(row);
        }
        // A 1-D array is one row.
        if dims < 2 {
            return vec![out.into_iter().flatten().collect()];
        }
        return out;
    }
    vec![vec![v.clone()]]
}

fn number(v: &VARIANT) -> Option<f64> {
    let (vt, _) = raw(v);
    if vt == VT_EMPTY.0 || vt == VT_BSTR.0 || vt == VT_ERROR.0 {
        return None;
    }
    let mut out = VARIANT::default();
    unsafe { VariantChangeType(&mut out, v, VAR_CHANGE_FLAGS(0), VT_R8).ok()? };
    f64::try_from(&out).ok()
}

/// Type information for the five methods: Excel lists an Automation
/// add-in's functions from its ITypeInfo (with none, every call is #NAME?).
fn type_info() -> Option<ITypeInfo> {
    thread_local! {
        static INFO: Option<ITypeInfo> = unsafe {
            let arity = [2u32, 1, 1, 1, 1];
            let wide = |s: &str| -> PWSTR {
                let units: Vec<u16> = s.encode_utf16().chain(std::iter::once(0)).collect();
                PWSTR(Box::leak(units.into_boxed_slice()).as_mut_ptr())
            };
            let methods: Vec<METHODDATA> = NAMES
                .iter()
                .enumerate()
                .map(|(i, name)| {
                    let params: Vec<PARAMDATA> = (0..arity[i]).map(|k| PARAMDATA { szName: wide(&format!("arg{}", k + 1)), vt: VT_VARIANT }).collect();
                    METHODDATA {
                        szName: wide(name),
                        ppdata: Box::leak(params.into_boxed_slice()).as_mut_ptr(),
                        dispid: i as i32 + 1,
                        iMeth: 7 + i as u32,
                        cc: CC_STDCALL,
                        cArgs: arity[i],
                        wFlags: DISPATCH_METHOD.0,
                        vtReturn: VT_VARIANT,
                    }
                })
                .collect();
            let mut data = INTERFACEDATA { pmethdata: Box::leak(methods.into_boxed_slice()).as_mut_ptr(), cMembers: NAMES.len() as u32 };
            let mut info: Option<ITypeInfo> = None;
            CreateDispTypeInfo(&mut data, 0, &mut info).ok().and(info)
        };
    }
    INFO.with(Clone::clone)
}

/// The class's default interface as its type library names it: no methods
/// of its own beyond IDispatch, but asked for by its IID when Excel loads the
/// add-in (answering only IID_IDispatch leaves every function #NAME?).
#[windows::core::interface("6F1C2B7E-3A55-4E2B-9C1D-0A7E5B4C3D23")]
unsafe trait ITestCom: IDispatch {}

#[implement(ITestCom, IDispatch)]
struct TestCom;

impl ITestCom_Impl for TestCom_Impl {}

impl IDispatch_Impl for TestCom_Impl {
    fn GetTypeInfoCount(&self) -> Result<u32> {
        trace("GetTypeInfoCount".into());
        Ok(u32::from(type_info().is_some()))
    }
    fn GetTypeInfo(&self, _: u32, _: u32) -> Result<ITypeInfo> {
        // The registered library's own description of the interface when
        // there is one (Excel looks a name up in it before it calls); the
        // built one otherwise.
        let registered = unsafe { LoadRegTypeLib(&LIBID_TEST, 1, 0, 0).and_then(|lib| lib.GetTypeInfoOfGuid(&IID_TEST)) };
        trace(format!("GetTypeInfo registered={}", registered.is_ok()));
        registered.or_else(|_| type_info().ok_or_else(|| E_NOTIMPL.into()))
    }
    fn GetIDsOfNames(&self, _: *const GUID, names: *const PCWSTR, count: u32, _: u32, ids: *mut i32) -> Result<()> {
        let mut known = true;
        for i in 0..count as usize {
            let name = unsafe { (*names.add(i)).to_string().unwrap_or_default() };
            trace(format!("GetIDsOfNames {name} (of {count})"));
            let id = NAMES.iter().position(|n| n.eq_ignore_ascii_case(&name)).map_or(-1, |p| p as i32 + 1);
            known &= id > 0;
            unsafe { *ids.add(i) = id };
        }
        if known { Ok(()) } else { Err(DISP_E_UNKNOWNNAME.into()) }
    }
    fn Invoke(&self, id: i32, _: *const GUID, _: u32, flags: DISPATCH_FLAGS, params: *const DISPPARAMS, result: *mut VARIANT, _: *mut EXCEPINFO, _: *mut u32) -> Result<()> {
        trace(format!("Invoke {id} flags {} args {}", flags.0, unsafe { (*params).cArgs }));
        // Arguments come last first.
        let args: Vec<VARIANT> = unsafe {
            let p = &*params;
            (0..p.cArgs as usize).rev().map(|i| (*p.rgvarg.add(i)).clone()).collect()
        };
        let arg = |i: usize| args.get(i).cloned().unwrap_or_default();
        let answer: VARIANT = match id {
            1 => VARIANT::from(number(&arg(0)).unwrap_or(0.0) + number(&arg(1)).unwrap_or(0.0)),
            2 => {
                let text = BSTR::try_from(&arg(0)).map(|s| s.to_string()).unwrap_or_default();
                VARIANT::from(BSTR::from(text.to_uppercase()))
            }
            3 => {
                let cells = unsafe { grid(&arg(0)) };
                VARIANT::from(cells.iter().flatten().filter_map(number).sum::<f64>())
            }
            4 => {
                let cells = unsafe { grid(&arg(0)) };
                VARIANT::from(BSTR::from(format!("{}x{}", cells.len(), cells.first().map_or(0, Vec::len))))
            }
            5 => unsafe {
                let cells = grid(&arg(0));
                let (rows, cols) = (cells.len() as u32, cells.first().map_or(0, Vec::len) as u32);
                let bounds = [SAFEARRAYBOUND { cElements: rows, lLbound: 1 }, SAFEARRAYBOUND { cElements: cols, lLbound: 1 }];
                let array = SafeArrayCreate(VT_VARIANT, 2, bounds.as_ptr());
                for (r, row) in cells.iter().enumerate() {
                    for (c, cell) in row.iter().enumerate() {
                        let value = VARIANT::from(number(cell).unwrap_or(0.0) * 2.0);
                        let index = [r as i32 + 1, c as i32 + 1];
                        let _ = SafeArrayPutElement(array, index.as_ptr(), &value as *const VARIANT as *const c_void);
                    }
                }
                make(VT_ARRAY_ | VT_VARIANT.0, array as u64)
            },
            _ => return Err(DISP_E_MEMBERNOTFOUND.into()),
        };
        if !result.is_null() {
            unsafe { std::ptr::write(result, answer) };
        }
        Ok(())
    }
}

#[implement(IClassFactory)]
struct Factory;

impl IClassFactory_Impl for Factory_Impl {
    fn CreateInstance(&self, outer: Option<&IUnknown>, riid: *const GUID, object: *mut *mut c_void) -> Result<()> {
        if outer.is_some() {
            return Err(CLASS_E_NOAGGREGATION.into());
        }
        let instance: IDispatch = TestCom.into();
        let found = unsafe { instance.query(riid, object) };
        trace(format!("CreateInstance {:?} -> {:?}", unsafe { *riid }, found));
        found.ok()
    }
    fn LockServer(&self, _: BOOL) -> Result<()> {
        Ok(())
    }
}

#[no_mangle]
pub unsafe extern "system" fn DllGetClassObject(class: *const GUID, riid: *const GUID, object: *mut *mut c_void) -> HRESULT {
    trace(format!("DllGetClassObject {:?} {:?}", *class, *riid));
    if *class != CLSID_TEST {
        return CLASS_E_CLASSNOTAVAILABLE;
    }
    let factory: IClassFactory = Factory.into();
    factory.query(riid, object)
}

#[no_mangle]
pub extern "system" fn DllCanUnloadNow() -> HRESULT {
    S_FALSE
}

pub const LIBID_TEST: GUID = GUID::from_u128(0x6f1c2b7e_3a55_4e2b_9c1d_0a7e5b4c3d22);
pub const IID_TEST: GUID = GUID::from_u128(0x6f1c2b7e_3a55_4e2b_9c1d_0a7e5b4c3d23);

/// Writes a type library for the class at `path` and registers it for this
/// user: what Excel reads to know an Automation add-in's functions.
#[no_mangle]
pub unsafe extern "system" fn OxiMakeTypeLib(path: PCWSTR) -> HRESULT {
    match make_type_lib(path) {
        Ok(()) => S_OK,
        Err(error) => error.code(),
    }
}

#[no_mangle]
pub unsafe extern "system" fn OxiDropTypeLib() -> HRESULT {
    match UnRegisterTypeLibForUser(&LIBID_TEST, 1, 0, 0, SYS_WIN64) {
        Ok(()) => S_OK,
        Err(error) => error.code(),
    }
}

unsafe fn make_type_lib(path: PCWSTR) -> Result<()> {
    let arity = [2u32, 1, 1, 1, 1];
    let library = CreateTypeLib2(SYS_WIN64, path)?;
    library.SetGuid(&LIBID_TEST)?;
    library.SetName(w!("OxiTestCom"))?;
    library.SetVersion(1, 0)?;
    library.SetLcid(0)?;
    let iface = library.CreateTypeInfo(w!("ITestCom"), TKIND_DISPATCH)?;
    iface.SetGuid(&IID_TEST)?;
    let standard = LoadTypeLibEx(w!("stdole2.tlb"), REGKIND_NONE)?;
    let dispatch = standard.GetTypeInfoOfGuid(&IDispatch::IID)?;
    let mut base = 0u32;
    iface.AddRefTypeInfo(&dispatch, &mut base)?;
    iface.AddImplType(0, base)?;
    for (i, name) in NAMES.iter().enumerate() {
        let n = arity[i] as usize;
        let mut params: Vec<ELEMDESC> = (0..n)
            .map(|_| {
                let mut e: ELEMDESC = std::mem::zeroed();
                e.tdesc.vt = VT_VARIANT;
                e
            })
            .collect();
        let mut function: FUNCDESC = std::mem::zeroed();
        function.memid = i as i32 + 1;
        function.lprgelemdescParam = params.as_mut_ptr();
        function.funckind = FUNC_DISPATCH;
        function.invkind = INVOKE_FUNC;
        function.callconv = CC_STDCALL;
        function.cParams = n as i16;
        function.elemdescFunc.tdesc.vt = VT_VARIANT;
        iface.AddFuncDesc(i as u32, &function)?;
        let words: Vec<HSTRING> = std::iter::once(HSTRING::from(*name)).chain((1..=n).map(|k| HSTRING::from(format!("arg{k}")))).collect();
        let pointers: Vec<PCWSTR> = words.iter().map(|w| PCWSTR(w.as_ptr())).collect();
        iface.SetFuncAndParamNames(i as u32, &pointers)?;
    }
    iface.LayOut()?;
    let class = library.CreateTypeInfo(w!("TestCom"), TKIND_COCLASS)?;
    class.SetGuid(&CLSID_TEST)?;
    let described: ITypeInfo = iface.cast()?;
    let mut reference = 0u32;
    class.AddRefTypeInfo(&described, &mut reference)?;
    class.AddImplType(0, reference)?;
    class.SetImplTypeFlags(0, IMPLTYPEFLAG_FDEFAULT)?;
    class.LayOut()?;
    library.SaveAllChanges()?;
    let loaded = LoadTypeLibEx(path, REGKIND_NONE)?;
    RegisterTypeLibForUser(&loaded, path, PCWSTR::null())?;
    Ok(())
}
