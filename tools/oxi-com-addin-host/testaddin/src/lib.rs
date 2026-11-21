// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! A COM add-in (`Oxi.TestAddin`): on connection it works the workbook
//! through the Application object it is handed, as add-ins do -- writes
//! text, a formula, a value computed from the formula's result, the
//! application's name and the connect mode into A1:E1 of the active sheet.

#![allow(non_snake_case, clippy::missing_safety_doc)]

use std::ffi::c_void;
use windows::core::*;
use windows::Win32::Foundation::*;
use windows::Win32::System::Com::*;
use windows::Win32::System::Variant::*;

pub const CLSID_ADDIN: GUID = GUID::from_u128(0x6f1c2b7e_3a55_4e2b_9c1d_0a7e5b4c3d31);

/// Office's add-in interface (msaddndr.dll), dual.
#[windows::core::interface("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")]
pub unsafe trait IDTExtensibility2: IDispatch {
    fn OnConnection(&self, application: *mut c_void, mode: i32, addin: *mut c_void, custom: *mut *mut SAFEARRAY) -> HRESULT;
    fn OnDisconnection(&self, mode: i32, custom: *mut *mut SAFEARRAY) -> HRESULT;
    fn OnAddInsUpdate(&self, custom: *mut *mut SAFEARRAY) -> HRESULT;
    fn OnStartupComplete(&self, custom: *mut *mut SAFEARRAY) -> HRESULT;
    fn OnBeginShutdown(&self, custom: *mut *mut SAFEARRAY) -> HRESULT;
}

fn raw(v: &VARIANT) -> (u16, u64) {
    unsafe {
        let p = v as *const VARIANT as *const u8;
        (*(p as *const u16), *(p.add(8) as *const u64))
    }
}

/// `object.name(args)` through IDispatch, with the given flags.
unsafe fn invoke(object: &IDispatch, name: &str, flags: DISPATCH_FLAGS, args: &[VARIANT]) -> Result<VARIANT> {
    let wide = HSTRING::from(name);
    let names = [PCWSTR(wide.as_ptr())];
    let mut id = 0i32;
    object.GetIDsOfNames(&GUID::zeroed(), names.as_ptr(), 1, 0, &mut id)?;
    let mut given: Vec<VARIANT> = args.iter().rev().cloned().collect();
    let mut put = -3i32; // DISPID_PROPERTYPUT
    let params = DISPPARAMS {
        rgvarg: given.as_mut_ptr(),
        rgdispidNamedArgs: if flags == DISPATCH_PROPERTYPUT { &mut put } else { std::ptr::null_mut() },
        cArgs: given.len() as u32,
        cNamedArgs: u32::from(flags == DISPATCH_PROPERTYPUT),
    };
    let mut result = VARIANT::default();
    object.Invoke(id, &GUID::zeroed(), 0, flags, &params, Some(&mut result), None, None)?;
    Ok(result)
}

unsafe fn object(v: &VARIANT) -> Result<IDispatch> {
    let (vt, data) = raw(v);
    if vt != VT_DISPATCH.0 || data == 0 {
        return Err(E_NOINTERFACE.into());
    }
    let pointer = data as *mut c_void;
    IDispatch::from_raw_borrowed(&pointer).cloned().ok_or_else(|| E_POINTER.into())
}

unsafe fn work(app: &IDispatch, mode: i32) -> Result<()> {
    let sheet = object(&invoke(app, "ActiveSheet", DISPATCH_PROPERTYGET, &[])?)?;
    let range = |a: &str| -> Result<IDispatch> { object(&invoke(&sheet, "Range", DISPATCH_PROPERTYGET | DISPATCH_METHOD, &[VARIANT::from(BSTR::from(a))])?) };
    invoke(&range("A1")?, "Value", DISPATCH_PROPERTYPUT, &[VARIANT::from(BSTR::from("hello from add-in"))])?;
    let b1 = range("B1")?;
    invoke(&b1, "Formula", DISPATCH_PROPERTYPUT, &[VARIANT::from(BSTR::from("=1+2"))])?;
    let three = f64::try_from(&invoke(&b1, "Value", DISPATCH_PROPERTYGET, &[])?).unwrap_or(-1.0);
    invoke(&range("C1")?, "Value", DISPATCH_PROPERTYPUT, &[VARIANT::from(three * 10.0)])?;
    let name = invoke(app, "Name", DISPATCH_PROPERTYGET, &[])?;
    invoke(&range("D1")?, "Value", DISPATCH_PROPERTYPUT, &[name])?;
    invoke(&range("E1")?, "Value", DISPATCH_PROPERTYPUT, &[VARIANT::from(mode)])?;
    Ok(())
}

#[implement(IDTExtensibility2, IDispatch)]
struct AddIn;

impl IDTExtensibility2_Impl for AddIn_Impl {
    unsafe fn OnConnection(&self, application: *mut c_void, mode: i32, _: *mut c_void, _: *mut *mut SAFEARRAY) -> HRESULT {
        let Some(app) = IDispatch::from_raw_borrowed(&application) else { return E_POINTER };
        match work(app, mode) {
            Ok(()) => S_OK,
            Err(e) => e.code(),
        }
    }
    unsafe fn OnDisconnection(&self, _: i32, _: *mut *mut SAFEARRAY) -> HRESULT {
        S_OK
    }
    unsafe fn OnAddInsUpdate(&self, _: *mut *mut SAFEARRAY) -> HRESULT {
        S_OK
    }
    unsafe fn OnStartupComplete(&self, _: *mut *mut SAFEARRAY) -> HRESULT {
        S_OK
    }
    unsafe fn OnBeginShutdown(&self, _: *mut *mut SAFEARRAY) -> HRESULT {
        S_OK
    }
}

impl IDispatch_Impl for AddIn_Impl {
    fn GetTypeInfoCount(&self) -> Result<u32> {
        Ok(0)
    }
    fn GetTypeInfo(&self, _: u32, _: u32) -> Result<ITypeInfo> {
        Err(E_NOTIMPL.into())
    }
    fn GetIDsOfNames(&self, _: *const GUID, _: *const PCWSTR, _: u32, _: u32, _: *mut i32) -> Result<()> {
        Err(DISP_E_UNKNOWNNAME.into())
    }
    fn Invoke(&self, _: i32, _: *const GUID, _: u32, _: DISPATCH_FLAGS, _: *const DISPPARAMS, _: *mut VARIANT, _: *mut EXCEPINFO, _: *mut u32) -> Result<()> {
        Err(DISP_E_MEMBERNOTFOUND.into())
    }
}

#[implement(IClassFactory)]
struct Factory;

impl IClassFactory_Impl for Factory_Impl {
    fn CreateInstance(&self, outer: Option<&IUnknown>, riid: *const GUID, object: *mut *mut c_void) -> Result<()> {
        if outer.is_some() {
            return Err(CLASS_E_NOAGGREGATION.into());
        }
        let instance: IDTExtensibility2 = AddIn.into();
        unsafe { instance.query(riid, object).ok() }
    }
    fn LockServer(&self, _: BOOL) -> Result<()> {
        Ok(())
    }
}

#[no_mangle]
pub unsafe extern "system" fn DllGetClassObject(class: *const GUID, riid: *const GUID, object: *mut *mut c_void) -> HRESULT {
    if *class != CLSID_ADDIN {
        return CLASS_E_CLASSNOTAVAILABLE;
    }
    let factory: IClassFactory = Factory.into();
    factory.query(riid, object)
}

#[no_mangle]
pub extern "system" fn DllCanUnloadNow() -> HRESULT {
    S_FALSE
}
