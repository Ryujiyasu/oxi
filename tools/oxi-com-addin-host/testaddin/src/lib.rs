// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! A COM add-in (`Oxi.TestAddin`): on connection it works the workbook
//! through the Application object it is handed, as add-ins do -- writes
//! text, a formula, a value computed from the formula's result, the
//! application's name and the connect mode into A1:E1 of the active sheet --
//! and listens to the application's events: when A2 changes it writes the
//! changed cell's address into F1 and twice its value into G1.

#![allow(non_snake_case, clippy::missing_safety_doc)]

use std::ffi::c_void;
use windows::core::*;
use windows::Win32::Foundation::*;
use windows::Win32::System::Com::*;
use windows::Win32::System::Variant::*;

/// What the add-in saw, appended to %TEMP%\oxi_testaddin.log.
fn trace(what: String) {
    use std::io::Write;
    if let Ok(mut f) = std::fs::OpenOptions::new().create(true).append(true).open(std::env::temp_dir().join("oxi_testaddin.log")) {
        let _ = writeln!(f, "{what}");
    }
}

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

/// Excel's application events (AppEvents): Advise asks the sink for it.
#[windows::core::interface("00024413-0000-0000-C000-000000000046")]
unsafe trait IAppEvents: IDispatch {}

const DISPID_SHEET_CHANGE: i32 = 0x61C;

/// SheetChange events seen so far: the add-in's own writes are changes too.
static SEEN: std::sync::atomic::AtomicU32 = std::sync::atomic::AtomicU32::new(0);

#[implement(IAppEvents, IDispatch)]
struct Sink;

impl IAppEvents_Impl for Sink_Impl {}

impl IDispatch_Impl for Sink_Impl {
    fn GetTypeInfoCount(&self) -> Result<u32> {
        Ok(0)
    }
    fn GetTypeInfo(&self, _: u32, _: u32) -> Result<ITypeInfo> {
        Err(E_NOTIMPL.into())
    }
    fn GetIDsOfNames(&self, _: *const GUID, _: *const PCWSTR, _: u32, _: u32, _: *mut i32) -> Result<()> {
        Err(DISP_E_UNKNOWNNAME.into())
    }
    fn Invoke(&self, id: i32, _: *const GUID, _: u32, _: DISPATCH_FLAGS, params: *const DISPPARAMS, _: *mut VARIANT, _: *mut EXCEPINFO, _: *mut u32) -> Result<()> {
        unsafe {
            let p = &*params;
            let kinds: Vec<u16> = (0..p.cArgs as usize).map(|i| raw(&*p.rgvarg.add(i)).0).collect();
            trace(format!("event {id:#x} args {kinds:?}"));
        }
        if id != DISPID_SHEET_CHANGE {
            return Ok(());
        }
        SEEN.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
        unsafe {
            let p = &*params;
            if p.cArgs < 2 {
                return Ok(());
            }
            // Measured: Excel hands an event's arguments in their declared
            // order, not last first -- rgvarg[0] is Sh, rgvarg[1] is Target.
            let sheet = object(&*p.rgvarg)?;
            let target = object(&*p.rgvarg.add(1))?;
            let address = invoke(&target, "Address", DISPATCH_PROPERTYGET, &[]);
            trace(format!("Address -> {:?}", address.as_ref().map(|v| raw(v).0)));
            let address = address?;
            let text = BSTR::try_from(&address).map(|s| s.to_string()).unwrap_or_default();
            trace(format!("address text {text:?}"));
            // Only A2: the add-in's own writes below are changes too.
            if text != "$A$2" {
                return Ok(());
            }
            let value = f64::try_from(&invoke(&target, "Value", DISPATCH_PROPERTYGET, &[])?).unwrap_or(-1.0);
            let range = |a: &str| -> Result<IDispatch> { object(&invoke(&sheet, "Range", DISPATCH_PROPERTYGET | DISPATCH_METHOD, &[VARIANT::from(BSTR::from(a))])?) };
            let wrote = range("F1").and_then(|f1| invoke(&f1, "Value", DISPATCH_PROPERTYPUT, &[address]));
            trace(format!("F1 write {:?}", wrote.as_ref().map(|_| ())));
            wrote?;
            invoke(&range("G1")?, "Value", DISPATCH_PROPERTYPUT, &[VARIANT::from(value * 2.0)])?;
            // How many changes were told by now: A2's, F1's and G1's.
            let seen = f64::from(SEEN.load(std::sync::atomic::Ordering::SeqCst));
            invoke(&range("H1")?, "Value", DISPATCH_PROPERTYPUT, &[VARIANT::from(seen)])?;
        }
        Ok(())
    }
}

/// The sink's connection: kept so OnDisconnection can let it go.
static mut ADVISED: Option<(IConnectionPoint, u32)> = None;

unsafe fn listen(app: &IDispatch) -> Result<()> {
    let container: IConnectionPointContainer = app.cast()?;
    let point = container.FindConnectionPoint(&<IAppEvents as Interface>::IID)?;
    let sink: IAppEvents = Sink.into();
    let unknown: IUnknown = sink.cast()?;
    let cookie = point.Advise(&unknown)?;
    ADVISED = Some((point, cookie));
    Ok(())
}

/// Office's ribbon interface (mso.dll), dual: the add-in's own controls.
#[windows::core::interface("000C0396-0000-0000-C000-000000000046")]
pub unsafe trait IRibbonExtensibility: IDispatch {
    fn GetCustomUI(&self, ribbon_id: std::mem::ManuallyDrop<BSTR>, xml: *mut BSTR) -> HRESULT;
}

/// One button on a tab of its own; pressing it calls OnOxiButton.
const RIBBON: &str = r#"<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui"><ribbon><tabs><tab id="oxiTab" label="Oxi"><group id="oxiGroup" label="Test"><button id="oxiButton" label="Press" onAction="OnOxiButton"/></group></tab></tabs></ribbon></customUI>"#;

/// The application, kept from OnConnection for the ribbon callback.
static mut APP: Option<IDispatch> = None;

#[implement(IDTExtensibility2, IRibbonExtensibility, IDispatch)]
struct AddIn;

impl IRibbonExtensibility_Impl for AddIn_Impl {
    unsafe fn GetCustomUI(&self, ribbon_id: std::mem::ManuallyDrop<BSTR>, xml: *mut BSTR) -> HRESULT {
        trace(format!("GetCustomUI {}", ribbon_id.to_string()));
        std::ptr::write(xml, BSTR::from(RIBBON));
        S_OK
    }
}

impl IDTExtensibility2_Impl for AddIn_Impl {
    unsafe fn OnConnection(&self, application: *mut c_void, mode: i32, _: *mut c_void, _: *mut *mut SAFEARRAY) -> HRESULT {
        let Some(app) = IDispatch::from_raw_borrowed(&application) else { return E_POINTER };
        APP = Some(app.clone());
        let listened = work(app, mode).and_then(|()| listen(app));
        trace(format!("connected mode {mode}: {listened:?}"));
        match listened {
            Ok(()) => S_OK,
            Err(e) => e.code(),
        }
    }
    unsafe fn OnDisconnection(&self, _: i32, _: *mut *mut SAFEARRAY) -> HRESULT {
        if let Some((point, cookie)) = (*std::ptr::addr_of_mut!(ADVISED)).take() {
            let _ = point.Unadvise(cookie);
        }
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

/// The ribbon's callbacks by name: OnOxiButton(control) writes
/// "pressed <control.Id>" into I1.
impl IDispatch_Impl for AddIn_Impl {
    fn GetTypeInfoCount(&self) -> Result<u32> {
        Ok(0)
    }
    fn GetTypeInfo(&self, _: u32, _: u32) -> Result<ITypeInfo> {
        Err(E_NOTIMPL.into())
    }
    fn GetIDsOfNames(&self, _: *const GUID, names: *const PCWSTR, count: u32, _: u32, ids: *mut i32) -> Result<()> {
        let name = unsafe { (*names).to_string().unwrap_or_default() };
        trace(format!("callback name {name}"));
        if count == 1 && name.eq_ignore_ascii_case("OnOxiButton") {
            unsafe { *ids = 1 };
            Ok(())
        } else {
            Err(DISP_E_UNKNOWNNAME.into())
        }
    }
    fn Invoke(&self, id: i32, _: *const GUID, _: u32, _: DISPATCH_FLAGS, params: *const DISPPARAMS, _: *mut VARIANT, _: *mut EXCEPINFO, _: *mut u32) -> Result<()> {
        if id != 1 {
            return Err(DISP_E_MEMBERNOTFOUND.into());
        }
        unsafe {
            let p = &*params;
            let control = object(&*p.rgvarg.add(p.cArgs as usize - 1))?;
            let control_id = invoke(&control, "Id", DISPATCH_PROPERTYGET, &[])?;
            let text = format!("pressed {}", BSTR::try_from(&control_id).map(|s| s.to_string()).unwrap_or_default());
            let Some(app) = (*std::ptr::addr_of!(APP)).clone() else { return Err(E_FAIL.into()) };
            let sheet = object(&invoke(&app, "ActiveSheet", DISPATCH_PROPERTYGET, &[])?)?;
            let i1 = object(&invoke(&sheet, "Range", DISPATCH_PROPERTYGET | DISPATCH_METHOD, &[VARIANT::from(BSTR::from("I1"))])?)?;
            invoke(&i1, "Value", DISPATCH_PROPERTYPUT, &[VARIANT::from(BSTR::from(text))])?;
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
