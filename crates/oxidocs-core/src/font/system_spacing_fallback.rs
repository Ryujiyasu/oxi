// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Platform font mapping for missing spacing glyphs. Font coverage is checked
//! by the caller; this module does not encode a family or codepoint priority.
//! Native Word controls confirm the system mapping for missing spacing glyphs.
//! Other scripts retain their separately measured document/theme rules.

#[derive(Clone)]
pub(crate) struct Mapping {
    pub family: String,
    pub scale: f32,
}

pub(crate) fn select(base: &str, ch: char, bold: bool, italic: bool, locale: &str) -> Option<Mapping> {
    #[cfg(windows)]
    {
        use std::{collections::HashMap, sync::{OnceLock, RwLock}};
        type Key = (String, char, bool, bool, String);
        static CACHE: OnceLock<RwLock<HashMap<Key, Option<Mapping>>>> = OnceLock::new();
        let cache = CACHE.get_or_init(|| RwLock::new(HashMap::new()));
        let key = (base.to_owned(), ch, bold, italic, locale.to_owned());
        if let Some(mapping) = cache.read().ok()?.get(&key) { return mapping.clone(); }
        let mapping = unsafe { native::map(base, ch, bold, italic, locale) };
        cache.write().ok()?.insert(key, mapping.clone());
        mapping
    }
    #[cfg(not(windows))]
    { let _ = (base, ch, bold, italic, locale); None }
}

#[cfg(windows)]
mod native {
    use super::Mapping;
    use std::{ffi::c_void, ptr, sync::atomic::{AtomicU32, Ordering}};

    #[repr(C)]
    #[derive(PartialEq)]
    struct Guid { a: u32, b: u16, c: u16, d: [u8; 8] }
    const FACTORY2: Guid = Guid { a: 0x0439fc60, b: 0xca44, c: 0x4994, d: [0x8d,0xee,0x3a,0x9a,0xf7,0xb7,0x32,0xec] };
    const UNKNOWN: Guid = Guid { a: 0, b: 0, c: 0, d: [0xc0,0,0,0,0,0,0,0x46] };
    const ANALYSIS_SOURCE: Guid = Guid { a: 0x688e1a58, b: 0x5094, c: 0x47c8, d: [0xad,0xc8,0xfb,0xce,0xa6,0x0a,0xe9,0x2b] };
    #[link(name="dwrite")]
    extern "system" { fn DWriteCreateFactory(kind: u32, iid: *const Guid, out: *mut *mut c_void) -> i32; }

    struct Com(*mut c_void);
    impl Com {
        unsafe fn slot(&self, index: usize) -> *const c_void {
            let vtable = *(self.0 as *const *const *const c_void);
            *vtable.add(index)
        }
    }
    impl Drop for Com {
        fn drop(&mut self) {
            if !self.0.is_null() { unsafe {
                let release: unsafe extern "system" fn(*mut c_void)->u32 = std::mem::transmute(self.slot(2));
                release(self.0);
            }}
        }
    }

    #[repr(C)]
    struct Source { vtable: *const SourceVtable, text: Vec<u16>, locale: Vec<u16>, refs: AtomicU32 }
    #[repr(C)]
    struct SourceVtable {
        query: unsafe extern "system" fn(*mut Source,*const Guid,*mut *mut c_void)->i32,
        add: unsafe extern "system" fn(*mut Source)->u32,
        release: unsafe extern "system" fn(*mut Source)->u32,
        at: unsafe extern "system" fn(*mut Source,u32,*mut *const u16,*mut u32)->i32,
        before: unsafe extern "system" fn(*mut Source,u32,*mut *const u16,*mut u32)->i32,
        direction: unsafe extern "system" fn(*mut Source)->u32,
        locale: unsafe extern "system" fn(*mut Source,u32,*mut u32,*mut *const u16)->i32,
        substitution: unsafe extern "system" fn(*mut Source,u32,*mut u32,*mut *mut c_void)->i32,
    }
    unsafe extern "system" fn query(this:*mut Source,iid:*const Guid,out:*mut *mut c_void)->i32 {
        if iid.is_null() || out.is_null() { return 0x80004003u32 as i32; }
        *out = ptr::null_mut();
        if *iid != UNKNOWN && *iid != ANALYSIS_SOURCE { return 0x80004002u32 as i32; }
        *out = this.cast(); add(this); 0
    }
    unsafe extern "system" fn add(this:*mut Source)->u32 { (*this).refs.fetch_add(1,Ordering::Relaxed)+1 }
    unsafe extern "system" fn release(this:*mut Source)->u32 { (*this).refs.fetch_sub(1,Ordering::Relaxed)-1 }
    unsafe extern "system" fn at(this:*mut Source,pos:u32,out:*mut *const u16,len:*mut u32)->i32 {
        let source=&*this; let count=source.text.len()-1;
        if pos as usize >= count { *out=ptr::null(); *len=0; }
        else { *out=source.text.as_ptr().add(pos as usize); *len=(count-pos as usize) as u32; }
        0
    }
    unsafe extern "system" fn before(this:*mut Source,pos:u32,out:*mut *const u16,len:*mut u32)->i32 {
        let source=&*this; let count=source.text.len()-1;
        if pos==0 || pos as usize>count { *out=ptr::null(); *len=0; }
        else { *out=source.text.as_ptr(); *len=pos; }
        0
    }
    unsafe extern "system" fn direction(_: *mut Source)->u32 { 0 }
    unsafe extern "system" fn locale(this:*mut Source,pos:u32,len:*mut u32,out:*mut *const u16)->i32 {
        let source=&*this; *len=(source.text.len()-1).saturating_sub(pos as usize) as u32;
        *out=source.locale.as_ptr(); 0
    }
    unsafe extern "system" fn substitution(this:*mut Source,pos:u32,len:*mut u32,out:*mut *mut c_void)->i32 {
        *len=((*this).text.len()-1).saturating_sub(pos as usize) as u32; *out=ptr::null_mut(); 0
    }
    static VTABLE: SourceVtable = SourceVtable { query,add,release,at,before,direction,locale,substitution };
    fn wide(text:&str)->Vec<u16> { text.encode_utf16().chain(Some(0)).collect() }

    pub(super) unsafe fn map(base:&str,ch:char,bold:bool,italic:bool,language:&str)->Option<Mapping> {
        let mut factory=Com(ptr::null_mut());
        if DWriteCreateFactory(0,&FACTORY2,&mut factory.0)<0 { return None; }
        let get_fallback: unsafe extern "system" fn(*mut c_void,*mut *mut c_void)->i32 = std::mem::transmute(factory.slot(26));
        let mut fallback=Com(ptr::null_mut());
        if get_fallback(factory.0,&mut fallback.0)<0 || fallback.0.is_null() { return None; }
        let mut source=Source { vtable:&VTABLE,text:wide(&ch.to_string()),locale:wide(language),refs:AtomicU32::new(1) };
        let base=wide(base); let mut length=0; let mut scale=0.0; let mut font=Com(ptr::null_mut());
        type Map = unsafe extern "system" fn(*mut c_void,*mut Source,u32,u32,*mut c_void,*const u16,u32,u32,u32,*mut u32,*mut *mut c_void,*mut f32)->i32;
        let map:Map=std::mem::transmute(fallback.slot(3));
        let text_length=(source.text.len()-1) as u32;
        if map(fallback.0,&mut source,0,text_length,ptr::null_mut(),base.as_ptr(),if bold {700}else{400},if italic{2}else{0},5,&mut length,&mut font.0,&mut scale)<0
            || font.0.is_null() || length==0 || !scale.is_finite() || scale<=0.0 { return None; }
        type GetObject=unsafe extern "system" fn(*mut c_void,*mut *mut c_void)->i32;
        let get_family:GetObject=std::mem::transmute(font.slot(3)); let mut family=Com(ptr::null_mut());
        if get_family(font.0,&mut family.0)<0 || family.0.is_null() {return None;}
        let get_names:GetObject=std::mem::transmute(family.slot(6));let mut names=Com(ptr::null_mut());
        if get_names(family.0,&mut names.0)<0 || names.0.is_null(){return None;}
        let find:unsafe extern "system" fn(*mut c_void,*const u16,*mut u32,*mut i32)->i32=std::mem::transmute(names.slot(4));
        let mut index=0;let mut exists=0;let english=wide("en-us");
        if find(names.0,english.as_ptr(),&mut index,&mut exists)<0{return None;}
        if exists==0 {index=0;}
        let get_length:unsafe extern "system" fn(*mut c_void,u32,*mut u32)->i32=std::mem::transmute(names.slot(7));
        let mut name_length=0;if get_length(names.0,index,&mut name_length)<0 || name_length>4096{return None;}
        let mut name=vec![0u16;name_length as usize+1];
        let get_string:unsafe extern "system" fn(*mut c_void,u32,*mut u16,u32)->i32=std::mem::transmute(names.slot(8));
        if get_string(names.0,index,name.as_mut_ptr(),name.len() as u32)<0{return None;}
        Some(Mapping {family:String::from_utf16(&name[..name_length as usize]).ok()?,scale})
    }

    #[cfg(test)]
    mod tests {
        use super::*;
        fn source()->Source { Source { vtable:&VTABLE,text:wide("A\u{1f600}"),locale:wide("en-US"),refs:AtomicU32::new(1) } }
        #[test]
        fn text_source_returns_utf16_units_and_null_after_end() {
            let mut source=source();let mut pointer=ptr::null();let mut length=0;
            unsafe {
                at(&mut source,1,&mut pointer,&mut length);
                assert_eq!(length,2);assert_eq!(*pointer,0xd83d);
                at(&mut source,2,&mut pointer,&mut length);
                assert_eq!(length,1);assert_eq!(*pointer,0xde00);
                at(&mut source,3,&mut pointer,&mut length);
                assert_eq!(length,0);assert!(pointer.is_null());
                before(&mut source,3,&mut pointer,&mut length);
                assert_eq!(length,3);assert_eq!(*pointer,65);
                before(&mut source,4,&mut pointer,&mut length);
                assert_eq!(length,0);assert!(pointer.is_null());
            }
        }
        #[test]
        fn text_source_refuses_extended_interfaces_without_refcount_change() {
            let mut source=source();let unsupported=Guid {a:1,b:0,c:0,d:[0;8]};let mut result=ptr::null_mut();
            unsafe {
                assert_eq!(query(&mut source,&unsupported,&mut result),0x80004002u32 as i32);
                assert!(result.is_null());assert_eq!(source.refs.load(Ordering::Relaxed),1);
                assert_eq!(query(&mut source,&ANALYSIS_SOURCE,&mut result),0);
                assert_eq!(result,(&mut source as *mut Source).cast());
                assert_eq!(release(&mut source),1);
            }
        }
    }

}
