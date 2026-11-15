// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Virtual BSTR allocation identity. These are interpreter addresses, never
//! native pointers. Empty allocated BSTRs and null BSTRs remain distinct.
use super::*;
use std::sync::atomic::{AtomicU64, Ordering};
static NEXT: AtomicU64 = AtomicU64::new(0x1_2000_0000);

/// Reserve disjoint interpreter storage, sharing the BSTR address space.
/// No returned integer is a native address or can be dereferenced by Rust.
pub(super) fn reserve_storage(bytes: u64) -> Result<i64, RuntimeError> {
    let size = bytes.checked_add(15).map(|value| (value & !15).max(32))
        .ok_or_else(|| error(RuntimeErrorKind::Overflow, "VBA address range overflow", None))?;
    let mut address = NEXT.load(Ordering::Relaxed);
    loop {
        let end = address.checked_add(size).filter(|end| *end <= i64::MAX as u64)
            .ok_or_else(|| error(RuntimeErrorKind::Overflow, "VBA address range overflow", None))?;
        match NEXT.compare_exchange_weak(address, end, Ordering::Relaxed, Ordering::Relaxed) {
            Ok(_) => return Ok(address as i64),
            Err(current) => address = current,
        }
    }
}
#[derive(Clone)]
pub(super) struct Allocation {
    pub(super) pointer: i64,
    pub(super) units: Vec<u16>,
}
fn allocate(units: Vec<u16>) -> Allocation {
    let size = ((units.len() as u64 * 2 + 31) & !15).max(32);
    Allocation {
        pointer: NEXT.fetch_add(size, Ordering::Relaxed) as i64 + 8,
        units,
    }
}
pub(super) fn slot_pointer(slot: &ValueSlot) -> Result<i64, RuntimeError> {
    let value = slot.snapshot();
    if matches!(value, Value::NullString) {
        return Ok(0);
    }
    let Some(units) = value.string_units() else {
        return value_pointer(value);
    };
    let units = units.into_owned();
    let mut strings = slot.storage.strings.borrow_mut();
    let entry = strings
        .entry(slot.path.clone())
        .or_insert_with(|| allocate(units.clone()));
    if entry.units != units {
        *entry = allocate(units);
    }
    Ok(entry.pointer)
}
pub(super) fn value_pointer(value: Value) -> Result<i64, RuntimeError> {
    if matches!(value, Value::NullString) {
        return Ok(0);
    }
    let units = if let Some(units) = value.string_units() {
        units.into_owned()
    } else {
        match value {
            Value::Null => return Err(invalid_null(None)),
            Value::Nothing => {
                return Err(error(
                    RuntimeErrorKind::ObjectVariableNotSet,
                    "Nothing has no String",
                    None,
                ))
            }
            Value::Error(_) | Value::Array(_) | Value::Record(_) | Value::Object(_) => {
                return Err(error(
                    RuntimeErrorKind::TypeMismatch,
                    "StrPtr cannot convert this value to String",
                    None,
                ))
            }
            _ => text(&value)
                .map_err(|m| error(RuntimeErrorKind::TypeMismatch, m, None))?
                .encode_utf16()
                .collect(),
        }
    };
    Ok(allocate(units).pointer)
}
pub(super) fn literal_pointer(
    value: Value,
    pool: &mut BTreeMap<Vec<u16>, Allocation>,
) -> Result<i64, RuntimeError> {
    if matches!(value, Value::NullString) {
        return Ok(0);
    }
    let Some(units) = value.string_units() else {
        return value_pointer(value);
    };
    let units = units.into_owned();
    Ok(pool
        .entry(units.clone())
        .or_insert_with(|| allocate(units))
        .pointer)
}
