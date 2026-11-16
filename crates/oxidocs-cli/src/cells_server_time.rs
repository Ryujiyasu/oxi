// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Host file dates, converted to local whole seconds before encoding a DATE.

use std::fs::Metadata;

fn local_date(year: i64, month: i64, day: i64, hour: i64, minute: i64, second: i64) -> Result<f64, String> {
    if !(100..=9999).contains(&year) || !(1..=12).contains(&month) || !(1..=31).contains(&day)
        || !(0..=23).contains(&hour) || !(0..=59).contains(&minute) || !(0..=59).contains(&second) {
        return Err("file modification time is outside the VBA Date range".into());
    }
    // Gregorian days relative to the Unix epoch, with a March-based year.
    let year = year - i64::from(month <= 2);
    let era = year.div_euclid(400);
    let year_in_era = year - era * 400;
    let march_month = month + if month > 2 { -3 } else { 9 };
    let day_in_year = (153 * march_month + 2) / 5 + day - 1;
    let day_in_era = year_in_era * 365 + year_in_era / 4 - year_in_era / 100 + day_in_year;
    let days = era * 146_097 + day_in_era - 719_468 + 25_569;
    let fraction = (hour * 3600 + minute * 60 + second) as f64 / 86_400.0;
    // A negative OLE DATE carries its time in the absolute fractional part.
    Ok(days as f64 + if days < 0 { -fraction } else { fraction })
}

#[cfg(windows)]
pub(super) fn modified_time(metadata: &Metadata) -> Result<f64, String> {
    use std::os::windows::fs::MetadataExt;
    use windows_sys::Win32::Foundation::{FILETIME, SYSTEMTIME};
    use windows_sys::Win32::System::Time::{FileTimeToSystemTime, SystemTimeToTzSpecificLocalTimeEx};
    let ticks = metadata.last_write_time();
    let file_time = FILETIME { dwLowDateTime: ticks as u32, dwHighDateTime: (ticks >> 32) as u32 };
    let mut utc: SYSTEMTIME = unsafe { std::mem::zeroed() };
    let mut local: SYSTEMTIME = unsafe { std::mem::zeroed() };
    // Both outputs are initialized POD records; inputs remain valid for each
    // call. A null timezone selects the OS timezone without changing it.
    if unsafe { FileTimeToSystemTime(&file_time, &mut utc) } == 0
        || unsafe { SystemTimeToTzSpecificLocalTimeEx(std::ptr::null(), &utc, &mut local) } == 0 {
        return Err(format!("cannot convert file modification time: {}", std::io::Error::last_os_error()));
    }
    local_date(i64::from(local.wYear), i64::from(local.wMonth), i64::from(local.wDay),
        i64::from(local.wHour), i64::from(local.wMinute), i64::from(local.wSecond))
}

#[cfg(unix)]
pub(super) fn modified_time(metadata: &Metadata) -> Result<f64, String> {
    use std::time::UNIX_EPOCH;
    let modified = metadata.modified().map_err(|error| error.to_string())?;
    let seconds: i128 = match modified.duration_since(UNIX_EPOCH) {
        Ok(duration) => i128::from(duration.as_secs()),
        Err(error) => -i128::from(error.duration().as_secs()) - i128::from(error.duration().subsec_nanos() != 0),
    };
    let native: libc::time_t = seconds.try_into().map_err(|_| "file modification time exceeds host time_t")?;
    let mut local: libc::tm = unsafe { std::mem::zeroed() };
    // The caller owns both records and localtime_r writes only to this output.
    if unsafe { libc::localtime_r(&native, &mut local) }.is_null() {
        return Err(format!("cannot convert file modification time: {}", std::io::Error::last_os_error()));
    }
    local_date(i64::from(local.tm_year) + 1900, i64::from(local.tm_mon) + 1, i64::from(local.tm_mday),
        i64::from(local.tm_hour), i64::from(local.tm_min), i64::from(local.tm_sec))
}

#[cfg(not(any(windows, unix)))]
pub(super) fn modified_time(_metadata: &Metadata) -> Result<f64, String> {
    Err("local file modification times are unavailable on this host".into())
}

#[cfg(test)]
mod tests {
    use super::local_date;

    #[test]
    fn file_dates_keep_native_negative_fractions_and_gregorian_boundaries() {
        for (date, expected) in [
            ((1899, 12, 29, 12, 0, 0), -1.5),
            ((1899, 12, 30, 12, 0, 0), 0.5),
            ((1899, 12, 31, 12, 0, 0), 1.5),
            ((1900, 1, 1, 0, 0, 0), 2.0),
            ((1900, 2, 28, 0, 0, 0), 60.0),
            ((1900, 3, 1, 0, 0, 0), 61.0),
            ((1970, 1, 1, 0, 0, 0), 25569.0),
            ((2020, 2, 29, 12, 4, 5), 43890.50283564815),
            ((2030, 1, 2, 21, 34, 56), 47485.89925925926),
        ] {
            assert_eq!(local_date(date.0, date.1, date.2, date.3, date.4, date.5).unwrap(), expected);
        }
        assert!(local_date(99, 12, 31, 0, 0, 0).is_err());
        assert!(local_date(10000, 1, 1, 0, 0, 0).is_err());
    }
}
