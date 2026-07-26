// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Excel's date serial numbers.
//!
//! A date is a number: the integer part counts days, the fractional part is the
//! time of day. Serial 1 is 1900-01-01, so `=A1+1` on a date means "tomorrow".
//!
//! # The 1900 leap year bug
//!
//! Excel believes 1900 was a leap year. It was not: century years are leap years
//! only when divisible by 400. Serial 60 therefore denotes 1900-02-29, a day
//! that never existed, and every serial from 61 onward is shifted by one day
//! relative to a naive "days since 1900-01-01" count.
//!
//! The bug was inherited from Lotus 1-2-3 and kept deliberately for
//! compatibility. It is reproduced here, because the goal is to agree with
//! Excel, not to be right about the Gregorian calendar. Getting this wrong
//! shifts every date in a workbook by one day — silently.
//!
//! Volatile functions (`NOW`, `TODAY`) are not implemented: they read the wall
//! clock and would make recalculation non-reproducible. They need an injected
//! clock before they can be added.

use crate::value::ExcelError;

/// Days from the Unix epoch (1970-01-01) back to 1900-01-01.
const UNIX_DAYS_TO_1900: i64 = -25_567;
/// Offset applied to serials at or before 1900-02-28 (serial 59).
const OFFSET_BEFORE_PHANTOM: i64 = 25_568;
/// Offset applied from 1900-03-01 (serial 61) onward, absorbing the phantom day.
const OFFSET_AFTER_PHANTOM: i64 = 25_569;
/// The serial Excel assigns to the day that never existed, 1900-02-29.
pub const PHANTOM_LEAP_DAY: i64 = 60;
/// Last serial Excel can represent: 9999-12-31.
pub const MAX_SERIAL: i64 = 2_958_465;

/// Convert a calendar date to days since the Unix epoch.
///
/// Howard Hinnant's `days_from_civil`, valid for any proleptic Gregorian date.
fn days_from_civil(year: i64, month: i64, day: i64) -> i64 {
    let y = year - if month <= 2 { 1 } else { 0 };
    let era = if y >= 0 { y } else { y - 399 } / 400;
    let yoe = y - era * 400;
    let mp = month + if month > 2 { -3 } else { 9 };
    let doy = (153 * mp + 2) / 5 + day - 1;
    let doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
    era * 146_097 + doe - 719_468
}

/// Inverse of [`days_from_civil`].
fn civil_from_days(days: i64) -> (i64, i64, i64) {
    let z = days + 719_468;
    let era = if z >= 0 { z } else { z - 146_096 } / 146_097;
    let doe = z - era * 146_097;
    let yoe = (doe - doe / 1460 + doe / 36_524 - doe / 146_096) / 365;
    let y = yoe + era * 400;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let d = doy - (153 * mp + 2) / 5 + 1;
    let m = mp + if mp < 10 { 3 } else { -9 };
    (y + i64::from(m <= 2), m, d)
}

/// A decomposed date. `day` is `0` only for serial `0`, which Excel renders as
/// "January 0, 1900".
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Date {
    pub year: i64,
    pub month: i64,
    pub day: i64,
}

/// Split a serial into its date part.
///
/// Serial 60 yields 1900-02-29 — the day that does not exist — because that is
/// what Excel reports for it.
pub fn date_from_serial(serial: i64) -> Result<Date, ExcelError> {
    if !(0..=MAX_SERIAL).contains(&serial) {
        return Err(ExcelError::Num);
    }
    if serial == 0 {
        return Ok(Date {
            year: 1900,
            month: 1,
            day: 0,
        });
    }
    if serial == PHANTOM_LEAP_DAY {
        return Ok(Date {
            year: 1900,
            month: 2,
            day: 29,
        });
    }

    let offset = if serial < PHANTOM_LEAP_DAY {
        OFFSET_BEFORE_PHANTOM
    } else {
        OFFSET_AFTER_PHANTOM
    };
    let (year, month, day) = civil_from_days(serial - offset);
    Ok(Date { year, month, day })
}

/// Build a serial from a calendar date.
///
/// Out-of-range months and days roll over the way Excel's `DATE` does:
/// `DATE(2026,13,1)` is 2027-01-01 and `DATE(2026,1,32)` is 2026-02-01.
pub fn serial_from_date(year: i64, month: i64, day: i64) -> Result<i64, ExcelError> {
    // Excel maps a two-digit style year into the 1900s: DATE(26,1,1) is 1926.
    let year = if (0..1900).contains(&year) {
        year + 1900
    } else {
        year
    };

    // Normalise the month first so that day roll-over lands in the right month.
    let month_index = year * 12 + (month - 1);
    let (year, month) = (month_index.div_euclid(12), month_index.rem_euclid(12) + 1);

    let unix_days = days_from_civil(year, month, 1) + (day - 1);
    let serial = if unix_days <= UNIX_DAYS_TO_1900 + 58 {
        // On or before 1900-02-28.
        unix_days + OFFSET_BEFORE_PHANTOM
    } else {
        unix_days + OFFSET_AFTER_PHANTOM
    };

    if !(0..=MAX_SERIAL).contains(&serial) {
        return Err(ExcelError::Num);
    }
    Ok(serial)
}

/// Day of week for a serial, as `WEEKDAY` return type 1: Sunday = 1.
///
/// The phase is anchored so that modern dates are right, which forces serial 1
/// to be a Sunday. 1900-01-01 was really a **Monday** — the phantom leap day
/// (§ module docs) shifts every weekday before 1900-03-01 by one, and Excel
/// reports the shifted value. Anchoring on the true weekday of 1900-01-01
/// instead would put every date after March 1900 off by a day.
pub fn weekday_sunday_one(serial: i64) -> i64 {
    (serial - 1).rem_euclid(7) + 1
}

/// Split the fractional part of a serial into hours, minutes and seconds.
///
/// The fraction is rounded to the nearest second first. Without that,
/// `TIME(11,59,59)` round-trips through binary floating point as
/// 11:59:58.999999 and `SECOND` reports 58.
pub fn time_from_fraction(value: f64) -> (i64, i64, i64) {
    let frac = value - value.floor();
    let total = (frac * 86_400.0).round() as i64 % 86_400;
    (total / 3600, (total % 3600) / 60, total % 60)
}

/// Build the fractional part of a serial from a clock time.
///
/// Components roll over: `TIME(25,0,0)` is 01:00, matching Excel.
pub fn fraction_from_time(hours: f64, minutes: f64, seconds: f64) -> f64 {
    let total = hours * 3600.0 + minutes * 60.0 + seconds;
    let day = total / 86_400.0;
    day - day.floor()
}

/// Add whole months to a serial, clamping the day to the end of the target
/// month: one month after 2026-01-31 is 2026-02-28, not 2026-03-03.
pub fn add_months(serial: i64, months: i64) -> Result<i64, ExcelError> {
    let date = date_from_serial(serial)?;
    let month_index = date.year * 12 + (date.month - 1) + months;
    let (year, month) = (month_index.div_euclid(12), month_index.rem_euclid(12) + 1);
    let day = date.day.min(days_in_month(year, month));
    serial_from_date(year, month, day)
}

/// Serial of the last day of the month `months` away from `serial`.
pub fn end_of_month(serial: i64, months: i64) -> Result<i64, ExcelError> {
    let date = date_from_serial(serial)?;
    let month_index = date.year * 12 + (date.month - 1) + months;
    let (year, month) = (month_index.div_euclid(12), month_index.rem_euclid(12) + 1);
    serial_from_date(year, month, days_in_month(year, month))
}

fn is_leap(year: i64) -> bool {
    (year % 4 == 0 && year % 100 != 0) || year % 400 == 0
}

fn days_in_month(year: i64, month: i64) -> i64 {
    match month {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
        4 | 6 | 9 | 11 => 30,
        // Note: the real calendar, not Excel's. Excel's phantom 1900-02-29 only
        // affects serial arithmetic, not the length it reports for a month.
        2 if is_leap(year) => 29,
        _ => 28,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn d(year: i64, month: i64, day: i64) -> Date {
        Date { year, month, day }
    }

    #[test]
    fn anchor_serials_match_excel() {
        assert_eq!(date_from_serial(1).unwrap(), d(1900, 1, 1));
        assert_eq!(date_from_serial(59).unwrap(), d(1900, 2, 28));
        assert_eq!(date_from_serial(61).unwrap(), d(1900, 3, 1));
        // 1970-01-01 is serial 25569 in Excel.
        assert_eq!(date_from_serial(25_569).unwrap(), d(1970, 1, 1));
        assert_eq!(date_from_serial(45_000).unwrap(), d(2023, 3, 15));
        assert_eq!(date_from_serial(MAX_SERIAL).unwrap(), d(9999, 12, 31));
    }

    #[test]
    fn serial_sixty_is_the_day_that_never_existed() {
        // Excel insists 1900 was a leap year; reproduce it rather than correct it.
        assert_eq!(date_from_serial(60).unwrap(), d(1900, 2, 29));
        // The real calendar disagrees, which is exactly why the offset changes
        // either side of this serial.
        assert!(!is_leap(1900));
    }

    #[test]
    fn serials_round_trip_across_the_phantom_day() {
        for serial in [1i64, 2, 58, 59, 61, 62, 1000, 25_569, 45_000, MAX_SERIAL] {
            let date = date_from_serial(serial).unwrap();
            assert_eq!(
                serial_from_date(date.year, date.month, date.day),
                Ok(serial),
                "serial {serial} -> {date:?}"
            );
        }
    }

    #[test]
    fn out_of_range_serials_are_num_errors() {
        assert_eq!(date_from_serial(-1), Err(ExcelError::Num));
        assert_eq!(date_from_serial(MAX_SERIAL + 1), Err(ExcelError::Num));
    }

    #[test]
    fn date_rolls_over_out_of_range_parts() {
        // DATE(2026,13,1) == DATE(2027,1,1)
        assert_eq!(
            serial_from_date(2026, 13, 1),
            serial_from_date(2027, 1, 1)
        );
        // DATE(2026,1,32) == DATE(2026,2,1)
        assert_eq!(serial_from_date(2026, 1, 32), serial_from_date(2026, 2, 1));
        // DATE(2026,0,1) == DATE(2025,12,1)
        assert_eq!(serial_from_date(2026, 0, 1), serial_from_date(2025, 12, 1));
    }

    #[test]
    fn two_digit_years_land_in_the_1900s() {
        assert_eq!(serial_from_date(26, 1, 1), serial_from_date(1926, 1, 1));
    }

    #[test]
    fn weekday_phase_is_anchored_on_modern_dates() {
        // 2026-07-26 is a Sunday, and 2026-07-27 a Monday.
        assert_eq!(weekday_sunday_one(serial_from_date(2026, 7, 26).unwrap()), 1);
        assert_eq!(weekday_sunday_one(serial_from_date(2026, 7, 27).unwrap()), 2);
        // 1900-03-01, the first day after the phantom, was a Thursday.
        assert_eq!(weekday_sunday_one(61), 5);
    }

    #[test]
    fn pre_march_1900_weekdays_are_wrong_the_way_excel_is_wrong() {
        // 1900-01-01 was a Monday in reality. Excel says Sunday, because the
        // phantom leap day has not been passed yet. Agreement with Excel wins.
        assert_eq!(weekday_sunday_one(1), 1);
    }

    #[test]
    fn time_components_survive_the_float_round_trip() {
        let frac = fraction_from_time(11.0, 59.0, 59.0);
        assert_eq!(time_from_fraction(frac), (11, 59, 59));
        assert_eq!(time_from_fraction(fraction_from_time(0.0, 0.0, 0.0)), (0, 0, 0));
        assert_eq!(time_from_fraction(fraction_from_time(23.0, 59.0, 59.0)), (23, 59, 59));
        // Components roll over: 25:00 is 01:00.
        assert_eq!(time_from_fraction(fraction_from_time(25.0, 0.0, 0.0)), (1, 0, 0));
    }

    #[test]
    fn month_arithmetic_clamps_to_the_end_of_the_month() {
        let jan31 = serial_from_date(2026, 1, 31).unwrap();
        // EDATE: one month after 31 Jan is 28 Feb, not 3 Mar.
        assert_eq!(
            date_from_serial(add_months(jan31, 1).unwrap()).unwrap(),
            d(2026, 2, 28)
        );
        // A leap year gives 29.
        let jan31_2024 = serial_from_date(2024, 1, 31).unwrap();
        assert_eq!(
            date_from_serial(add_months(jan31_2024, 1).unwrap()).unwrap(),
            d(2024, 2, 29)
        );
    }

    #[test]
    fn end_of_month_covers_the_japanese_fiscal_year() {
        // Fiscal year 2026 in Japan runs to 2027-03-31.
        let apr1 = serial_from_date(2026, 4, 1).unwrap();
        assert_eq!(
            date_from_serial(end_of_month(apr1, 11).unwrap()).unwrap(),
            d(2027, 3, 31)
        );
        assert_eq!(
            date_from_serial(end_of_month(apr1, 0).unwrap()).unwrap(),
            d(2026, 4, 30)
        );
    }
}
