// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! The coupon-bond functions: the coupon dates and day counts (COUPPCD,
//! COUPNCD, COUPNUM, COUPDAYBS, COUPDAYS, COUPDAYSNC), PRICE, YIELD,
//! DURATION, MDURATION, ACCRINT, ACCRINTM, PRICEMAT and YIELDMAT. The
//! coupon calendar steps back from maturity keeping its day of the month
//! (the last day when maturity falls on one), the way the analysis add-in
//! has always done it.

use crate::datetime;
use crate::functions::{days_30_360, num, yearfrac, Arg};
use crate::value::{ExcelError, Value};

pub(crate) const NAMES: &[&str] = &[
    "ACCRINT", "ACCRINTM", "COUPDAYBS", "COUPDAYS", "COUPDAYSNC", "COUPNCD", "COUPNUM", "COUPPCD", "DURATION",
    "MDURATION", "PRICE", "PRICEMAT", "YIELD", "YIELDMAT", "ODDFPRICE", "ODDFYIELD", "ODDLPRICE", "ODDLYIELD",
];

/// A coupon date: its year and month, the day of the month maturity has,
/// and whether that day is the last of its month.
#[derive(Clone, Copy)]
struct CouponDate {
    year: i64,
    month: i64,
    day: i64,
    last: bool,
}

impl CouponDate {
    fn from_serial(serial: i64) -> Result<Self, ExcelError> {
        let d = datetime::date_from_serial(serial)?;
        Ok(CouponDate { year: d.year, month: d.month, day: d.day, last: d.day == datetime::days_in_month(d.year, d.month) })
    }
    fn serial(&self) -> Result<i64, ExcelError> {
        let length = datetime::days_in_month(self.year, self.month);
        let day = if self.last { length } else { self.day.min(length) };
        datetime::serial_from_date(self.year, self.month, day)
    }
    fn add_months(&mut self, months: i64) {
        let index = self.year * 12 + self.month - 1 + months;
        self.year = index.div_euclid(12);
        self.month = index.rem_euclid(12) + 1;
    }
}

/// The coupon date on or before settlement.
fn previous_coupon(settle: i64, maturity: i64, frequency: i64) -> Result<CouponDate, ExcelError> {
    let mut date = CouponDate::from_serial(maturity)?;
    date.year = datetime::date_from_serial(settle)?.year;
    if date.serial()? < settle {
        date.year += 1;
    }
    while date.serial()? > settle {
        date.add_months(-12 / frequency);
    }
    Ok(date)
}

/// The coupon date after settlement.
fn next_coupon(settle: i64, maturity: i64, frequency: i64) -> Result<CouponDate, ExcelError> {
    let mut date = CouponDate::from_serial(maturity)?;
    date.year = datetime::date_from_serial(settle)?.year;
    if date.serial()? > settle {
        date.year -= 1;
    }
    while date.serial()? <= settle {
        date.add_months(12 / frequency);
    }
    Ok(date)
}

/// Days between two dates on the basis: 30/360 US or European, else actual.
fn days_between(start: i64, end: i64, basis: i64) -> Result<f64, ExcelError> {
    Ok(match basis {
        0 => days_30_360(start, end, false)? as f64,
        4 => days_30_360(start, end, true)? as f64,
        _ => (end - start) as f64,
    })
}

fn coupon_number(settle: i64, maturity: i64, frequency: i64) -> Result<f64, ExcelError> {
    let before = previous_coupon(settle, maturity, frequency)?;
    let end = datetime::date_from_serial(maturity)?;
    let months = (end.year - before.year) * 12 + end.month - before.month;
    Ok((months * frequency / 12) as f64)
}

fn coupon_days(settle: i64, maturity: i64, frequency: i64, basis: i64) -> Result<f64, ExcelError> {
    Ok(match basis {
        // On the actual basis the period is counted from the previous
        // coupon date found WITHOUT the month-end rule, to the same day
        // one period on. Measured, with maturity 28 Feb 2030 and
        // settlement 25 Jan 2024: 365 yearly, 184 half-yearly, 92
        // quarterly; maturity 31 Aug 2029, settlement 31 May 2024: 182.
        1 => {
            let mut start = CouponDate::from_serial(maturity)?;
            start.last = false;
            start.year = datetime::date_from_serial(settle)?.year;
            if start.serial()? < settle {
                start.year += 1;
            }
            while start.serial()? > settle {
                start.add_months(-12 / frequency);
            }
            let from = start.serial()?;
            let mut end = CouponDate::from_serial(from)?;
            end.last = false;
            end.add_months(12 / frequency);
            (end.serial()? - from) as f64
        }
        3 => 365.0 / frequency as f64,
        _ => 360.0 / frequency as f64,
    })
}

fn days_before(settle: i64, maturity: i64, frequency: i64, basis: i64) -> Result<f64, ExcelError> {
    days_between(previous_coupon(settle, maturity, frequency)?.serial()?, settle, basis)
}

fn days_after(settle: i64, maturity: i64, frequency: i64, basis: i64) -> Result<f64, ExcelError> {
    if basis != 0 && basis != 4 {
        return Ok((next_coupon(settle, maturity, frequency)?.serial()? - settle) as f64);
    }
    // European 30/360 counts straight to the next coupon: measured, 90
    // from 31 May to 31 Aug where the period less the days before is 89.
    if basis == 4 {
        return days_between(settle, next_coupon(settle, maturity, frequency)?.serial()?, basis);
    }
    Ok(coupon_days(settle, maturity, frequency, basis)? - days_before(settle, maturity, frequency, basis)?)
}

fn price(settle: i64, maturity: i64, rate: f64, yld: f64, redemption: f64, frequency: i64, basis: i64) -> Result<f64, ExcelError> {
    let f = frequency as f64;
    let e = coupon_days(settle, maturity, frequency, basis)?;
    let dsc_e = days_after(settle, maturity, frequency, basis)? / e;
    let n = coupon_number(settle, maturity, frequency)?;
    let a = days_before(settle, maturity, frequency, basis)?;
    let mut answer = redemption / (1.0 + yld / f).powf(n - 1.0 + dsc_e);
    answer -= 100.0 * rate / f * a / e;
    let (t1, t2) = (100.0 * rate / f, 1.0 + yld / f);
    let mut k = 0.0;
    while k < n {
        answer += t1 / t2.powf(k + dsc_e);
        k += 1.0;
    }
    Ok(answer)
}

/// Macaulay duration, each flow timed at k - 1 + DSC/E periods: measured,
/// DURATION(1 Jul 2018, 1 Jan 2048, 8%, 9%, 2, 1) 10.9191452815919 (timing
/// by YEARFRAC instead gives 10.92157...).
fn duration(settle: i64, maturity: i64, coupon: f64, yld: f64, frequency: i64, basis: i64) -> Result<f64, ExcelError> {
    let f = frequency as f64;
    let coupons = coupon_number(settle, maturity, frequency)?;
    let cash = coupon * 100.0 / f;
    let growth = 1.0 + yld / f;
    let shift = days_after(settle, maturity, frequency, basis)? / coupon_days(settle, maturity, frequency, basis)? - 1.0;
    let mut weighted = 0.0;
    let mut present = 0.0;
    let mut t = 1.0;
    while t < coupons {
        weighted += (t + shift) * cash / growth.powf(t + shift);
        present += cash / growth.powf(t + shift);
        t += 1.0;
    }
    weighted += (coupons + shift) * (cash + 100.0) / growth.powf(coupons + shift);
    present += (cash + 100.0) / growth.powf(coupons + shift);
    Ok(weighted / present / f)
}

/// The length of a quasi-coupon period on the basis: its actual days on
/// the actual basis, else the basis's year over the frequency.
fn normal_length(start: i64, end: i64, frequency: i64, basis: i64) -> f64 {
    match basis {
        1 => (end - start) as f64,
        3 => 365.0 / frequency as f64,
        _ => 360.0 / frequency as f64,
    }
}

/// A bond whose first period is odd, short or long, as Excel documents it
/// with quasi-coupon periods stepped back from the first coupon:
/// P = rv/v^(N+t) + c*sum(DC_i/NL_i)/v^t + sum_{k=1..N} c/v^(k+t)
///     - c*sum(A_i/NL_i), t = Nq + DSC/E.
/// Measured: ODDFPRICE(11 Nov 2008, 1 Mar 2021, 15 Oct 2008, 1 Mar 2009,
/// 7.85%, 6.25%, 100, 2, 1) 113.597717474079.
#[allow(clippy::too_many_arguments)]
fn odd_first_price(
    settle: i64,
    maturity: i64,
    issue: i64,
    first: i64,
    rate: f64,
    yld: f64,
    redemption: f64,
    frequency: i64,
    basis: i64,
) -> Result<f64, ExcelError> {
    let f = frequency as f64;
    let mut dates = vec![first];
    let mut date = CouponDate::from_serial(first)?;
    loop {
        date.add_months(-12 / frequency);
        let serial = date.serial()?;
        dates.push(serial);
        if serial <= issue {
            break;
        }
    }
    dates.reverse();
    let periods = dates.len() - 1;
    let (mut counted, mut accrued) = (0.0, 0.0);
    let (mut whole, mut dsc, mut e) = (0.0, 0.0, 1.0);
    for i in 0..periods {
        let (from, to) = (dates[i], dates[i + 1]);
        let length = normal_length(from, to, frequency, basis);
        let start = from.max(issue);
        counted += days_between(start, to, basis)? / length;
        if settle > start {
            accrued += days_between(start, settle.min(to), basis)? / length;
        }
        if from <= settle && settle < to {
            whole = (periods - 1 - i) as f64;
            dsc = days_between(settle, to, basis)?;
            e = length;
        }
    }
    let n = coupon_number(first, maturity, frequency)?;
    let c = 100.0 * rate / f;
    let v = 1.0 + yld / f;
    let t = whole + dsc / e;
    let mut answer = redemption / v.powf(n + t) + c * counted / v.powf(t) - c * accrued;
    let mut k = 1.0;
    while k <= n {
        answer += c / v.powf(k + t);
        k += 1.0;
    }
    Ok(answer)
}

/// The sums over a bond's odd last period, its quasi-coupon periods
/// stepped forward from the last coupon: (sum DC_i/NL_i, sum A_i/NL_i,
/// sum DSC_i/NL_i).
fn odd_last_sums(settle: i64, maturity: i64, last: i64, frequency: i64, basis: i64) -> Result<(f64, f64, f64), ExcelError> {
    let (mut counted, mut accrued, mut remaining) = (0.0, 0.0, 0.0);
    let mut date = CouponDate::from_serial(last)?;
    let mut from = last;
    while from < maturity {
        date.add_months(12 / frequency);
        let to = date.serial()?;
        // The odd last period counts its quasi-coupon periods at their
        // actual length on every actual basis: measured, bases 2 and 3 price
        // as basis 1 (99.8791676815291).
        let length = normal_length(from, to, frequency, if matches!(basis, 2 | 3) { 1 } else { basis });
        let end = to.min(maturity);
        counted += days_between(from, end, basis)? / length;
        if settle > from {
            accrued += days_between(from, settle.min(to), basis)? / length;
        }
        let start = settle.max(from);
        if end > start {
            remaining += days_between(start, end, basis)? / length;
        }
        from = to;
    }
    Ok((counted, accrued, remaining))
}

pub(crate) fn call(name: &str, args: &[Arg]) -> Result<Value, ExcelError> {
    let at = |i: usize| -> Result<f64, ExcelError> {
        match args.get(i) {
            Some(arg) => num(arg),
            None => Err(ExcelError::Value),
        }
    };
    let optional = |i: usize, default: f64| -> Result<f64, ExcelError> {
        match args.get(i) {
            Some(arg) if !matches!(arg.scalar(), Value::Blank) => num(arg),
            _ => Ok(default),
        }
    };
    let day = |i: usize| -> Result<i64, ExcelError> {
        let value = at(i)?.trunc();
        if value < 0.0 {
            return Err(ExcelError::Num);
        }
        Ok(value as i64)
    };
    let basis_at = |i: usize| -> Result<i64, ExcelError> {
        let basis = optional(i, 0.0)?.trunc();
        if !(0.0..=4.0).contains(&basis) {
            return Err(ExcelError::Num);
        }
        Ok(basis as i64)
    };
    let frequency_at = |i: usize| -> Result<i64, ExcelError> {
        match at(i)?.trunc() as i64 {
            f @ (1 | 2 | 4) => Ok(f),
            _ => Err(ExcelError::Num),
        }
    };
    let count = |low: usize, high: usize| {
        if args.len() < low || args.len() > high {
            Err(ExcelError::Value)
        } else {
            Ok(())
        }
    };
    let number = |value: f64| {
        if value.is_finite() {
            Ok(Value::Number(value))
        } else {
            Err(ExcelError::Num)
        }
    };
    match name {
        "COUPPCD" | "COUPNCD" | "COUPNUM" | "COUPDAYBS" | "COUPDAYS" | "COUPDAYSNC" => {
            count(3, 4)?;
            let (settle, maturity) = (day(0)?, day(1)?);
            let frequency = frequency_at(2)?;
            let basis = basis_at(3)?;
            if settle >= maturity {
                return Err(ExcelError::Num);
            }
            number(match name {
                "COUPPCD" => previous_coupon(settle, maturity, frequency)?.serial()? as f64,
                "COUPNCD" => next_coupon(settle, maturity, frequency)?.serial()? as f64,
                "COUPNUM" => coupon_number(settle, maturity, frequency)?,
                "COUPDAYBS" => days_before(settle, maturity, frequency, basis)?,
                "COUPDAYS" => coupon_days(settle, maturity, frequency, basis)?,
                _ => days_after(settle, maturity, frequency, basis)?,
            })
        }
        "PRICE" => {
            count(6, 7)?;
            let (settle, maturity) = (day(0)?, day(1)?);
            let (rate, yld, redemption) = (at(2)?, at(3)?, at(4)?);
            let (frequency, basis) = (frequency_at(5)?, basis_at(6)?);
            if settle >= maturity || rate < 0.0 || yld < 0.0 || redemption <= 0.0 {
                return Err(ExcelError::Num);
            }
            number(price(settle, maturity, rate, yld, redemption, frequency, basis)?)
        }
        "YIELD" => {
            count(6, 7)?;
            let (settle, maturity) = (day(0)?, day(1)?);
            let (rate, target, redemption) = (at(2)?, at(3)?, at(4)?);
            let (frequency, basis) = (frequency_at(5)?, basis_at(6)?);
            if settle >= maturity || rate < 0.0 || target <= 0.0 || redemption <= 0.0 {
                return Err(ExcelError::Num);
            }
            let coupons = coupon_number(settle, maturity, frequency)?;
            if coupons <= 1.0 {
                // One period left: the simple yield Excel documents.
                let f = frequency as f64;
                let e = coupon_days(settle, maturity, frequency, basis)?;
                let a = days_before(settle, maturity, frequency, basis)?;
                let dsr = e - a;
                let answer = ((redemption / 100.0 + rate / f) - (target / 100.0 + a / e * rate / f))
                    / (target / 100.0 + a / e * rate / f)
                    * f
                    * e
                    / dsr;
                return number(answer);
            }
            // Bracketing and false position between yields of 0 and 1, the
            // bracket doubled while the price is still above the target.
            let at_yield = |y: f64| price(settle, maturity, rate, y, redemption, frequency, basis);
            let equal = |a: f64, b: f64| a == b || (a - b).abs() <= 1e-15 * a.abs().max(b.abs());
            let (mut y1, mut y2) = (0.0, 1.0);
            let (mut p1, mut p2) = (at_yield(y1)?, at_yield(y2)?);
            let mut yn = (y2 - y1) * 0.5;
            let mut pn = 0.0;
            for _ in 0..100 {
                if equal(pn, target) {
                    break;
                }
                pn = at_yield(yn)?;
                if equal(target, p1) {
                    return number(y1);
                } else if equal(target, p2) {
                    return number(y2);
                } else if equal(target, pn) {
                    return number(yn);
                } else if target < p2 {
                    y2 *= 2.0;
                    p2 = at_yield(y2)?;
                    yn = (y2 - y1) * 0.5;
                } else {
                    if target < pn {
                        y1 = yn;
                        p1 = pn;
                    } else {
                        y2 = yn;
                        p2 = pn;
                    }
                    yn = y2 - (y2 - y1) * ((target - p2) / (p1 - p2));
                }
            }
            number(yn)
        }
        "DURATION" | "MDURATION" => {
            count(5, 6)?;
            let (settle, maturity) = (day(0)?, day(1)?);
            let (coupon, yld) = (at(2)?, at(3)?);
            let (frequency, basis) = (frequency_at(4)?, basis_at(5)?);
            if settle >= maturity || coupon < 0.0 || yld < 0.0 {
                return Err(ExcelError::Num);
            }
            let d = duration(settle, maturity, coupon, yld, frequency, basis)?;
            number(if name == "MDURATION" { d / (1.0 + yld / frequency as f64) } else { d })
        }
        // ACCRINT = par x rate/frequency x the sum, over the coupon periods
        // the accrual runs through, of the days accrued in each over that
        // period's length. Measured: 1 Jan to 15 Mar 2024 half-yearly on
        // the actual basis is 25 x 74/182 = 10.1648351648352.
        "ACCRINT" => {
            count(6, 8)?;
            let (issue, first, settle) = (day(0)?, day(1)?, day(2)?);
            let (rate, par) = (at(3)?, optional(4, 1000.0)?);
            let frequency = frequency_at(5)?;
            let basis = basis_at(6)?;
            let from_issue = match args.get(7).map(Arg::scalar) {
                None | Some(Value::Blank) => true,
                Some(value) => value.to_logical()?,
            };
            if issue >= settle || rate <= 0.0 || par <= 0.0 {
                return Err(ExcelError::Num);
            }
            let start = if !from_issue && settle > first { first } else { issue };
            let step = 12 / frequency;
            let mut edge = CouponDate::from_serial(first)?;
            while edge.serial()? > start {
                edge.add_months(-step);
            }
            let mut total = 0.0;
            loop {
                let period_start = edge.serial()?;
                if period_start >= settle {
                    break;
                }
                let mut next = edge;
                next.add_months(step);
                let period_end = next.serial()?;
                let length = match basis {
                    1 => (period_end - period_start) as f64,
                    3 => 365.0 / frequency as f64,
                    _ => 360.0 / frequency as f64,
                };
                let (from, to) = (start.max(period_start), settle.min(period_end));
                if to > from {
                    total += days_between(from, to, basis)? / length;
                }
                edge = next;
            }
            number(par * rate / frequency as f64 * total)
        }
        "ACCRINTM" => {
            count(3, 5)?;
            let (issue, settle) = (day(0)?, day(1)?);
            let (rate, par) = (at(2)?, optional(3, 1000.0)?);
            let basis = basis_at(4)?;
            if issue >= settle || rate <= 0.0 || par <= 0.0 {
                return Err(ExcelError::Num);
            }
            number(par * rate * yearfrac(issue, settle, basis)?)
        }
        "PRICEMAT" | "YIELDMAT" => {
            count(5, 6)?;
            let (settle, maturity, issue) = (day(0)?, day(1)?, day(2)?);
            let (rate, fourth) = (at(3)?, at(4)?);
            let basis = basis_at(5)?;
            if settle >= maturity || rate < 0.0 || fourth < 0.0 {
                return Err(ExcelError::Num);
            }
            // Written as Excel documents them, in days over the year's days:
            // measured, PRICEMAT 99.9844988755569 and YIELDMAT
            // 0.0609543336915387 come out that way (year fractions give
            // the last digit differently).
            let (dim, a, dsm, b) = if basis == 1 {
                (yearfrac(issue, maturity, basis)?, yearfrac(issue, settle, basis)?, yearfrac(settle, maturity, basis)?, 1.0)
            } else {
                let year = if basis == 3 { 365.0 } else { 360.0 };
                (days_between(issue, maturity, basis)?, days_between(issue, settle, basis)?, days_between(settle, maturity, basis)?, year)
            };
            number(if name == "PRICEMAT" {
                (100.0 + dim / b * rate * 100.0) / (1.0 + dsm / b * fourth) - a / b * rate * 100.0
            } else {
                let paid = fourth / 100.0 + a / b * rate;
                ((1.0 + dim / b * rate) - paid) / paid * (b / dsm)
            })
        }
        // ODDFPRICE and ODDFYIELD: settlement, maturity, issue, first
        // coupon, rate, yield or price, redemption, frequency, [basis];
        // issue < settlement < first coupon < maturity, else #NUM!.
        "ODDFPRICE" | "ODDFYIELD" => {
            count(8, 9)?;
            let (settle, maturity, issue, first) = (day(0)?, day(1)?, day(2)?, day(3)?);
            let (rate, given, redemption) = (at(4)?, at(5)?, at(6)?);
            let (frequency, basis) = (frequency_at(7)?, basis_at(8)?);
            if !(issue < settle && settle < first && first < maturity) || rate < 0.0 || given < 0.0 || redemption <= 0.0 {
                return Err(ExcelError::Num);
            }
            let price_at = |yld: f64| odd_first_price(settle, maturity, issue, first, rate, yld, redemption, frequency, basis);
            if name == "ODDFPRICE" {
                return number(price_at(given)?);
            }
            // The yield that prices the bond at `given`, by Newton's method.
            let mut yld = rate.max(0.01);
            for _ in 0..100 {
                let here = price_at(yld)? - given;
                let step = 1e-7;
                let slope = (price_at(yld + step)? - given - here) / step;
                if slope == 0.0 {
                    break;
                }
                let next = yld - here / slope;
                if (next - yld).abs() < 1e-12 {
                    yld = next;
                    break;
                }
                yld = next;
            }
            number(yld)
        }
        // ODDLPRICE and ODDLYIELD: settlement, maturity, last coupon, rate,
        // yield or price, redemption, frequency, [basis]; last coupon <
        // settlement < maturity, else #NUM!. Measured: ODDLPRICE(7 Feb 2008,
        // 15 Jun 2008, 15 Oct 2007, 3.75%, 4.05%, 100, 2, 0) 99.8782860147213.
        "ODDLPRICE" | "ODDLYIELD" => {
            count(7, 8)?;
            let (settle, maturity, last) = (day(0)?, day(1)?, day(2)?);
            let (rate, given, redemption) = (at(3)?, at(4)?, at(5)?);
            let (frequency, basis) = (frequency_at(6)?, basis_at(7)?);
            if !(last < settle && settle < maturity) || rate < 0.0 || given < 0.0 || redemption <= 0.0 {
                return Err(ExcelError::Num);
            }
            let f = frequency as f64;
            let c = 100.0 * rate / f;
            let (counted, accrued, remaining) = odd_last_sums(settle, maturity, last, frequency, basis)?;
            if name == "ODDLPRICE" {
                return number((redemption + counted * c) / (1.0 + remaining * given / f) - accrued * c);
            }
            let paid = given + accrued * c;
            number((redemption + counted * c - paid) / paid * f / remaining)
        }
        _ => Err(ExcelError::Name),
    }
}
