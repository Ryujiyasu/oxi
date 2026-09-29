// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Worksheet functions added in a second pass: the reciprocal trigonometric
//! functions, bit shifts, the precise roundings, matrix functions, the
//! discount-security and T-bill money functions, the cumulative annuity
//! parts, XNPV and XIRR, VDB, and the tests CHISQ.TEST, F.TEST and Z.TEST.
//! Each value quoted below was read from Excel.

use crate::distributions as d;
use crate::functions::{
    block_of, fin_fv_raw, fin_pmt_raw, norm_cdf, num, numeric_operands, reach, truncate_significant, yearfrac,
    Arg, RangeData,
};
use crate::value::{ExcelError, Value};

/// The names this file answers for.
pub(crate) const NAMES: &[&str] = &[
    "ACOT", "ACOTH", "BINOM.DIST.RANGE", "BITLSHIFT", "BITRSHIFT", "CEILING.PRECISE", "CHISQ.TEST",
    "CHITEST", "COT", "COTH", "CSC", "CSCH", "CUMIPMT", "CUMPRINC", "DISC", "DOLLARDE", "DOLLARFR",
    "ECMA.CEILING", "EFFECT", "F.TEST", "FLOOR.PRECISE", "FTEST", "FVSCHEDULE", "INTRATE", "ISO.CEILING",
    "ISPMT", "MDETERM", "NOMINAL", "PDURATION", "PERCENTRANK.EXC", "PERMUTATIONA", "PRICEDISC", "PROB",
    "RECEIVED", "RRI", "SEC", "SECH", "SERIESSUM", "SKEW.P", "STDEVPA", "TBILLEQ", "TBILLPRICE",
    "TBILLYIELD", "VARPA", "VDB", "XIRR", "XNPV", "YIELDDISC", "Z.TEST", "ZTEST", "PERCENTOF", "ENCODEURL",
    "AMORLINC", "AMORDEGRC", "PHONETIC", "REGEXTEST", "REGEXREPLACE", "BAHTTEXT",
];

/// A whole number in Thai words, millions in blocks of six digits. A final
/// one is เอ็ด when anything stands above it, หนึ่ง alone.
fn thai_words(n: u64, above: bool) -> String {
    if n >= 1_000_000 {
        let mut out = thai_words(n / 1_000_000, above);
        out.push_str("ล้าน");
        out.push_str(&thai_block(n % 1_000_000, true));
        return out;
    }
    thai_block(n, above)
}

fn thai_block(n: u64, above: bool) -> String {
    const DIGITS: [&str; 10] = ["", "หนึ่ง", "สอง", "สาม", "สี่", "ห้า", "หก", "เจ็ด", "แปด", "เก้า"];
    const PLACES: [&str; 4] = ["แสน", "หมื่น", "พัน", "ร้อย"];
    let mut out = String::new();
    for (i, place) in PLACES.iter().enumerate() {
        let digit = (n / 10u64.pow(5 - i as u32) % 10) as usize;
        if digit > 0 {
            out.push_str(DIGITS[digit]);
            out.push_str(place);
        }
    }
    match n / 10 % 10 {
        0 => {}
        1 => out.push_str("สิบ"),
        2 => out.push_str("ยี่สิบ"),
        digit => {
            out.push_str(DIGITS[digit as usize]);
            out.push_str("สิบ");
        }
    }
    match n % 10 {
        0 => {}
        1 if above || n >= 10 => out.push_str("เอ็ด"),
        digit => out.push_str(DIGITS[digit as usize]),
    }
    out
}

/// A pattern of REGEXTEST, REGEXEXTRACT or REGEXREPLACE, case-blind when
/// the flag at `flag` is 1; one that will not compile is #VALUE!.
fn pattern(args: &[Arg], flag: usize) -> Result<crate::regex::Regex, ExcelError> {
    let text = args[1].scalar().to_text()?;
    let blind = match args.get(flag).map(Arg::scalar) {
        None | Some(Value::Blank) => false,
        Some(value) => match value.to_number()? {
            0.0 => false,
            1.0 => true,
            _ => return Err(ExcelError::Value),
        },
    };
    crate::regex::Regex::new(&text, blind, false).map_err(|_| ExcelError::Value)
}

fn at(args: &[Arg], i: usize) -> Result<f64, ExcelError> {
    match args.get(i) {
        Some(arg) => num(arg),
        None => Err(ExcelError::Value),
    }
}

fn optional(args: &[Arg], i: usize, default: f64) -> Result<f64, ExcelError> {
    match args.get(i) {
        Some(arg) if !matches!(arg.scalar(), Value::Blank) => num(arg),
        _ => Ok(default),
    }
}

fn count(args: &[Arg], low: usize, high: usize) -> Result<(), ExcelError> {
    if args.len() < low || args.len() > high {
        Err(ExcelError::Value)
    } else {
        Ok(())
    }
}

fn finite(value: f64) -> Result<Value, ExcelError> {
    if value.is_finite() {
        Ok(Value::Number(if value == 0.0 { 0.0 } else { value }))
    } else {
        Err(ExcelError::Num)
    }
}

/// Numbers of a block or list, in order, skipping text and blanks.
fn numbers_of(arg: &Arg) -> Result<Vec<f64>, ExcelError> {
    numeric_operands(std::slice::from_ref(arg))
}

/// Settlement and maturity as whole days, the first before the second.
fn span(args: &[Arg]) -> Result<(i64, i64), ExcelError> {
    let (settle, maturity) = (at(args, 0)?.trunc(), at(args, 1)?.trunc());
    if settle < 0.0 || maturity < 0.0 || settle >= maturity {
        return Err(ExcelError::Num);
    }
    Ok((settle as i64, maturity as i64))
}

fn basis(args: &[Arg], i: usize) -> Result<i64, ExcelError> {
    let basis = optional(args, i, 0.0)?.trunc();
    if !(0.0..=4.0).contains(&basis) {
        return Err(ExcelError::Num);
    }
    Ok(basis as i64)
}

pub(crate) fn call(name: &str, args: &[Arg]) -> Result<Value, ExcelError> {
    match name {
        // ---- the reciprocal trigonometric functions ----------------------
        "ACOT" => {
            count(args, 1, 1)?;
            finite(std::f64::consts::FRAC_PI_2 - at(args, 0)?.atan())
        }
        "ACOTH" => {
            count(args, 1, 1)?;
            let x = at(args, 0)?;
            if x.abs() <= 1.0 {
                return Err(ExcelError::Num);
            }
            finite(0.5 * ((x + 1.0) / (x - 1.0)).ln())
        }
        "COT" | "CSC" | "SEC" | "COTH" | "CSCH" | "SECH" => {
            count(args, 1, 1)?;
            let x = at(args, 0)?;
            if matches!(name, "COT" | "CSC" | "SEC") && x.abs() >= 134_217_728.0 {
                return Err(ExcelError::Num);
            }
            let below = match name {
                "COT" => x.tan(),
                "CSC" => x.sin(),
                "SEC" => x.cos(),
                "COTH" => x.tanh(),
                "CSCH" => x.sinh(),
                _ => x.cosh(),
            };
            if below == 0.0 {
                return Err(ExcelError::DivZero);
            }
            finite(1.0 / below)
        }

        // ---- bits and roundings ------------------------------------------
        // BITLSHIFT(5,2) 20, BITRSHIFT(20,-2) 80: a whole number below 2^48,
        // shifted no further than 53 places.
        "BITLSHIFT" | "BITRSHIFT" => {
            count(args, 2, 2)?;
            let (n, shift) = (at(args, 0)?, at(args, 1)?.trunc());
            if n < 0.0 || n.fract() != 0.0 || n >= 281_474_976_710_656.0 || shift.abs() > 53.0 {
                return Err(ExcelError::Num);
            }
            let left = if name == "BITLSHIFT" { shift } else { -shift };
            let answer = if left >= 0.0 { n * 2f64.powf(left) } else { (n / 2f64.powf(-left)).floor() };
            if answer >= 281_474_976_710_656.0 {
                return Err(ExcelError::Num);
            }
            finite(answer)
        }
        // CEILING.PRECISE(-4.3,2) -4, FLOOR.PRECISE(-4.3,2) -6: the sign of
        // the step does not matter; ISO.CEILING and ECMA.CEILING round up the
        // same way.
        "CEILING.PRECISE" | "ISO.CEILING" | "ECMA.CEILING" | "FLOOR.PRECISE" => {
            count(args, 1, 2)?;
            let x = at(args, 0)?;
            let step = optional(args, 1, 1.0)?.abs();
            if step == 0.0 {
                return Ok(Value::Number(0.0));
            }
            let steps = x / step;
            let whole = if name == "FLOOR.PRECISE" { steps.floor() } else { steps.ceil() };
            finite(whole * step)
        }
        "PERMUTATIONA" => {
            count(args, 2, 2)?;
            let (n, k) = (at(args, 0)?.trunc(), at(args, 1)?.trunc());
            if n < 0.0 || k < 0.0 {
                return Err(ExcelError::Num);
            }
            finite(n.powf(k))
        }
        // SERIESSUM(2,1,2,{1,2,3}) = 1*2 + 2*2^3 + 3*2^5 = 114.
        "SERIESSUM" => {
            count(args, 4, 4)?;
            let (x, first, step) = (at(args, 0)?, at(args, 1)?, at(args, 2)?);
            let mut total = 0.0;
            for (i, coefficient) in args[3].flatten().iter().enumerate() {
                let coefficient = match coefficient {
                    Value::Number(n) => *n,
                    Value::Error(e) => return Err(*e),
                    _ => return Err(ExcelError::Value),
                };
                total += coefficient * x.powf(first + i as f64 * step);
            }
            finite(total)
        }

        // ---- statistics ------------------------------------------------------
        "SKEW.P" => {
            let values = numeric_operands(args)?;
            let n = values.len() as f64;
            if n < 1.0 {
                return Err(ExcelError::DivZero);
            }
            let mean = values.iter().sum::<f64>() / n;
            let spread = (values.iter().map(|x| (x - mean).powi(2)).sum::<f64>() / n).sqrt();
            if spread == 0.0 {
                return Err(ExcelError::DivZero);
            }
            // The cubes summed first, divided by the spread cubed after:
            // measured, SKEW.P(1,2,3,4,6) 0.395870337343817.
            finite(values.iter().map(|x| (x - mean).powi(3)).sum::<f64>() / n / spread.powi(3))
        }
        // The whole-population forms of STDEVA and VARA: text is 0, TRUE 1.
        "STDEVPA" | "VARPA" => {
            let mut values = Vec::new();
            for arg in args {
                match arg {
                    Arg::Value(value) => match value {
                        Value::Error(e) => return Err(*e),
                        Value::Blank => {}
                        other => values.push(other.to_number()?),
                    },
                    Arg::Range(block) => {
                        for value in &block.cells {
                            match value {
                                Value::Error(e) => return Err(*e),
                                Value::Number(n) => values.push(*n),
                                Value::Logical(b) => values.push(f64::from(*b)),
                                Value::Text(_) => values.push(0.0),
                                Value::Blank => {}
                            }
                        }
                    }
                }
            }
            if values.is_empty() {
                return Err(ExcelError::DivZero);
            }
            let n = values.len() as f64;
            let mean = values.iter().sum::<f64>() / n;
            let variance = values.iter().map(|x| (x - mean).powi(2)).sum::<f64>() / n;
            finite(if name == "VARPA" { variance } else { variance.sqrt() })
        }
        // PERCENTRANK.EXC(1,2,3,4,6; 3) 0.5: the place counted from one over
        // one more than the count, truncated to three digits by default.
        "PERCENTRANK.EXC" => {
            count(args, 2, 3)?;
            let mut values = numbers_of(&args[0])?;
            if values.is_empty() {
                return Err(ExcelError::Num);
            }
            values.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
            let x = at(args, 1)?;
            let digits = optional(args, 2, 3.0)?.trunc();
            if digits < 1.0 {
                return Err(ExcelError::Num);
            }
            let n = values.len();
            if x < values[0] || x > values[n - 1] {
                return Err(ExcelError::NA);
            }
            let mut place = (n - 1) as f64;
            for i in 0..n {
                if values[i] == x {
                    place = i as f64;
                    break;
                }
                if i + 1 < n && values[i] < x && x < values[i + 1] {
                    place = i as f64 + (x - values[i]) / (values[i + 1] - values[i]);
                    break;
                }
            }
            finite(truncate_significant((place + 1.0) / (n as f64 + 1.0), digits as i32))
        }
        // PROB(x, chances, low, [high]): the chance of the values from low to
        // high, the chances adding up to one.
        "PROB" => {
            count(args, 3, 4)?;
            let (xs, chances) = (args[0].flatten(), args[1].flatten());
            if xs.len() != chances.len() {
                return Err(ExcelError::NA);
            }
            let low = at(args, 2)?;
            let high = optional(args, 3, low)?;
            let mut total = 0.0;
            let mut sum = 0.0;
            for (x, chance) in xs.iter().zip(&chances) {
                let chance = match chance {
                    Value::Number(n) => *n,
                    Value::Error(e) => return Err(*e),
                    _ => continue,
                };
                if chance <= 0.0 || chance > 1.0 {
                    return Err(ExcelError::Num);
                }
                total += chance;
                if let Value::Number(x) = x {
                    if *x >= low && *x <= high {
                        sum += chance;
                    }
                }
            }
            if (total - 1.0).abs() > 1e-9 {
                return Err(ExcelError::Num);
            }
            finite(sum)
        }
        // BINOM.DIST.RANGE(10,0.3,2,4) 0.7004233215.
        "BINOM.DIST.RANGE" => {
            count(args, 3, 4)?;
            let (n, p, low) = (at(args, 0)?.trunc(), at(args, 1)?, at(args, 2)?.trunc());
            let high = optional(args, 3, low)?.trunc();
            if n < 0.0 || !(0.0..=1.0).contains(&p) || low < 0.0 || low > n || high < low || high > n {
                return Err(ExcelError::Num);
            }
            let mut total = 0.0;
            let mut k = low;
            while k <= high {
                total += binomial(n, k) * p.powf(k) * (1.0 - p).powf(n - k);
                k += 1.0;
            }
            finite(total)
        }
        // CHISQ.TEST: the chance of a chi-squared this large, with (rows-1) x
        // (columns-1) degrees of freedom, or one less than the count for a
        // single line.
        "CHISQ.TEST" | "CHITEST" => {
            count(args, 2, 2)?;
            let (actual, expected) = (block_of(&args[0]), block_of(&args[1]));
            if actual.width != expected.width || actual.height != expected.height {
                return Err(ExcelError::NA);
            }
            let mut chi = 0.0;
            for (a, e) in actual.cells.iter().zip(&expected.cells) {
                if let (Value::Number(a), Value::Number(e)) = (a, e) {
                    if *e == 0.0 {
                        return Err(ExcelError::DivZero);
                    }
                    chi += (a - e).powi(2) / e;
                }
            }
            let freedom = if actual.width > 1 && actual.height > 1 {
                ((actual.width - 1) * (actual.height - 1)) as f64
            } else {
                (actual.cells.len() - 1) as f64
            };
            if freedom < 1.0 {
                return Err(ExcelError::NA);
            }
            finite(d::regularized_gamma_q(freedom / 2.0, chi / 2.0))
        }
        // F.TEST: two tails of the F distribution at the ratio of the two
        // sample variances.
        "F.TEST" | "FTEST" => {
            count(args, 2, 2)?;
            let (first, second) = (numbers_of(&args[0])?, numbers_of(&args[1])?);
            if first.len() < 2 || second.len() < 2 {
                return Err(ExcelError::DivZero);
            }
            let variance = |values: &[f64]| {
                let n = values.len() as f64;
                let mean = values.iter().sum::<f64>() / n;
                values.iter().map(|x| (x - mean).powi(2)).sum::<f64>() / (n - 1.0)
            };
            let (v1, v2) = (variance(&first), variance(&second));
            if v1 == 0.0 || v2 == 0.0 {
                return Err(ExcelError::DivZero);
            }
            let (d1, d2) = ((first.len() - 1) as f64, (second.len() - 1) as f64);
            let lower = d::f_cdf(v1 / v2, d1, d2);
            finite(2.0 * lower.min(1.0 - lower))
        }
        // Z.TEST(1,2,3,4,6; 2) 0.0815121921532852: one tail above the mean,
        // the sample spread standing in for sigma when none is given.
        "Z.TEST" | "ZTEST" => {
            count(args, 2, 3)?;
            let values = numbers_of(&args[0])?;
            let n = values.len() as f64;
            if n < 1.0 {
                return Err(ExcelError::NA);
            }
            let mean = values.iter().sum::<f64>() / n;
            let sigma = match args.get(2) {
                Some(arg) => num(arg)?,
                None => {
                    if n < 2.0 {
                        return Err(ExcelError::DivZero);
                    }
                    (values.iter().map(|x| (x - mean).powi(2)).sum::<f64>() / (n - 1.0)).sqrt()
                }
            };
            if sigma == 0.0 {
                return Err(ExcelError::DivZero);
            }
            // z as (mean - x) * sqrt(n) / sigma: measured, the digits of
            // Z.TEST(1,2,3,4,6; 2) 0.0815121921532852 come out that way.
            finite(norm_cdf(-((mean - at(args, 1)?) * n.sqrt() / sigma)))
        }

        // ---- matrices --------------------------------------------------------
        "MDETERM" => {
            count(args, 1, 1)?;
            let matrix = square(&args[0])?;
            finite(determinant(matrix))
        }

        // ---- money -----------------------------------------------------------
        // CUMIPMT(0.01,12,1000,1,12,0) -66.1854641401001 and CUMPRINC -1000:
        // the interest and principal parts of payments start to end.
        "CUMIPMT" | "CUMPRINC" => {
            count(args, 6, 6)?;
            let (rate, periods, present) = (at(args, 0)?, at(args, 1)?, at(args, 2)?);
            let (start, end, kind) = (at(args, 3)?.trunc(), at(args, 4)?.trunc(), at(args, 5)?);
            if rate <= 0.0 || periods <= 0.0 || present <= 0.0 || start < 1.0 || end < start || end > periods
                || (kind != 0.0 && kind != 1.0)
            {
                return Err(ExcelError::Num);
            }
            // The principal is what the balance fell by, from before the first
            // payment to after the last; the interest is the payments less
            // that. Measured: CUMPRINC exactly -1000, CUMIPMT -66.1854641401001
            // (adding the parts one by one misses both in the last digit).
            let payment = fin_pmt_raw(rate, periods, present, 0.0, kind)?;
            // Paid at the start of each period, the parts are added up period
            // by period, each period's interest on the balance before its
            // payment: measured, CUMPRINC(0.05/12,60,10000,13,24,1) is
            // -1890.05280395008.
            if kind == 1.0 {
                let balance = |before: f64| fin_fv_raw(rate, before, payment, present, 1.0);
                let (mut principal_total, mut interest_total) = (0.0, 0.0);
                let mut at = start;
                while at <= end {
                    let interest = if at == 1.0 { 0.0 } else { (balance(at - 2.0)? - payment) * rate };
                    principal_total += payment - interest;
                    interest_total += interest;
                    at += 1.0;
                }
                return finite(if name == "CUMPRINC" { principal_total } else { interest_total });
            }
            let principal =
                fin_fv_raw(rate, start - 1.0, payment, present, kind)? - fin_fv_raw(rate, end, payment, present, kind)?;
            finite(if name == "CUMPRINC" { principal } else { payment * (end - start + 1.0) - principal })
        }
        "EFFECT" => {
            count(args, 2, 2)?;
            let (rate, times) = (at(args, 0)?, at(args, 1)?.trunc());
            if rate <= 0.0 || times < 1.0 {
                return Err(ExcelError::Num);
            }
            // Raised by squaring and multiplying, as Excel does: measured,
            // EFFECT(0.05,12) 0.0511618978817334 (powf gives ...330).
            let mut base = 1.0 + rate / times;
            let mut left = times as u64;
            let mut grown = 1.0;
            while left > 0 {
                if left & 1 == 1 {
                    grown *= base;
                }
                base *= base;
                left >>= 1;
            }
            finite(grown - 1.0)
        }
        "NOMINAL" => {
            count(args, 2, 2)?;
            let (rate, times) = (at(args, 0)?, at(args, 1)?.trunc());
            if rate <= 0.0 || times < 1.0 {
                return Err(ExcelError::Num);
            }
            finite(times * ((1.0 + rate).powf(1.0 / times) - 1.0))
        }
        "FVSCHEDULE" => {
            count(args, 2, 2)?;
            let mut value = at(args, 0)?;
            for rate in args[1].flatten() {
                match rate {
                    Value::Number(r) => value *= 1.0 + r,
                    Value::Blank => {}
                    Value::Error(e) => return Err(e),
                    _ => return Err(ExcelError::Value),
                }
            }
            finite(value)
        }
        "PDURATION" => {
            count(args, 3, 3)?;
            let (rate, present, future) = (at(args, 0)?, at(args, 1)?, at(args, 2)?);
            if rate <= 0.0 || present <= 0.0 || future <= 0.0 {
                return Err(ExcelError::Num);
            }
            finite((future.ln() - present.ln()) / (1.0 + rate).ln())
        }
        "RRI" => {
            count(args, 3, 3)?;
            let (periods, present, future) = (at(args, 0)?, at(args, 1)?, at(args, 2)?);
            if periods <= 0.0 || present == 0.0 {
                return Err(ExcelError::Num);
            }
            finite((future / present).powf(1.0 / periods) - 1.0)
        }
        "ISPMT" => {
            count(args, 4, 4)?;
            let (rate, period, periods, present) = (at(args, 0)?, at(args, 1)?, at(args, 2)?, at(args, 3)?);
            if periods == 0.0 {
                return Err(ExcelError::DivZero);
            }
            finite(present * rate * (period / periods - 1.0))
        }
        // DOLLARDE(1.02,16) 1.125 and DOLLARFR(1.125,16) 1.02: the fraction
        // written as so many sixteenths, and back.
        "DOLLARDE" | "DOLLARFR" => {
            count(args, 2, 2)?;
            let (value, fraction) = (at(args, 0)?, at(args, 1)?.trunc());
            if fraction < 0.0 {
                return Err(ExcelError::Num);
            }
            if fraction == 0.0 {
                return Err(ExcelError::DivZero);
            }
            let scale = 10f64.powf(fraction.log10().ceil());
            let whole = value.trunc();
            let part = value - whole;
            finite(if name == "DOLLARDE" {
                whole + part * scale / fraction
            } else {
                whole + part * fraction / scale
            })
        }
        // The discount securities, on the day-count basis given (30/360 by
        // default): DISC(45000,45365,97,100) 0.0300835654596101.
        "DISC" | "INTRATE" | "RECEIVED" | "PRICEDISC" | "YIELDDISC" => {
            count(args, 4, 5)?;
            let (settle, maturity) = span(args)?;
            let (third, fourth) = (at(args, 2)?, at(args, 3)?);
            if third <= 0.0 || fourth <= 0.0 {
                return Err(ExcelError::Num);
            }
            let years = yearfrac(settle, maturity, basis(args, 4)?)?;
            finite(match name {
                // Measured: 0.0300835654596101, written as 1 - price/redemption.
                "DISC" => (1.0 - third / fourth) / years,
                "INTRATE" => (fourth - third) / third / years,
                "RECEIVED" => {
                    let rest = 1.0 - fourth * years;
                    if rest <= 0.0 {
                        return Err(ExcelError::Num);
                    }
                    third / rest
                }
                "PRICEDISC" => fourth - third * fourth * years,
                _ => (fourth - third) / third / years,
            })
        }
        // T-bills run at most a year: TBILLPRICE(45000,45180,0.05) 97.5.
        "TBILLPRICE" | "TBILLYIELD" | "TBILLEQ" => {
            count(args, 3, 3)?;
            let (settle, maturity) = span(args)?;
            let days = (maturity - settle) as f64;
            if days > 365.0 {
                return Err(ExcelError::Num);
            }
            let given = at(args, 2)?;
            if given <= 0.0 {
                return Err(ExcelError::Num);
            }
            finite(match name {
                "TBILLPRICE" => {
                    let price = 100.0 * (1.0 - given * days / 360.0);
                    if price <= 0.0 {
                        return Err(ExcelError::Num);
                    }
                    price
                }
                "TBILLYIELD" => (100.0 - given) / given * 360.0 / days,
                _ => {
                    if days <= 182.0 {
                        365.0 * given / (360.0 - given * days)
                    } else {
                        let price = 100.0 * (1.0 - given * days / 360.0);
                        let term = days / 365.0;
                        (-term + (term * term - (2.0 * term - 1.0) * (1.0 - 100.0 / price)).sqrt()) / (term - 0.5)
                    }
                }
            })
        }
        // XNPV over dates counted from the first, at 365 a year.
        "XNPV" => {
            count(args, 3, 3)?;
            let rate = at(args, 0)?;
            let (flows, dates) = flows_and_dates(&args[1], &args[2])?;
            finite(xnpv(rate, &flows, &dates))
        }
        "XIRR" => {
            count(args, 2, 3)?;
            let (flows, dates) = flows_and_dates(&args[0], &args[1])?;
            if !flows.iter().any(|f| *f > 0.0) || !flows.iter().any(|f| *f < 0.0) {
                return Err(ExcelError::Num);
            }
            let mut rate = optional(args, 2, 0.1)?;
            for _ in 0..100 {
                let value = xnpv(rate, &flows, &dates);
                let slope: f64 = flows
                    .iter()
                    .zip(&dates)
                    .map(|(flow, date)| {
                        let years = (date - dates[0]) / 365.0;
                        -years * flow / (1.0 + rate).powf(years + 1.0)
                    })
                    .sum();
                if slope == 0.0 {
                    return Err(ExcelError::Num);
                }
                let next = rate - value / slope;
                if !next.is_finite() || next <= -1.0 {
                    return Err(ExcelError::Num);
                }
                if (next - rate).abs() < 1e-12 {
                    return finite(next);
                }
                rate = next;
            }
            Err(ExcelError::Num)
        }
        "VDB" => {
            count(args, 5, 7)?;
            let (cost, salvage, life, start, end) = (at(args, 0)?, at(args, 1)?, at(args, 2)?, at(args, 3)?, at(args, 4)?);
            let factor = optional(args, 5, 2.0)?;
            let no_switch = match args.get(6) {
                Some(arg) => arg.scalar().to_logical()?,
                None => false,
            };
            if cost < 0.0 || salvage < 0.0 || life <= 0.0 || start < 0.0 || end < start || end > life || factor <= 0.0 {
                return Err(ExcelError::Num);
            }
            finite(vdb(cost, salvage, life, start, end, factor, no_switch))
        }
        // Measured: REGEXTEST("ABC","abc") FALSE and with 1 TRUE; "(" #VALUE!.
        "REGEXTEST" => {
            count(args, 2, 3)?;
            let text: Vec<char> = args[0].scalar().to_text()?.chars().collect();
            Ok(Value::Logical(pattern(args, 2)?.find_at(&text, 0).is_some()))
        }
        // REGEXREPLACE(text, pattern, replacement, [occurrence], [case]):
        // every match, or the nth (from the end when negative); $n, ${n},
        // $0 and $$ in the replacement. Measured: "a1b2c3" with occurrence
        // 2 is a1b#c3, with -1 a1b2c#; "aaa" by a* is "--".
        "REGEXREPLACE" => {
            count(args, 3, 5)?;
            let text: Vec<char> = args[0].scalar().to_text()?.chars().collect();
            let regex = pattern(args, 4)?;
            let replacement = args[2].scalar().to_text()?;
            let occurrence = optional(args, 3, 0.0)?.trunc() as i64;
            let found = regex.find_all(&text);
            let chosen: Option<usize> = match occurrence {
                0 => None,
                n if n > 0 => Some(n as usize - 1),
                n => found.len().checked_sub(n.unsigned_abs() as usize),
            };
            if occurrence != 0 && chosen.is_none_or(|i| i >= found.len()) {
                return Ok(Value::Text(text.iter().collect()));
            }
            let mut out = String::new();
            let mut last = 0;
            for (i, caps) in found.iter().enumerate() {
                if chosen.is_some_and(|c| c != i) {
                    continue;
                }
                let (s, e) = caps[0].unwrap_or((last, last));
                out.extend(&text[last..s]);
                out.push_str(&crate::regex::expand(&replacement, &text, caps));
                last = e;
            }
            out.extend(&text[last..]);
            Ok(Value::Text(out))
        }
        // PERCENTOF: the one sum over the other. Measured: 1 of 1,2,3,4 is 0.1.
        "PERCENTOF" => {
            count(args, 2, 2)?;
            let (part, whole) = (numbers_of(&args[0])?, numbers_of(&args[1])?);
            let whole: f64 = whole.iter().sum();
            if whole == 0.0 {
                return Err(ExcelError::DivZero);
            }
            finite(part.iter().sum::<f64>() / whole)
        }
        // BAHTTEXT: the amount in Thai words, baht then satang rounded to
        // two places. Measured: 1234 หนึ่งพันสองร้อยสามสิบสี่บาทถ้วน, 1000001
        // หนึ่งล้านเอ็ดบาทถ้วน, 0.5 ห้าสิบสตางค์ (no baht), -5 ลบห้าบาทถ้วน,
        // 1.005 หนึ่งบาทหนึ่งสตางค์.
        "BAHTTEXT" => {
            count(args, 1, 1)?;
            let amount = at(args, 0)?;
            // Rounded as the decimal the number shows, so 1.005 is 1.01.
            let cents: f64 = format!("{:.13}", amount.abs() * 100.0).parse::<f64>().unwrap_or(0.0).round();
            if cents >= 1e19 {
                return Err(ExcelError::Num);
            }
            let cents = cents as u64;
            let (baht, satang) = (cents / 100, cents % 100);
            let mut out = String::new();
            if amount < 0.0 && cents > 0 {
                out.push_str("ลบ");
            }
            if baht > 0 || satang == 0 {
                out.push_str(if baht == 0 { "ศูนย์" } else { "" });
                out.push_str(&thai_words(baht, false));
                out.push_str("บาท");
            }
            if satang == 0 {
                out.push_str("ถ้วน");
            } else {
                out.push_str(&thai_block(satang, false));
                out.push_str("สตางค์");
            }
            Ok(Value::Text(out))
        }
        // ENCODEURL: UTF-8, every byte but A-Z a-z 0-9 - _ . written %XX.
        // Measured: "a b&c" is a%20b%26c and ~ becomes %7E.
        "ENCODEURL" => {
            count(args, 1, 1)?;
            let text = args[0].scalar().to_text()?;
            let mut out = String::with_capacity(text.len() * 3);
            for byte in text.bytes() {
                if byte.is_ascii_alphanumeric() || matches!(byte, b'-' | b'_' | b'.') {
                    out.push(byte as char);
                } else {
                    out.push_str(&format!("%{byte:02X}"));
                }
            }
            Ok(Value::Text(out))
        }
        // PHONETIC reads the reading stored with a cell; a cell with none
        // shows its text, and text handed over directly is #VALUE!.
        "PHONETIC" => {
            count(args, 1, 1)?;
            match &args[0] {
                Arg::Range(block) => Ok(Value::Text(block.cells.first().cloned().unwrap_or(Value::Blank).to_text()?)),
                Arg::Value(_) => Err(ExcelError::Value),
            }
        }
        // The French depreciation of the analysis add-in: AMORLINC straight
        // line, AMORDEGRC declining with the rate raised by the life's
        // coefficient and each year rounded. Measured: 360 and 776.
        "AMORLINC" | "AMORDEGRC" => {
            count(args, 6, 7)?;
            let (cost, bought, first_end, salvage) = (at(args, 0)?, at(args, 1)?.trunc(), at(args, 2)?.trunc(), at(args, 3)?);
            let (period, rate) = (at(args, 4)?.trunc(), at(args, 5)?);
            let basis = basis(args, 6)?;
            if basis == 2 || rate <= 0.0 || cost < 0.0 || salvage < 0.0 || period < 0.0 || bought > first_end {
                return Err(ExcelError::Num);
            }
            let first = yearfrac(bought as i64, first_end as i64, basis)?;
            if name == "AMORLINC" {
                let one = cost * rate;
                let first_part = first * rate * cost;
                let full = ((cost - salvage - first_part) / one).trunc();
                let answer = if period == 0.0 {
                    first_part
                } else if period <= full {
                    one
                } else if period == full + 1.0 {
                    cost - salvage - one * full - first_part
                } else {
                    0.0
                };
                return finite(answer.max(0.0));
            }
            let life = 1.0 / rate;
            let coefficient = if life < 3.0 {
                1.0
            } else if life < 5.0 {
                1.5
            } else if life <= 6.0 {
                2.0
            } else {
                2.5
            };
            let rate = rate * coefficient;
            let round = |x: f64| x.round();
            let mut this = round(first * rate * cost);
            let mut left = cost - this;
            let mut rest = left - salvage;
            let mut n = 0.0;
            while n < period {
                this = round(rate * left);
                rest -= this;
                if rest < 0.0 {
                    return finite(if period - n <= 1.0 { round(left * 0.5) } else { 0.0 });
                }
                left -= this;
                n += 1.0;
            }
            finite(this)
        }
        _ => Err(ExcelError::Name),
    }
}

/// The functions here that answer with a block.
pub(crate) fn call_block(name: &str, args: &[Arg]) -> Option<Arg> {
    let answer = match name {
        "MINVERSE" => (|| {
            if args.len() != 1 {
                return Err(ExcelError::Value);
            }
            let matrix = square(&args[0])?;
            inverse(matrix)
        })(),
        "MUNIT" => (|| {
            if args.len() != 1 {
                return Err(ExcelError::Value);
            }
            let size = at(args, 0)?.trunc();
            if size < 1.0 {
                return Err(ExcelError::Value);
            }
            let size = size as usize;
            let cells = (0..size * size)
                .map(|i| Value::Number(if i / size == i % size { 1.0 } else { 0.0 }))
                .collect();
            Ok(RangeData { width: size, height: size, cells })
        })(),
        "LINEST" | "LOGEST" => regression_block(name == "LOGEST", args),
        // REGEXEXTRACT(text, pattern, [mode], [case]): 0 the first match,
        // 1 every match, 2 the groups of the first -- the last two as a row.
        // Measured: no match is #N/A; a group that took no part is empty.
        "REGEXEXTRACT" => {
            return Some((|| -> Result<Arg, ExcelError> {
                count(args, 2, 4)?;
                let text: Vec<char> = args[0].scalar().to_text()?.chars().collect();
                let regex = pattern(args, 3)?;
                let piece = |span: Option<(usize, usize)>| {
                    Value::Text(span.map(|(s, e)| text[s..e].iter().collect()).unwrap_or_default())
                };
                match optional(args, 2, 0.0)?.trunc() as i64 {
                    0 => {
                        let caps = regex.find_at(&text, 0).ok_or(ExcelError::NA)?;
                        Ok(Arg::Value(piece(caps[0])))
                    }
                    1 => {
                        let all = regex.find_all(&text);
                        if all.is_empty() {
                            return Err(ExcelError::NA);
                        }
                        let cells: Vec<Value> = all.iter().map(|caps| piece(caps[0])).collect();
                        Ok(Arg::Range(RangeData { width: cells.len(), height: 1, cells }))
                    }
                    2 => {
                        let caps = regex.find_at(&text, 0).ok_or(ExcelError::NA)?;
                        let cells: Vec<Value> = if caps.len() > 1 { caps[1..].iter().map(|c| piece(*c)).collect() } else { vec![piece(caps[0])] };
                        Ok(Arg::Range(RangeData { width: cells.len(), height: 1, cells }))
                    }
                    _ => Err(ExcelError::Value),
                }
            })()
            .unwrap_or_else(|why| Arg::Value(Value::Error(why))));
        }
        "GROWTH" => growth(args),
        _ => return None,
    };
    Some(match answer {
        Ok(block) => Arg::Range(block),
        Err(why) => Arg::Value(Value::Error(why)),
    })
}

fn binomial(n: f64, k: f64) -> f64 {
    let k = k.min(n - k);
    let mut acc = 1.0f64;
    let mut i = 0.0;
    while i < k {
        acc = acc * (n - i) / (i + 1.0);
        i += 1.0;
    }
    acc.round()
}

/// A square block of numbers, row by row.
fn square(arg: &Arg) -> Result<Vec<Vec<f64>>, ExcelError> {
    let block = block_of(arg);
    if block.width != block.height || block.cells.is_empty() {
        return Err(ExcelError::Value);
    }
    let mut rows = Vec::with_capacity(block.height);
    for row in 0..block.height {
        let mut line = Vec::with_capacity(block.width);
        for col in 0..block.width {
            match reach(&block, col, row) {
                Some(Value::Number(n)) => line.push(n),
                _ => return Err(ExcelError::Value),
            }
        }
        rows.push(line);
    }
    Ok(rows)
}

/// The determinant by elimination with partial pivoting.
fn determinant(mut m: Vec<Vec<f64>>) -> f64 {
    let n = m.len();
    let mut det = 1.0;
    for col in 0..n {
        let pivot = (col..n)
            .max_by(|&a, &b| m[a][col].abs().partial_cmp(&m[b][col].abs()).unwrap_or(std::cmp::Ordering::Equal))
            .unwrap_or(col);
        if m[pivot][col] == 0.0 {
            return 0.0;
        }
        if pivot != col {
            m.swap(pivot, col);
            det = -det;
        }
        det *= m[col][col];
        for row in col + 1..n {
            let ratio = m[row][col] / m[col][col];
            for k in col..n {
                m[row][k] -= ratio * m[col][k];
            }
        }
    }
    det
}

/// The inverse by Gauss-Jordan elimination; #NUM! for a singular matrix.
fn inverse(mut m: Vec<Vec<f64>>) -> Result<RangeData, ExcelError> {
    let n = m.len();
    let mut out: Vec<Vec<f64>> = (0..n).map(|i| (0..n).map(|j| if i == j { 1.0 } else { 0.0 }).collect()).collect();
    for col in 0..n {
        let pivot = (col..n)
            .max_by(|&a, &b| m[a][col].abs().partial_cmp(&m[b][col].abs()).unwrap_or(std::cmp::Ordering::Equal))
            .unwrap_or(col);
        if m[pivot][col] == 0.0 {
            return Err(ExcelError::Num);
        }
        m.swap(pivot, col);
        out.swap(pivot, col);
        let lead = m[col][col];
        for k in 0..n {
            m[col][k] /= lead;
            out[col][k] /= lead;
        }
        for row in 0..n {
            if row != col {
                let ratio = m[row][col];
                if ratio != 0.0 {
                    for k in 0..n {
                        m[row][k] -= ratio * m[col][k];
                        out[row][k] -= ratio * out[col][k];
                    }
                }
            }
        }
    }
    Ok(RangeData { width: n, height: n, cells: out.into_iter().flatten().map(Value::Number).collect() })
}

fn flows_and_dates(flows: &Arg, dates: &Arg) -> Result<(Vec<f64>, Vec<f64>), ExcelError> {
    let (flows, dates) = (flows.flatten(), dates.flatten());
    if flows.len() != dates.len() || flows.is_empty() {
        return Err(ExcelError::Num);
    }
    let mut out_flows = Vec::with_capacity(flows.len());
    let mut out_dates = Vec::with_capacity(dates.len());
    for (flow, date) in flows.iter().zip(&dates) {
        let flow = match flow {
            Value::Number(n) => *n,
            Value::Error(e) => return Err(*e),
            _ => return Err(ExcelError::Value),
        };
        let date = match date {
            Value::Number(n) => n.trunc(),
            Value::Error(e) => return Err(*e),
            _ => return Err(ExcelError::Value),
        };
        out_flows.push(flow);
        out_dates.push(date);
    }
    if out_dates.iter().any(|date| *date < out_dates[0]) {
        return Err(ExcelError::Num);
    }
    Ok((out_flows, out_dates))
}

fn xnpv(rate: f64, flows: &[f64], dates: &[f64]) -> f64 {
    flows
        .iter()
        .zip(dates)
        .map(|(flow, date)| flow / (1.0 + rate).powf((date - dates[0]) / 365.0))
        .sum()
}

/// Declining balance for one period, never below the salvage.
fn period_ddb(cost: f64, salvage: f64, life: f64, period: f64, factor: f64) -> f64 {
    let mut rate = factor / life;
    let old = if rate >= 1.0 {
        rate = 1.0;
        if period == 1.0 { cost } else { 0.0 }
    } else {
        cost * (1.0 - rate).powf(period - 1.0)
    };
    let new = cost * (1.0 - rate).powf(period);
    let ddb = if new < salvage { old - salvage } else { old - new };
    ddb.max(0.0)
}

/// Declining balance switching to straight line once that gives more,
/// over the first `period` periods of what is left.
fn inter_vdb(cost: f64, salvage: f64, life: f64, life_left: f64, period: f64, factor: f64) -> f64 {
    let end = period.ceil();
    let last = end as u64;
    let mut total = 0.0;
    let mut left_to_write = cost - salvage;
    let mut straight = false;
    let mut line = 0.0;
    for i in 1..=last {
        let mut term = if !straight {
            let ddb = period_ddb(cost, salvage, life, i as f64, factor);
            line = left_to_write / (life_left - (i - 1) as f64);
            if line > ddb {
                straight = true;
                line
            } else {
                left_to_write -= ddb;
                ddb
            }
        } else {
            line
        };
        if i == last {
            term *= period + 1.0 - end;
        }
        total += term;
    }
    total
}

/// VDB, worked the way LibreOffice works it: VDB(2400,300,10,0,1) 480.
fn vdb(mut cost: f64, salvage: f64, mut life: f64, mut start: f64, mut end: f64, factor: f64, no_switch: bool) -> f64 {
    if no_switch {
        let (first, last) = (start.floor(), end.ceil());
        let mut total = 0.0;
        let mut i = first + 1.0;
        while i <= last {
            let mut term = period_ddb(cost, salvage, life, i, factor);
            if i == first + 1.0 {
                term *= end.min(first + 1.0) - start;
            } else if i == last {
                term *= end + 1.0 - last;
            }
            total += term;
            i += 1.0;
        }
        return total;
    }
    if start != start.floor() && factor > 1.0 && start >= life / 2.0 {
        let part = start - life / 2.0;
        start = life / 2.0;
        end -= part;
        life += 1.0;
    }
    cost -= inter_vdb(cost, salvage, life, life, start, factor);
    inter_vdb(cost, salvage, life, life - start, end - start, factor)
}

/// The known y values and x values of LINEST, LOGEST and GROWTH: the x values
/// as one row of variables per observation. A y column with an x block as
/// tall has a variable per x column; a y row, one per x row; x values the
/// same size as the y values are one variable; no x values are 1, 2, 3, ...
fn observations(args: &[Arg], logs: bool) -> Result<(Vec<f64>, Vec<Vec<f64>>), ExcelError> {
    let given = |i: usize| match args.get(i) {
        Some(Arg::Value(Value::Blank)) | None => None,
        Some(arg) => Some(arg),
    };
    let y_block = block_of(&args[0]);
    let mut ys = Vec::with_capacity(y_block.cells.len());
    for value in &y_block.cells {
        let y = match value {
            Value::Number(n) => *n,
            Value::Error(e) => return Err(*e),
            _ => return Err(ExcelError::Value),
        };
        if logs && y <= 0.0 {
            return Err(ExcelError::Num);
        }
        ys.push(if logs { y.ln() } else { y });
    }
    let n = ys.len();
    let x_block = match given(1) {
        Some(arg) => block_of(arg),
        None => RangeData { width: 1, height: n, cells: (1..=n).map(|i| Value::Number(i as f64)).collect() },
    };
    let number = |value: &Value| match value {
        Value::Number(n) => Ok(*n),
        Value::Error(e) => Err(*e),
        _ => Err(ExcelError::Value),
    };
    let mut rows = Vec::with_capacity(n);
    if x_block.cells.len() == n {
        for value in &x_block.cells {
            rows.push(vec![number(value)?]);
        }
    } else if y_block.width == 1 && x_block.height == n {
        for row in 0..n {
            rows.push((0..x_block.width).map(|col| number(&x_block.at(col, row))).collect::<Result<_, _>>()?);
        }
    } else if y_block.height == 1 && x_block.width == n {
        for col in 0..n {
            rows.push((0..x_block.height).map(|row| number(&x_block.at(col, row))).collect::<Result<_, _>>()?);
        }
    } else {
        return Err(ExcelError::Ref);
    }
    Ok((ys, rows))
}

type Fit = (Vec<f64>, f64, Vec<Vec<f64>>, Vec<f64>);

/// Least squares: the slopes, the constant, the inverse of the normal
/// matrix and the x means, the data centred first when there is a constant.
fn least_squares(ys: &[f64], rows: &[Vec<f64>], constant: bool) -> Result<Fit, ExcelError> {
    let n = ys.len();
    let k = rows.first().map_or(0, Vec::len);
    let (x_mean, y_mean) = if constant {
        let x_mean: Vec<f64> = (0..k).map(|j| rows.iter().map(|r| r[j]).sum::<f64>() / n as f64).collect();
        (x_mean, ys.iter().sum::<f64>() / n as f64)
    } else {
        (vec![0.0; k], 0.0)
    };
    let mut normal = vec![vec![0.0; k]; k];
    let mut right = vec![0.0; k];
    for (row, y) in rows.iter().zip(ys) {
        for a in 0..k {
            let xa = row[a] - x_mean[a];
            right[a] += xa * (y - y_mean);
            for b in 0..k {
                normal[a][b] += xa * (row[b] - x_mean[b]);
            }
        }
    }
    let first_square = normal.first().and_then(|row| row.first()).copied().unwrap_or(0.0);
    let inverted = inverse(normal).map_err(|_| ExcelError::Num)?;
    let matrix: Vec<Vec<f64>> = (0..k)
        .map(|a| {
            (0..k)
                .map(|b| match &inverted.cells[a * k + b] {
                    Value::Number(v) => *v,
                    _ => 0.0,
                })
                .collect()
        })
        .collect();
    // One variable is worked as TREND works it, so the two agree.
    let slopes: Vec<f64> = if k == 1 {
        vec![right[0] / first_square]
    } else {
        (0..k).map(|a| (0..k).map(|b| matrix[a][b] * right[b]).sum()).collect()
    };
    let intercept = if constant { y_mean - (0..k).map(|j| slopes[j] * x_mean[j]).sum::<f64>() } else { 0.0 };
    Ok((slopes, intercept, matrix, x_mean))
}

/// LINEST and LOGEST: the slopes last variable first, then the constant;
/// with statistics, four rows more as Excel lays them out, #N/A where a row
/// has nothing more to say.
fn regression_block(logs: bool, args: &[Arg]) -> Result<RangeData, ExcelError> {
    if args.is_empty() || args.len() > 4 {
        return Err(ExcelError::Value);
    }
    let flag = |i: usize, default: bool| match args.get(i).map(Arg::scalar) {
        None | Some(Value::Blank) => Ok(default),
        Some(value) => value.to_logical(),
    };
    let constant = flag(2, true)?;
    let stats = flag(3, false)?;
    let (ys, rows) = observations(args, logs)?;
    let n = ys.len();
    let k = rows.first().map_or(0, Vec::len);
    let (slopes, intercept, matrix, x_mean) = least_squares(&ys, &rows, constant)?;
    let shown = |v: f64| Value::Number(if logs { v.exp() } else { v });
    let width = k + 1;
    let mut cells: Vec<Value> = slopes.iter().rev().map(|m| shown(*m)).collect();
    cells.push(shown(intercept));
    if !stats {
        return Ok(RangeData { width, height: 1, cells });
    }
    let fitted: Vec<f64> =
        rows.iter().map(|r| intercept + r.iter().zip(&slopes).map(|(x, m)| x * m).sum::<f64>()).collect();
    let y_mean = ys.iter().sum::<f64>() / n as f64;
    let ss_resid: f64 = ys.iter().zip(&fitted).map(|(y, f)| (y - f).powi(2)).sum();
    let ss_total: f64 =
        if constant { ys.iter().map(|y| (y - y_mean).powi(2)).sum() } else { ys.iter().map(|y| y * y).sum() };
    let ss_reg = ss_total - ss_resid;
    let freedom = n as f64 - k as f64 - if constant { 1.0 } else { 0.0 };
    let na = Value::Error(ExcelError::NA);
    let variance = if freedom > 0.0 { ss_resid / freedom } else { f64::NAN };
    let mut errors: Vec<Value> = (0..k).rev().map(|j| Value::Number((matrix[j][j] * variance).sqrt())).collect();
    if constant {
        let mut extra = 1.0 / n as f64;
        for a in 0..k {
            for b in 0..k {
                extra += x_mean[a] * matrix[a][b] * x_mean[b];
            }
        }
        errors.push(Value::Number((extra * variance).sqrt()));
    } else {
        errors.push(na.clone());
    }
    cells.extend(errors);
    for row in [
        // r squared as one less the unexplained share: measured, LOGEST's
        // 0.720442481699332 comes out that way and not as ss_reg/ss_total.
        vec![Value::Number(1.0 - ss_resid / ss_total), Value::Number(variance.sqrt())],
        vec![Value::Number((ss_reg / k as f64) / variance), Value::Number(freedom)],
        vec![Value::Number(ss_reg), Value::Number(ss_resid)],
    ] {
        let mut row = row;
        while row.len() < width {
            row.push(na.clone());
        }
        cells.extend(row.into_iter().take(width));
    }
    Ok(RangeData { width, height: 5, cells })
}

/// TREND with several x variables, the one-variable case being left to
/// TREND itself.
pub(crate) fn trend_many(args: &[Arg]) -> Option<Result<RangeData, ExcelError>> {
    let x = args.get(1).filter(|arg| !matches!(arg, Arg::Value(Value::Blank)))?;
    let (y, x) = (block_of(&args[0]), block_of(x));
    if x.cells.len() == y.cells.len() {
        return None;
    }
    Some((|| {
        let constant = match args.get(3).map(Arg::scalar) {
            None | Some(Value::Blank) => true,
            Some(value) => value.to_logical()?,
        };
        let (ys, rows) = observations(args, false)?;
        let (slopes, intercept, _, _) = least_squares(&ys, &rows, constant)?;
        let k = slopes.len();
        let new_block = match args.get(2).filter(|arg| !matches!(arg, Arg::Value(Value::Blank))) {
            Some(arg) => block_of(arg),
            None => x.clone(),
        };
        let number = |value: &Value| match value {
            Value::Number(n) => Ok(*n),
            Value::Error(e) => Err(*e),
            _ => Err(ExcelError::Value),
        };
        let (points, across) = if new_block.width == k { (new_block.height, true) } else { (new_block.width, false) };
        let mut cells = Vec::with_capacity(points);
        for p in 0..points {
            let mut fitted = intercept;
            for (j, slope) in slopes.iter().enumerate() {
                let value = if across { new_block.at(j, p) } else { new_block.at(p, j) };
                fitted += slope * number(&value)?;
            }
            cells.push(Value::Number(fitted));
        }
        Ok(if across {
            RangeData { width: 1, height: points, cells }
        } else {
            RangeData { width: points, height: 1, cells }
        })
    })())
}

/// GROWTH: the exponential curve through the known points, at the new x values.
fn growth(args: &[Arg]) -> Result<RangeData, ExcelError> {
    if args.is_empty() || args.len() > 4 {
        return Err(ExcelError::Value);
    }
    let constant = match args.get(3).map(Arg::scalar) {
        None | Some(Value::Blank) => true,
        Some(value) => value.to_logical()?,
    };
    let (ys, rows) = observations(args, true)?;
    let (slopes, intercept, _, _) = least_squares(&ys, &rows, constant)?;
    let k = slopes.len();
    let blank = |i: usize| matches!(args.get(i), Some(Arg::Value(Value::Blank)) | None);
    let new_block = if !blank(2) {
        block_of(&args[2])
    } else if !blank(1) {
        block_of(&args[1])
    } else {
        let y = block_of(&args[0]);
        RangeData { width: y.width, height: y.height, cells: (1..=y.cells.len()).map(|i| Value::Number(i as f64)).collect() }
    };
    let number = |value: &Value| match value {
        Value::Number(n) => Ok(*n),
        Value::Error(e) => Err(*e),
        _ => Err(ExcelError::Value),
    };
    if k == 1 {
        let cells = new_block
            .cells
            .iter()
            .map(|x| number(x).map(|x| Value::Number((intercept + slopes[0] * x).exp())))
            .collect::<Result<Vec<_>, _>>()?;
        return Ok(RangeData { width: new_block.width, height: new_block.height, cells });
    }
    let (points, across) = if new_block.width == k { (new_block.height, true) } else { (new_block.width, false) };
    let mut cells = Vec::with_capacity(points);
    for p in 0..points {
        // b * m1^x1 * m2^x2 ..., as LOGEST states the curve: measured, the
        // last digit of GROWTH over two variables comes out that way.
        let mut fitted = intercept.exp();
        for (j, slope) in slopes.iter().enumerate() {
            let x = if across { new_block.at(j, p) } else { new_block.at(p, j) };
            fitted *= slope.exp().powf(number(&x)?);
        }
        cells.push(Value::Number(fitted));
    }
    Ok(if across {
        RangeData { width: 1, height: points, cells }
    } else {
        RangeData { width: points, height: 1, cells }
    })
}
