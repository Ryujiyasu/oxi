// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! The engineering functions on complex numbers written as text: COMPLEX,
//! IMREAL, IMSUM, IMSQRT and the rest. Each answer quoted here was read
//! from Excel.

use crate::functions::{num, Arg};
use crate::value::{ExcelError, Value};

pub(crate) const NAMES: &[&str] = &[
    "COMPLEX", "IMABS", "IMAGINARY", "IMARGUMENT", "IMCONJUGATE", "IMCOS", "IMCOSH", "IMCOT", "IMCSC",
    "IMCSCH", "IMDIV", "IMEXP", "IMLN", "IMLOG10", "IMLOG2", "IMPOWER", "IMPRODUCT", "IMREAL", "IMSEC",
    "IMSECH", "IMSIN", "IMSINH", "IMSQRT", "IMSUB", "IMSUM", "IMTAN",
];

#[derive(Clone, Copy)]
struct Complex {
    re: f64,
    im: f64,
}

impl Complex {
    fn new(re: f64, im: f64) -> Self {
        Complex { re, im }
    }
    fn mul(self, o: Complex) -> Complex {
        Complex::new(self.re * o.re - self.im * o.im, self.re * o.im + self.im * o.re)
    }
    fn div(self, o: Complex) -> Result<Complex, ExcelError> {
        let d = o.re * o.re + o.im * o.im;
        if d == 0.0 {
            return Err(ExcelError::Num);
        }
        Ok(Complex::new((self.re * o.re + self.im * o.im) / d, (self.im * o.re - self.re * o.im) / d))
    }
    fn abs(self) -> f64 {
        (self.re * self.re + self.im * self.im).sqrt()
    }
    fn arg(self) -> f64 {
        self.im.atan2(self.re)
    }
    fn ln(self) -> Result<Complex, ExcelError> {
        if self.re == 0.0 && self.im == 0.0 {
            return Err(ExcelError::Num);
        }
        Ok(Complex::new(self.abs().ln(), self.arg()))
    }
    fn exp(self) -> Complex {
        let e = self.re.exp();
        Complex::new(e * self.im.cos(), e * self.im.sin())
    }
    fn sin(self) -> Complex {
        Complex::new(self.re.sin() * self.im.cosh(), self.re.cos() * self.im.sinh())
    }
    fn cos(self) -> Complex {
        Complex::new(self.re.cos() * self.im.cosh(), -self.re.sin() * self.im.sinh())
    }
    fn sinh(self) -> Complex {
        Complex::new(self.re.sinh() * self.im.cos(), self.re.cosh() * self.im.sin())
    }
    fn cosh(self) -> Complex {
        Complex::new(self.re.cosh() * self.im.cos(), self.re.sinh() * self.im.sin())
    }
    /// z to a real power, by the polar form.
    fn powf(self, n: f64) -> Result<Complex, ExcelError> {
        let r = self.abs();
        if r == 0.0 {
            return if n > 0.0 { Ok(Complex::new(0.0, 0.0)) } else { Err(ExcelError::Num) };
        }
        let (rn, t) = (r.powf(n), self.arg() * n);
        Ok(Complex::new(rn * t.cos(), rn * t.sin()))
    }
}

/// A complex number read from text: `3`, `4i`, `3-4i`, `-j`, `1e2+1E-1i`.
/// Measured: "" and "1+2I" and "2i+1" are #NUM!; a leading space is let
/// through. Also the letter it was written with, None for a plain number.
fn parse(arg: &Arg) -> Result<(Complex, Option<char>), ExcelError> {
    let text = match arg.scalar() {
        Value::Number(n) => return Ok((Complex::new(n, 0.0), None)),
        Value::Logical(_) => return Err(ExcelError::Value),
        Value::Error(e) => return Err(e),
        Value::Blank => return Ok((Complex::new(0.0, 0.0), None)),
        Value::Text(t) => t,
    };
    let text = text.trim_start();
    if text.is_empty() {
        return Err(ExcelError::Num);
    }
    let number = |part: &str| -> Result<f64, ExcelError> {
        if part.is_empty() || part.contains(|c: char| !(c.is_ascii_digit() || matches!(c, '.' | 'e' | 'E' | '+' | '-'))) {
            return Err(ExcelError::Num);
        }
        part.parse::<f64>().map_err(|_| ExcelError::Num)
    };
    let last = text.chars().last().unwrap_or(' ');
    if last != 'i' && last != 'j' {
        return Ok((Complex::new(number(text)?, 0.0), None));
    }
    let body = &text[..text.len() - 1];
    // The sign that starts the imaginary part: the last + or - that is not
    // the first character and does not follow an exponent's e.
    let bytes = body.as_bytes();
    let split = (1..bytes.len())
        .rev()
        .find(|&at| matches!(bytes[at], b'+' | b'-') && !matches!(bytes[at - 1], b'e' | b'E'));
    let (real, imaginary) = match split {
        Some(at) => (number(&body[..at])?, &body[at..]),
        None => (0.0, body),
    };
    let imaginary = match imaginary {
        "" | "+" => 1.0,
        "-" => -1.0,
        other => number(other)?,
    };
    Ok((Complex::new(real, imaginary), Some(last)))
}

/// A number the way these functions write one: fifteen significant digits,
/// in full unless that runs past 21 characters, then as 1.2E-07.
/// Measured: 0.000000001, 123456789012345000, 100000000000000000000, but
/// 1.23456789012345E-07 and 1E+100.
fn written(x: f64) -> String {
    if x == 0.0 {
        return "0".to_string();
    }
    let scientific = format!("{:.14e}", x.abs());
    let (mantissa, exponent) = scientific.split_once('e').unwrap_or((&scientific, "0"));
    let exponent: i32 = exponent.parse().unwrap_or(0);
    let digits: String = mantissa.chars().filter(|c| c.is_ascii_digit()).collect();
    let digits = digits.trim_end_matches('0');
    let digits = if digits.is_empty() { "0" } else { digits };
    let fixed = if exponent >= 0 {
        let whole = exponent as usize + 1;
        if digits.len() <= whole {
            format!("{digits}{}", "0".repeat(whole - digits.len()))
        } else {
            format!("{}.{}", &digits[..whole], &digits[whole..])
        }
    } else {
        format!("0.{}{digits}", "0".repeat((-exponent - 1) as usize))
    };
    let body = if fixed.len() <= 21 {
        fixed
    } else {
        let mantissa = if digits.len() > 1 { format!("{}.{}", &digits[..1], &digits[1..]) } else { digits.to_string() };
        format!("{mantissa}E{}{:02}", if exponent < 0 { '-' } else { '+' }, exponent.abs())
    };
    if x < 0.0 {
        format!("-{body}")
    } else {
        body
    }
}

fn text_of(z: Complex, suffix: char) -> Value {
    // Rounded to what will be written, so 0.30000000000000004 reads 0.3 and
    // a part that rounds to nothing is left out.
    let (re, im) = (written(z.re), written(z.im));
    let imaginary = match im.as_str() {
        "1" => suffix.to_string(),
        "-1" => format!("-{suffix}"),
        other => format!("{other}{suffix}"),
    };
    Value::Text(if im == "0" {
        re
    } else if re == "0" {
        imaginary
    } else if imaginary.starts_with('-') {
        format!("{re}{imaginary}")
    } else {
        format!("{re}+{imaginary}")
    })
}

/// The letter the answer is written with: j when the arguments used j, i
/// otherwise; #VALUE! when they mix the two.
fn letter(found: &[Option<char>]) -> Result<char, ExcelError> {
    let mut chosen = None;
    for one in found.iter().flatten() {
        match chosen {
            None => chosen = Some(*one),
            Some(held) if held != *one => return Err(ExcelError::Value),
            _ => {}
        }
    }
    Ok(chosen.unwrap_or('i'))
}

fn finite(z: Complex) -> Result<Complex, ExcelError> {
    if z.re.is_finite() && z.im.is_finite() {
        Ok(z)
    } else {
        Err(ExcelError::Num)
    }
}

pub(crate) fn call(name: &str, args: &[Arg]) -> Result<Value, ExcelError> {
    let count = |low: usize, high: usize| {
        if args.len() < low || args.len() > high {
            Err(ExcelError::Value)
        } else {
            Ok(())
        }
    };
    match name {
        "COMPLEX" => {
            count(2, 3)?;
            let (re, im) = (num(&args[0])?, num(&args[1])?);
            let suffix = match args.get(2).map(Arg::scalar) {
                None | Some(Value::Blank) => 'i',
                Some(Value::Text(t)) if t == "i" || t.is_empty() => 'i',
                Some(Value::Text(t)) if t == "j" => 'j',
                Some(_) => return Err(ExcelError::Value),
            };
            Ok(text_of(Complex::new(re, im), suffix))
        }
        "IMSUM" | "IMPRODUCT" => {
            if args.is_empty() {
                return Err(ExcelError::Value);
            }
            let mut found = Vec::new();
            let mut total = if name == "IMSUM" { Complex::new(0.0, 0.0) } else { Complex::new(1.0, 0.0) };
            for arg in args {
                let values: Vec<Arg> = match arg {
                    Arg::Range(block) => block.cells.iter().map(|v| Arg::Value(v.clone())).collect(),
                    other => vec![other.clone()],
                };
                for one in values {
                    if matches!(one.scalar(), Value::Blank) && matches!(arg, Arg::Range(_)) {
                        continue;
                    }
                    let (z, suffix) = parse(&one)?;
                    found.push(suffix);
                    total = if name == "IMSUM" { Complex::new(total.re + z.re, total.im + z.im) } else { total.mul(z) };
                }
            }
            let suffix = letter(&found)?;
            Ok(text_of(finite(total)?, suffix))
        }
        "IMSUB" | "IMDIV" => {
            count(2, 2)?;
            let ((a, s1), (b, s2)) = (parse(&args[0])?, parse(&args[1])?);
            let suffix = letter(&[s1, s2])?;
            let z = if name == "IMSUB" { Complex::new(a.re - b.re, a.im - b.im) } else { a.div(b)? };
            Ok(text_of(finite(z)?, suffix))
        }
        "IMREAL" | "IMAGINARY" | "IMABS" | "IMARGUMENT" => {
            count(1, 1)?;
            let (z, _) = parse(&args[0])?;
            Ok(Value::Number(match name {
                "IMREAL" => z.re,
                "IMAGINARY" => z.im,
                "IMABS" => z.abs(),
                _ => {
                    if z.re == 0.0 && z.im == 0.0 {
                        return Err(ExcelError::DivZero);
                    }
                    z.arg()
                }
            }))
        }
        "IMPOWER" => {
            count(2, 2)?;
            let (z, suffix) = parse(&args[0])?;
            let n = num(&args[1])?;
            Ok(text_of(finite(z.powf(n)?)?, suffix.unwrap_or('i')))
        }
        _ => {
            count(1, 1)?;
            let (z, suffix) = parse(&args[0])?;
            let one = Complex::new(1.0, 0.0);
            let answer = match name {
                "IMCONJUGATE" => Complex::new(z.re, -z.im),
                "IMSQRT" => z.powf(0.5)?,
                "IMEXP" => z.exp(),
                "IMLN" => z.ln()?,
                "IMLOG10" | "IMLOG2" => {
                    let base = if name == "IMLOG10" { 10f64.ln() } else { 2f64.ln() };
                    let l = z.ln()?;
                    Complex::new(l.re / base, l.im / base)
                }
                "IMSIN" => z.sin(),
                "IMCOS" => z.cos(),
                "IMTAN" => z.sin().div(z.cos())?,
                "IMSINH" => z.sinh(),
                "IMCOSH" => z.cosh(),
                "IMSEC" => one.div(z.cos())?,
                "IMCSC" => one.div(z.sin())?,
                "IMCOT" => z.cos().div(z.sin())?,
                "IMSECH" => one.div(z.cosh())?,
                "IMCSCH" => one.div(z.sinh())?,
                _ => return Err(ExcelError::Name),
            };
            Ok(text_of(finite(answer)?, suffix.unwrap_or('i')))
        }
    }
}
