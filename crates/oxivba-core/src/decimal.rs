// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! VBA's `Decimal`: a 96-bit whole number and a power of ten, up to 28, to
//! divide it by.
//!
//! Every behaviour here was measured against Excel's VBA: `CDec(1) / 3` is
//! 0.3333333333333333333333333333 (28 places), `CDec(1) / 3 * 3` is
//! 0.9999999999999999999999999999, the largest is
//! 79228162514264337593543950335, a half goes to the even neighbour
//! (`Round(CDec("2.345"), 2)` is 2.34), and a Double becomes one through its
//! fifteen significant digits (`CDec(0.1) + CDec(0.2) = 0.3`).

use std::cmp::Ordering;

/// The largest magnitude a Decimal holds: 2^96 - 1.
const MAX: u128 = (1u128 << 96) - 1;
const MAX_SCALE: u32 = 28;

/// A Decimal: sign, magnitude (at most `MAX`) and scale (at most 28).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Dec {
    pub negative: bool,
    pub magnitude: u128,
    pub scale: u8,
}

/// A number too large for a Decimal.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Overflow;

/// An unsigned whole number of up to 256 bits, little-endian 64-bit limbs:
/// room for the product of two magnitudes and for one scaled up by 10^28.
#[derive(Clone, Copy, PartialEq, Eq)]
struct Wide([u64; 4]);

impl Wide {
    fn from_u128(value: u128) -> Wide {
        Wide([value as u64, (value >> 64) as u64, 0, 0])
    }

    fn is_zero(&self) -> bool {
        self.0.iter().all(|limb| *limb == 0)
    }

    fn fits(&self) -> Option<u128> {
        if self.0[2] != 0 || self.0[3] != 0 {
            return None;
        }
        let value = (self.0[0] as u128) | ((self.0[1] as u128) << 64);
        (value <= MAX).then_some(value)
    }

    fn mul_small(&self, factor: u64) -> Wide {
        let mut out = [0u64; 4];
        let mut carry = 0u128;
        for (at, limb) in self.0.iter().enumerate() {
            let product = (*limb as u128) * (factor as u128) + carry;
            out[at] = product as u64;
            carry = product >> 64;
        }
        Wide(out)
    }

    fn mul(a: u128, b: u128) -> Wide {
        let a = [a as u64, (a >> 64) as u64];
        let b = [b as u64, (b >> 64) as u64];
        let mut out = [0u64; 4];
        for (i, x) in a.iter().enumerate() {
            let mut carry = 0u128;
            for (j, y) in b.iter().enumerate() {
                let sum = (*x as u128) * (*y as u128) + out[i + j] as u128 + carry;
                out[i + j] = sum as u64;
                carry = sum >> 64;
            }
            let mut at = i + 2;
            while carry != 0 && at < 4 {
                let sum = out[at] as u128 + carry;
                out[at] = sum as u64;
                carry = sum >> 64;
                at += 1;
            }
        }
        Wide(out)
    }

    fn add(&self, other: &Wide) -> Wide {
        let mut out = [0u64; 4];
        let mut carry = 0u128;
        for at in 0..4 {
            let sum = self.0[at] as u128 + other.0[at] as u128 + carry;
            out[at] = sum as u64;
            carry = sum >> 64;
        }
        Wide(out)
    }

    /// `self - other`, with `self >= other`.
    fn sub(&self, other: &Wide) -> Wide {
        let mut out = [0u64; 4];
        let mut borrow = 0i128;
        for at in 0..4 {
            let difference = self.0[at] as i128 - other.0[at] as i128 - borrow;
            if difference < 0 {
                out[at] = (difference + (1i128 << 64)) as u64;
                borrow = 1;
            } else {
                out[at] = difference as u64;
                borrow = 0;
            }
        }
        Wide(out)
    }

    fn cmp(&self, other: &Wide) -> Ordering {
        for at in (0..4).rev() {
            match self.0[at].cmp(&other.0[at]) {
                Ordering::Equal => continue,
                unequal => return unequal,
            }
        }
        Ordering::Equal
    }

    /// Divide by a small number, handing back the remainder.
    fn div_small(&mut self, divisor: u64) -> u64 {
        let mut remainder = 0u128;
        for at in (0..4).rev() {
            let current = (remainder << 64) | self.0[at] as u128;
            self.0[at] = (current / divisor as u128) as u64;
            remainder = current % divisor as u128;
        }
        remainder as u64
    }
}

impl Dec {
    pub fn zero() -> Dec {
        Dec { negative: false, magnitude: 0, scale: 0 }
    }

    /// A whole number and scale brought into range: the scale is cut back,
    /// rounding half to even, until the magnitude fits and the scale is at
    /// most 28.
    fn settle(negative: bool, wide: Wide, scale: u32) -> Result<Dec, Overflow> {
        Dec::settle_beyond(negative, wide, scale, false)
    }

    /// `settle`, told whether anything was already left over below the
    /// digits it is given -- a division's remainder -- so that a cut digit
    /// of 5 is known to be more than half: measured, CDec(1) / CDec(7) ends
    /// ...429, not ...428.
    fn settle_beyond(negative: bool, mut wide: Wide, mut scale: u32, beyond: bool) -> Result<Dec, Overflow> {
        let mut last = 0u64;
        let mut sticky = beyond;
        let mut cut = false;
        loop {
            let fits = wide.fits();
            if fits.is_some() && scale <= MAX_SCALE {
                break;
            }
            if scale == 0 {
                return Err(Overflow);
            }
            sticky |= last != 0;
            last = wide.div_small(10);
            scale -= 1;
            cut = true;
        }
        if cut && (last > 5 || (last == 5 && (sticky || wide.0[0] & 1 == 1))) {
            wide = wide.add(&Wide::from_u128(1));
            if wide.fits().is_none() {
                // The carry made it one digit too long: cut once more.
                if scale == 0 {
                    return Err(Overflow);
                }
                let digit = wide.div_small(10);
                scale -= 1;
                if digit >= 5 {
                    wide = wide.add(&Wide::from_u128(1));
                }
                if wide.fits().is_none() {
                    return Err(Overflow);
                }
            }
        }
        let magnitude = wide.fits().ok_or(Overflow)?;
        Ok(Dec { negative: negative && magnitude != 0, magnitude, scale: scale as u8 })
    }

    pub fn from_i128(value: i128) -> Result<Dec, Overflow> {
        let magnitude = value.unsigned_abs();
        if magnitude > MAX {
            return Err(Overflow);
        }
        Ok(Dec { negative: value < 0, magnitude, scale: 0 })
    }

    /// Written the way Excel's VBA writes a Double it turns into a Decimal:
    /// through its fifteen significant digits.
    pub fn from_f64(value: f64) -> Result<Dec, Overflow> {
        if !value.is_finite() {
            return Err(Overflow);
        }
        if value == 0.0 {
            return Ok(Dec::zero());
        }
        Dec::parse(&format!("{:.14e}", value)).map(|held| held.reduced()).ok_or(Overflow)
    }

    /// The same value with no trailing zeros past the point.
    pub fn reduced(&self) -> Dec {
        let mut held = *self;
        while held.scale > 0 && held.magnitude % 10 == 0 {
            held.magnitude /= 10;
            held.scale -= 1;
        }
        held
    }

    /// Text such as `-12.5`, `1e3` or `.25`; None when it is not a number.
    pub fn parse(text: &str) -> Option<Dec> {
        let text = text.trim();
        let (negative, body) = match text.strip_prefix('-') {
            Some(rest) => (true, rest),
            None => (false, text.strip_prefix('+').unwrap_or(text)),
        };
        let (mantissa, exponent) = match body.find(['e', 'E']) {
            Some(at) => (&body[..at], body[at + 1..].parse::<i32>().ok()?),
            None => (body, 0),
        };
        let (whole, fraction) = mantissa.split_once('.').unwrap_or((mantissa, ""));
        if whole.is_empty() && fraction.is_empty() {
            return None;
        }
        if !whole.chars().chain(fraction.chars()).all(|ch| ch.is_ascii_digit()) {
            return None;
        }
        let digits: String = format!("{whole}{fraction}");
        let mut wide = Wide::from_u128(0);
        for digit in digits.chars() {
            wide = wide.mul_small(10).add(&Wide::from_u128(digit as u128 - '0' as u128));
            if wide.0[3] != 0 {
                return None;
            }
        }
        let mut scale = fraction.len() as i32 - exponent;
        while scale < 0 {
            wide = wide.mul_small(10);
            scale += 1;
        }
        Dec::settle(negative, wide, scale as u32).ok()
    }

    pub fn to_f64(&self) -> f64 {
        let value = self.magnitude as f64 / 10f64.powi(self.scale as i32);
        if self.negative { -value } else { value }
    }

    /// The magnitude brought to `scale` places, as a wide whole number.
    fn at_scale(&self, scale: u32) -> Wide {
        let mut wide = Wide::from_u128(self.magnitude);
        for _ in self.scale as u32..scale {
            wide = wide.mul_small(10);
        }
        wide
    }

    pub fn add(&self, other: &Dec) -> Result<Dec, Overflow> {
        let scale = self.scale.max(other.scale) as u32;
        let (a, b) = (self.at_scale(scale), other.at_scale(scale));
        if self.negative == other.negative {
            return Dec::settle(self.negative, a.add(&b), scale);
        }
        match a.cmp(&b) {
            Ordering::Less => Dec::settle(other.negative, b.sub(&a), scale),
            _ => Dec::settle(self.negative, a.sub(&b), scale),
        }
    }

    pub fn neg(&self) -> Dec {
        Dec { negative: !self.negative && self.magnitude != 0, ..*self }
    }

    pub fn sub(&self, other: &Dec) -> Result<Dec, Overflow> {
        self.add(&other.neg())
    }

    pub fn mul(&self, other: &Dec) -> Result<Dec, Overflow> {
        Dec::settle(
            self.negative != other.negative,
            Wide::mul(self.magnitude, other.magnitude),
            self.scale as u32 + other.scale as u32,
        )
    }

    /// Division carried to as many places as the answer can hold.
    pub fn div(&self, other: &Dec) -> Option<Result<Dec, Overflow>> {
        if other.magnitude == 0 {
            return None;
        }
        let divisor = other.magnitude;
        let mut quotient = Wide::from_u128(self.magnitude / divisor);
        let mut remainder = self.magnitude % divisor;
        // value = (a / b) * 10^(sb - sa); places counts the digits past the
        // point of a / b.
        let mut places: i32 = 0;
        let shift = other.scale as i32 - self.scale as i32;
        // Carry on while there is room for another digit.
        while remainder != 0 {
            let next = quotient.mul_small(10);
            let scale_after = places + 1 - shift;
            if next.fits().is_none() || scale_after > MAX_SCALE as i32 + 2 {
                break;
            }
            remainder *= 10;
            quotient = next.add(&Wide::from_u128(remainder / divisor));
            remainder %= divisor;
            places += 1;
        }
        // One more digit and the rest for rounding.
        let mut wide = quotient;
        let mut scale = places - shift;
        let mut beyond = false;
        if remainder != 0 {
            let doubled = remainder * 2;
            let bump = doubled > divisor || (doubled == divisor && wide.0[0] & 1 == 1);
            // Settle below cuts places; round what the division left here
            // only when no cutting follows.
            if scale >= 0 && scale as u32 <= MAX_SCALE && wide.fits().is_some() {
                if bump {
                    wide = wide.add(&Wide::from_u128(1));
                }
            } else {
                beyond = true;
            }
        }
        while scale < 0 {
            wide = wide.mul_small(10);
            scale += 1;
        }
        Some(Dec::settle_beyond(self.negative != other.negative, wide, scale as u32, beyond))
    }

    pub fn cmp(&self, other: &Dec) -> Ordering {
        let difference = match self.sub(other) {
            Ok(difference) => difference,
            Err(_) => {
                return if self.negative { Ordering::Less } else { Ordering::Greater };
            }
        };
        if difference.magnitude == 0 {
            Ordering::Equal
        } else if difference.negative {
            Ordering::Less
        } else {
            Ordering::Greater
        }
    }

    /// Rounded to `places` decimals, a half to the even neighbour.
    pub fn round(&self, places: u32) -> Dec {
        if self.scale as u32 <= places {
            return *self;
        }
        let mut wide = Wide::from_u128(self.magnitude);
        let mut scale = self.scale as u32;
        let mut last = 0u64;
        let mut sticky = false;
        while scale > places {
            sticky |= last != 0;
            last = wide.div_small(10);
            scale -= 1;
        }
        if last > 5 || (last == 5 && (sticky || wide.0[0] & 1 == 1)) {
            wide = wide.add(&Wide::from_u128(1));
        }
        let magnitude = wide.fits().unwrap_or(MAX);
        Dec { negative: self.negative && magnitude != 0, magnitude, scale: scale as u8 }
    }

    /// The whole part, toward zero (`Fix`) or down (`Int`).
    pub fn whole(&self, down: bool) -> Dec {
        let mut wide = Wide::from_u128(self.magnitude);
        let mut lost = false;
        for _ in 0..self.scale {
            lost |= wide.div_small(10) != 0;
        }
        if down && self.negative && lost {
            wide = wide.add(&Wide::from_u128(1));
        }
        let magnitude = wide.fits().unwrap_or(MAX);
        Dec { negative: self.negative && magnitude != 0, magnitude, scale: 0 }
    }

    pub fn abs(&self) -> Dec {
        Dec { negative: false, ..*self }
    }
}

impl std::fmt::Display for Dec {
    /// As `CStr` writes it: no trailing zeros, no point without a fraction.
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let digits = self.magnitude.to_string();
        let scale = self.scale as usize;
        let (whole, fraction) = if digits.len() > scale {
            (digits[..digits.len() - scale].to_string(), digits[digits.len() - scale..].to_string())
        } else {
            ("0".to_string(), format!("{}{}", "0".repeat(scale - digits.len()), digits))
        };
        let fraction = fraction.trim_end_matches('0');
        if self.negative {
            f.write_str("-")?;
        }
        f.write_str(&whole)?;
        if !fraction.is_empty() {
            f.write_str(".")?;
            f.write_str(fraction)?;
        }
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::Dec;

    fn d(text: &str) -> Dec {
        Dec::parse(text).unwrap()
    }

    /// Every answer here is Excel's VBA.
    #[test]
    fn decimal_arithmetic_as_excel_answers() {
        let one = d("1");
        let three = d("3");
        assert_eq!(one.div(&three).unwrap().unwrap().to_string(), "0.3333333333333333333333333333");
        let third = one.div(&three).unwrap().unwrap();
        assert_eq!(third.mul(&three).unwrap().to_string(), "0.9999999999999999999999999999");
        assert_eq!(d("7").div(&d("2")).unwrap().unwrap().to_string(), "3.5");
        assert_eq!(d("0.1").add(&d("0.2")).unwrap(), d("0.3"));
        assert_eq!(
            d("2.5").add(&d("0.00000000000000000000000001")).unwrap().to_string(),
            "2.50000000000000000000000001"
        );
        assert_eq!(d("79228162514264337593543950335").to_string(), "79228162514264337593543950335");
        assert!(d("100000000000000000000").mul(&d("1000000000")).is_err());
        assert_eq!(d("2.5").round(0).to_string(), "2");
        assert_eq!(d("2.345").round(2).to_string(), "2.34");
        assert_eq!(d("-2.5").whole(true).to_string(), "-3");
        assert_eq!(d("-2.5").whole(false).to_string(), "-2");
        assert_eq!(Dec::from_f64(0.1).unwrap(), d("0.1"));
        assert_eq!(d("12345678901234567890.123").to_string(), "12345678901234567890.123");
        assert!(Dec::from_f64(1e30).is_err());
    }
}
