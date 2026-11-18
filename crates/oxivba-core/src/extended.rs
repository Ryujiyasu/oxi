// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Finite intermediates for typed floating arithmetic. Storage and argument
//! conversions choose their own precision; no platform floating mode is changed.

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum Fault {
    Overflow,
    DivisionByZero,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) struct Extended {
    negative: bool,
    significand: u64,
    exponent: i32,
}

impl Extended {
    const MIN_QUANTUM: i32 = -16_445;

    pub(crate) fn from_f64(value: f64) -> Option<Self> {
        if !value.is_finite() {
            return None;
        }
        let bits = value.to_bits();
        let e = ((bits >> 52) & 2047) as i32;
        let fraction = bits & ((1_u64 << 52) - 1);
        let m = if e == 0 {
            fraction
        } else {
            fraction | (1_u64 << 52)
        };
        Some(Self::normalize(
            bits >> 63 != 0,
            m,
            if e == 0 { -1074 } else { e - 1023 - 52 },
        ))
    }

    pub(crate) fn from_i64(value: i64) -> Self {
        Self::normalize(value < 0, value.unsigned_abs(), 0)
    }

    fn zero(negative: bool) -> Self {
        Self {
            negative,
            significand: 0,
            exponent: 0,
        }
    }

    fn normalize(negative: bool, magnitude: u64, exponent: i32) -> Self {
        if magnitude == 0 {
            return Self::zero(negative);
        }
        let shift = magnitude.leading_zeros();
        Self {
            negative,
            significand: magnitude << shift,
            exponent: exponent - shift as i32,
        }
    }

    fn round_shift(value: u128, shift: u32) -> u128 {
        if shift == 0 {
            return value;
        }
        if shift > 128 {
            return 0;
        }
        if shift == 128 {
            return u128::from(value > (1_u128 << 127));
        }
        let q = value >> shift;
        let remainder = value & ((1_u128 << shift) - 1);
        let half = 1_u128 << (shift - 1);
        q + u128::from(remainder > half || (remainder == half && q & 1 != 0))
    }

    fn pack(negative: bool, magnitude: u128, exponent: i32) -> Result<Self, Fault> {
        if magnitude == 0 {
            return Ok(Self::zero(negative));
        }
        let bits = 128 - magnitude.leading_zeros();
        // Quantize subnormals directly, before rounding a normal significand.
        // Otherwise two successive roundings can change a tie.
        if exponent + bits as i32 - 1 < -16_382 {
            let shift = Self::MIN_QUANTUM - exponent;
            let m = if shift >= 0 {
                Self::round_shift(magnitude, shift as u32)
            } else {
                magnitude << (-shift as u32)
            };
            return Ok(Self::normalize(negative, m as u64, Self::MIN_QUANTUM));
        }
        let result = if bits <= 64 {
            Self::normalize(negative, magnitude as u64, exponent)
        } else {
            let shift = bits - 64;
            let m = Self::round_shift(magnitude, shift);
            if m == 1_u128 << 64 {
                Self {
                    negative,
                    significand: 1_u64 << 63,
                    exponent: exponent + shift as i32 + 1,
                }
            } else {
                Self {
                    negative,
                    significand: m as u64,
                    exponent: exponent + shift as i32,
                }
            }
        };
        if result.exponent + 63 > 16_383 {
            Err(Fault::Overflow)
        } else {
            Ok(result)
        }
    }

    fn jam(value: u128, shift: u32) -> u128 {
        if shift == 0 {
            return value;
        }
        if shift >= 128 {
            return u128::from(value != 0);
        }
        (value >> shift) | u128::from(value & ((1_u128 << shift) - 1) != 0)
    }

    pub(crate) fn is_zero(self) -> bool {
        self.significand == 0
    }
    pub(crate) fn is_negative(self) -> bool {
        self.negative
    }
    pub(crate) fn negated(mut self) -> Self {
        self.negative = !self.negative;
        self
    }
    pub(crate) fn abs(mut self) -> Self {
        self.negative = false;
        self
    }

    pub(crate) fn add(self, other: Self) -> Result<Self, Fault> {
        if self.is_zero() && other.is_zero() {
            return Ok(Self::zero(self.negative && other.negative));
        }
        if self.is_zero() {
            return Ok(other);
        }
        if other.is_zero() {
            return Ok(self);
        }
        let (a, b) = if (self.exponent, self.significand) >= (other.exponent, other.significand) {
            (self, other)
        } else {
            (other, self)
        };
        let large = u128::from(a.significand) << 3;
        let small = Self::jam(
            u128::from(b.significand) << 3,
            (a.exponent - b.exponent) as u32,
        );
        let magnitude = if a.negative == b.negative {
            large + small
        } else {
            large - small
        };
        if magnitude == 0 {
            return Ok(Self::zero(false));
        }
        Self::pack(a.negative, magnitude, a.exponent - 3)
    }

    pub(crate) fn sub(self, other: Self) -> Result<Self, Fault> {
        self.add(other.negated())
    }

    pub(crate) fn mul(self, other: Self) -> Result<Self, Fault> {
        Self::pack(
            self.negative ^ other.negative,
            u128::from(self.significand) * u128::from(other.significand),
            self.exponent + other.exponent,
        )
    }

    pub(crate) fn div(self, other: Self) -> Result<Self, Fault> {
        if other.is_zero() {
            return Err(Fault::DivisionByZero);
        }
        let negative = self.negative ^ other.negative;
        if self.is_zero() {
            return Ok(Self::zero(negative));
        }
        let exponent = self.exponent - other.exponent;
        let top = exponent - i32::from(self.significand < other.significand);
        if top < -16_382 {
            let shift = exponent - Self::MIN_QUANTUM;
            if shift < -1 {
                return Ok(Self::zero(negative));
            }
            let (numerator, denominator) = if shift < 0 {
                (
                    u128::from(self.significand),
                    u128::from(other.significand) << 1,
                )
            } else {
                (
                    u128::from(self.significand) << shift,
                    u128::from(other.significand),
                )
            };
            let q = Self::round_ratio(numerator, denominator);
            return Ok(Self::normalize(negative, q as u64, Self::MIN_QUANTUM));
        }
        let shift = if self.significand >= other.significand {
            63
        } else {
            64
        };
        let q = Self::round_ratio(
            u128::from(self.significand) << shift,
            u128::from(other.significand),
        );
        Self::pack(negative, q, exponent - shift)
    }

    fn round_ratio(numerator: u128, denominator: u128) -> u128 {
        let q = numerator / denominator;
        let remainder = numerator % denominator;
        q + u128::from(2 * remainder > denominator || (2 * remainder == denominator && q & 1 != 0))
    }

    pub(crate) fn to_f64(self) -> f64 {
        let sign = u64::from(self.negative) << 63;
        if self.is_zero() {
            return f64::from_bits(sign);
        }
        let mut e = self.exponent + 63;
        if e < -1022 {
            let m = Self::round_shift(u128::from(self.significand), (-1074 - self.exponent) as u32);
            return f64::from_bits(sign | m as u64);
        }
        let mut m = Self::round_shift(u128::from(self.significand), 11);
        if m == 1_u128 << 53 {
            m >>= 1;
            e += 1;
        }
        if e > 1023 {
            return f64::from_bits(sign | 0x7ff0_0000_0000_0000);
        }
        f64::from_bits(sign | (((e + 1023) as u64) << 52) | ((m as u64) & ((1_u64 << 52) - 1)))
    }

    pub(crate) fn to_f32(self) -> f32 {
        let sign = u32::from(self.negative) << 31;
        if self.is_zero() {
            return f32::from_bits(sign);
        }
        let mut e = self.exponent + 63;
        if e < -126 {
            let m = Self::round_shift(u128::from(self.significand), (-149 - self.exponent) as u32);
            return f32::from_bits(sign | m as u32);
        }
        let mut m = Self::round_shift(u128::from(self.significand), 40);
        if m == 1_u128 << 24 {
            m >>= 1;
            e += 1;
        }
        if e > 127 {
            return f32::from_bits(sign | 0x7f80_0000);
        }
        f32::from_bits(sign | (((e + 127) as u32) << 23) | ((m as u32) & ((1_u32 << 23) - 1)))
    }

    pub(crate) fn round_i64(self) -> Option<i64> {
        if self.is_zero() {
            return Some(0);
        }
        if self.exponent > 0 {
            return None;
        }
        let m = Self::round_shift(u128::from(self.significand), (-self.exponent) as u32);
        if self.negative {
            if m > 1_u128 << 63 {
                None
            } else if m == 1_u128 << 63 {
                Some(i64::MIN)
            } else {
                Some(-(m as i64))
            }
        } else if m > i64::MAX as u128 {
            None
        } else {
            Some(m as i64)
        }
    }

    pub(crate) fn floor(self) -> Self {
        if self.is_zero() || self.exponent >= 0 {
            return self;
        }
        let shift = (-self.exponent) as u32;
        if shift >= 64 {
            return if self.negative {
                Self::from_i64(-1)
            } else {
                Self::zero(false)
            };
        }
        let whole = self.significand >> shift;
        let fraction = self.significand & ((1_u64 << shift) - 1) != 0;
        Self::normalize(
            self.negative,
            whole + u64::from(self.negative && fraction),
            0,
        )
    }
    pub(crate) fn truncated(self) -> Self {
        let whole = self.abs().floor();
        if self.negative {
            whole.negated()
        } else {
            whole
        }
    }
}

#[cfg(test)]
#[path = "extended_tests.rs"]
mod tests;
