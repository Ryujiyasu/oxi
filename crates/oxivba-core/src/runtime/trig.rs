// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.
//
// Portions adapted from AMD libm's SSE2 sin/cos and remainder_piby2,
// distributed in ReactOS sdk/lib/crt/math/libm_sse2 under this notice:
//
// MIT License
// Copyright (c) 2002-2019 Advanced Micro Devices, Inc.
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this Software and associated documentaon files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE.

const LIMIT: f64 = 9.223372e18;
const QUARTER_PI: f64 = f64::from_bits(0x3fe9_21fb_5444_2d18);
const C: [f64; 6] = [
    f64::from_bits(0x3fa5_5555_5555_5555),
    f64::from_bits(0xbf56_c16c_16c1_6967),
    f64::from_bits(0x3efa_01a0_19f4_ec91),
    f64::from_bits(0xbe92_7e4f_a17f_667b),
    f64::from_bits(0x3e21_eeb6_9038_2eec),
    f64::from_bits(0xbda9_07db_4725_8aa7),
];
const S: [f64; 6] = [
    f64::from_bits(0xbfc5_5555_5555_5555),
    f64::from_bits(0x3f81_1111_1111_0bb3),
    f64::from_bits(0xbf2a_01a0_19e8_3e5c),
    f64::from_bits(0x3ec7_1de3_796c_de01),
    f64::from_bits(0xbe5a_e600_b42f_dfa7),
    f64::from_bits(0x3de5_e0b2_f9a4_3bb8),
];

// Keep the arithmetic order: each multiplication and addition rounds to Double.
pub(super) fn sin(value: f64) -> f64 {
    if value == 0.0 || value.is_nan() {
        return value;
    }
    let x = value.abs();
    if x >= LIMIT {
        // The caller rejects these arguments before evaluating this kernel.
        return value.sin();
    }
    if x <= QUARTER_PI {
        let z = value * value;
        let z6 = (z * z) * z;
        let a = (S[5] * z + S[4]) * z + S[3];
        let b = (S[2] * z + S[1]) * z + S[0];
        return value + (value * z) * (a * z6 + b);
    }
    let (region, r, rr) = reduce(x);
    let answer = if region & 1 != 0 {
        cosine_kernel(r, rr)
    } else {
        sine_kernel(r, rr)
    };
    if (region & 2 != 0) != value.is_sign_negative() {
        -answer
    } else {
        answer
    }
}

pub(super) fn cos(value: f64) -> f64 {
    if value.is_nan() {
        return value;
    }
    let x = value.abs();
    if x >= LIMIT {
        return value.cos();
    }
    if x < QUARTER_PI {
        if x < 1.0 / 8192.0 {
            return if x < 1.0 / 134217728.0 {
                1.0
            } else {
                1.0 - (x * x) * 0.5
            };
        }
        let z = x * x;
        let z4 = z * z;
        let z8 = z4 * z4;
        let a = (C[1] * z + C[0]) * z4;
        let b = (C[3] * z + C[2]) * z8;
        let d = (z4 * z8) * (C[5] * z + C[4]);
        let negative_half = z * -0.5;
        let t = negative_half + 1.0;
        return (((1.0 - t) + negative_half) + ((a + b) + d)) + t;
    }
    let (region, r, rr) = reduce(x);
    let answer = if region & 1 != 0 {
        sine_kernel(r, rr)
    } else {
        cosine_kernel(r, rr)
    };
    if (region + 1) & 2 != 0 {
        -answer
    } else {
        answer
    }
}

pub(super) fn tan(value: f64) -> f64 {
    if value == 0.0 || value.is_nan() {
        return value;
    }
    sin(value) / cos(value)
}

fn sine_kernel(r: f64, rr: f64) -> f64 {
    let z = r * r;
    let z6 = (z * z) * z;
    let a = (S[5] * z + S[4]) * z + S[3];
    let b = (S[2] * z + S[1]) * z + S[0];
    let polynomial = a * z6 + b;
    (rr + ((r * z) * polynomial - (z * 0.5) * rr)) + r
}

fn cosine_kernel(r: f64, rr: f64) -> f64 {
    let z = r * r;
    let z6 = (z * z) * z;
    let a = (C[5] * z + C[4]) * z + C[3];
    let b = (C[2] * z + C[1]) * z + C[0];
    let negative_t = 0.5 * z - 1.0;
    let q = ((negative_t + 1.0) - 0.5 * z) - r * rr;
    ((b + a * z6) * (z * z) + q) - negative_t
}

fn exponent(value: f64) -> i32 {
    ((value.to_bits() >> 52) & 2047) as i32
}

fn reduce(x: f64) -> (u32, f64, f64) {
    if x >= 500000.0 {
        return reduce_large(x);
    }
    let n = (x * f64::from_bits(0x3fe4_5f30_6dc9_c883) + 0.5) as u32;
    let multiple = f64::from(n);
    let mut head = x - multiple * f64::from_bits(0x3ff9_21fb_5440_0000);
    let mut tail = multiple * f64::from_bits(0x3dd0_b461_1a62_6331);
    let mut r = head - tail;
    if exponent(x) - exponent(r) > 15 {
        let t = head;
        let product = multiple * f64::from_bits(0x3dd0_b461_1a60_0000);
        head = t - product;
        tail = multiple * f64::from_bits(0x3ba3_198a_2e03_7073) - ((t - head) - product);
        r = head - tail;
    }
    (n, r, (head - r) - tail)
}

// The supported VBA argument range has exponent at most 62. These 10-bit
// chunks contain the complete 180-bit reduction window needed in that range.
const PI_BITS: [u64; 26] = [
    0, 0, 0, 0, 0, 0, 162, 998, 54, 915, 580, 84, 671, 777, 855, 839, 851, 311, 448, 877, 553, 358,
    316, 270, 260, 127,
];

fn reduce_large(value: f64) -> (u32, f64, f64) {
    let bits = value.to_bits();
    let xexp = ((bits >> 52) & 2047) as usize - 1023;
    let significand = (bits & ((1_u64 << 52) - 1)) | (1_u64 << 52);
    let first = xexp / 10;
    let remainder_exp = xexp - first * 10;
    let last = first + 18;
    let mut result = [0_u64; 20];
    let mut carry = 0;
    for i in (first..=last).rev() {
        let u = PI_BITS[i] * significand + carry;
        result[i - first] = u & 1023;
        carry = u >> 10;
    }
    let leading = (((result[0] << 10) | result[1]) >> (9 - remainder_exp)) & 7;
    let complement = leading & 1 != 0;
    let region = if complement {
        ((leading >> 1) + 1) & 3
    } else {
        leading >> 1
    } as u32;
    let mut i = 1;
    let mut mantissa =
        (if complement { !result[1] } else { result[1] }) & ((1_u64 << (10 - remainder_exp)) - 1);
    while mantissa < 1_u64 << 53 {
        i += 1;
        let chunk = if complement {
            !result[i] & 1023
        } else {
            result[i]
        };
        mantissa = (mantissa << 10) | chunk;
    }
    let mut high = (if complement {
        !result[i + 1]
    } else {
        result[i + 1]
    }) << 54;
    let mut rexp = 52 + remainder_exp as i32 - 10 * i as i32;
    while mantissa >= 1_u64 << 53 {
        rexp += 1;
        high = (high >> 1) | ((mantissa & 1) << 63);
        mantissa >>= 1;
    }
    let mut fraction_bits = (mantissa & ((1_u64 << 52) - 1)) | (((rexp + 1023) as u64) << 52);
    if complement {
        fraction_bits |= 1_u64 << 63;
    }
    let fraction = f64::from_bits(fraction_bits);
    high >>= 12;
    let scale = ((rexp + 1023 - 53) as u64) << 52;
    let mut tail = f64::from_bits(scale | high) - f64::from_bits(scale);
    tail *= if complement { -2.0 } else { 2.0 };
    let head = f64::from_bits(fraction.to_bits() & 0xffff_ffff_f800_0000);
    let rest = fraction - head;
    let lead = f64::from_bits(0x3ff9_21fb_5444_2d18);
    let p1 = f64::from_bits(0x3ff9_21fb_5000_0000);
    let p2 = f64::from_bits(0x3e51_10b4_6000_0000);
    let p3 = f64::from_bits(0x3c91_a626_3314_5c06);
    let c = lead * fraction;
    let cc =
        ((((p1 * head - c) + p1 * rest) + p2 * head) + p2 * rest) + (lead * tail + p3 * fraction);
    let r = c + cc;
    (region, r, (c - r) + cc)
}
