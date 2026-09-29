// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! The probability distributions behind T.DIST, CHISQ.DIST, F.DIST,
//! GAMMA.DIST, BETA.DIST and their inverses: the regularised incomplete
//! gamma and beta functions, and a root finder that turns any of the
//! cumulative functions round.

use crate::functions::{ln_gamma, regularized_beta};

/// P(a, x), the regularised lower incomplete gamma function: a series below
/// a + 1, a continued fraction for the upper part above it.
pub(crate) fn regularized_gamma_p(a: f64, x: f64) -> f64 {
    if x <= 0.0 {
        return 0.0;
    }
    // Below a + 1 the series e^-x x^a / Γ(a+1) * (1 + x/(a+1) + ...) is
    // summed straight, never as one minus the other part, which would lose
    // the last digits: measured, GAMMA.DIST(2,3,1,TRUE) 0.323323583816937.
    if x < a + 1.0 {
        let front = if a.fract() == 0.0 && a <= 170.0 {
            let mut product = (-x).exp();
            let mut k = 1.0;
            while k <= a {
                product *= x / k;
                k += 1.0;
            }
            product
        } else {
            (a * x.ln() - x - ln_gamma(a + 1.0)).exp()
        };
        let mut term = 1.0f64;
        let mut sum = 1.0f64;
        let mut n = a;
        for _ in 0..10_000 {
            n += 1.0;
            term *= x / n;
            sum += term;
            if term.abs() < sum.abs() * 1e-17 {
                break;
            }
        }
        return front * sum;
    }
    1.0 - regularized_gamma_q(a, x)
}

/// Q(a, x) = 1 - P(a, x), kept apart so a far tail keeps its digits.
pub(crate) fn regularized_gamma_q(a: f64, x: f64) -> f64 {
    if x <= 0.0 {
        return 1.0;
    }
    if x < a + 1.0 {
        return 1.0 - regularized_gamma_p(a, x);
    }
    if let Some(upper) = gamma_q_closed(a, x) {
        return upper;
    }
    regularized_gamma_q_fraction(a, x)
}

/// Q(a, x) in closed form for a whole or half-whole a up to a few hundred:
/// a Poisson sum, plus erfc(√x) for the half-whole ones.
fn gamma_q_closed(a: f64, x: f64) -> Option<f64> {
    if a > 300.0 || x > 700.0 {
        return None;
    }
    if a.fract() == 0.0 {
        let mut term = 1.0f64;
        let mut sum = 1.0f64;
        let mut k = 1.0;
        while k < a {
            term *= x / k;
            sum += term;
            k += 1.0;
        }
        return Some((-x).exp() * sum);
    }
    if (a - 0.5).fract() == 0.0 {
        let root = x.sqrt();
        let mut term = root / std::f64::consts::PI.sqrt() * 2.0;
        let mut sum = 0.0f64;
        let mut k = 0.5;
        while k < a {
            if k > 0.5 {
                term *= x / k;
            }
            sum += term;
            k += 1.0;
        }
        return Some(erfc(root) + (-x).exp() * sum);
    }
    None
}


fn regularized_gamma_q_fraction(a: f64, x: f64) -> f64 {
    let tiny = 1e-300;
    let mut b = x + 1.0 - a;
    let mut c = 1.0 / tiny;
    let mut d = 1.0 / b;
    let mut h = d;
    for i in 1..10_000 {
        let an = -(i as f64) * (i as f64 - a);
        b += 2.0;
        d = an * d + b;
        if d.abs() < tiny {
            d = tiny;
        }
        c = b + an / c;
        if c.abs() < tiny {
            c = tiny;
        }
        d = 1.0 / d;
        let step = d * c;
        h *= step;
        if (step - 1.0).abs() < 1e-17 {
            break;
        }
    }
    (-x + a * x.ln() - ln_gamma(a)).exp() * h
}

/// The x at which an increasing `cdf` reaches `p`, between `low` and
/// `high`: bisection to full precision, the bracket widened upward first
/// when it does not yet reach.
pub(crate) fn invert(p: f64, mut low: f64, mut high: f64, cdf: impl Fn(f64) -> f64) -> f64 {
    let mut grown = 0;
    while cdf(high) < p && grown < 2_000 {
        low = high;
        high *= 2.0;
        grown += 1;
    }
    for _ in 0..2_000 {
        let middle = 0.5 * (low + high);
        if middle <= low || middle >= high {
            break;
        }
        if cdf(middle) < p {
            low = middle;
        } else {
            high = middle;
        }
    }
    0.5 * (low + high)
}

/// The x at which a DECREASING `upper` tail comes down to `q`, searched
/// on the tail itself so that a small q keeps its digits.
pub(crate) fn invert_upper(q: f64, mut low: f64, mut high: f64, upper: impl Fn(f64) -> f64) -> f64 {
    let mut grown = 0;
    while upper(high) > q && grown < 2_000 {
        low = high;
        high *= 2.0;
        grown += 1;
    }
    for _ in 0..2_000 {
        let middle = 0.5 * (low + high);
        if middle <= low || middle >= high {
            break;
        }
        if upper(middle) > q {
            low = middle;
        } else {
            high = middle;
        }
    }
    0.5 * (low + high)
}

/// Student's t, the probability below x.
pub(crate) fn t_cdf(x: f64, df: f64) -> f64 {
    let tail = 0.5 * regularized_beta(df / (df + x * x), df / 2.0, 0.5);
    if x > 0.0 {
        1.0 - tail
    } else {
        tail
    }
}

pub(crate) fn t_pdf(x: f64, df: f64) -> f64 {
    (ln_gamma((df + 1.0) / 2.0) - ln_gamma(df / 2.0)).exp() / (df * std::f64::consts::PI).sqrt()
        * (1.0 + x * x / df).powf(-(df + 1.0) / 2.0)
}

/// Both tails of Student's t beyond |x|.
pub(crate) fn t_two_tailed(x: f64, df: f64) -> f64 {
    regularized_beta(df / (df + x * x), df / 2.0, 0.5)
}

/// The t below which a share p of the distribution lies.
pub(crate) fn t_inv(p: f64, df: f64) -> f64 {
    if p == 0.5 {
        return 0.0;
    }
    // Symmetric: find the upper-half point for the smaller tail, searched
    // on the tail itself.
    let tail = p.min(1.0 - p);
    let upper = invert_upper(tail, 0.0, 1.0, |x| 0.5 * t_two_tailed(x, df));
    if p < 0.5 {
        -upper
    } else {
        upper
    }
}

pub(crate) fn chisq_pdf(x: f64, k: f64) -> f64 {
    if x < 0.0 {
        return 0.0;
    }
    if x == 0.0 {
        return match k {
            k if k < 2.0 => f64::INFINITY,
            k if k == 2.0 => 0.5,
            _ => 0.0,
        };
    }
    ((k / 2.0 - 1.0) * x.ln() - x / 2.0 - (k / 2.0) * 2f64.ln() - ln_gamma(k / 2.0)).exp()
}

pub(crate) fn f_cdf(x: f64, d1: f64, d2: f64) -> f64 {
    if x <= 0.0 {
        return 0.0;
    }
    regularized_beta(d1 * x / (d1 * x + d2), d1 / 2.0, d2 / 2.0)
}

pub(crate) fn f_upper(x: f64, d1: f64, d2: f64) -> f64 {
    if x <= 0.0 {
        return 1.0;
    }
    regularized_beta(d2 / (d2 + d1 * x), d2 / 2.0, d1 / 2.0)
}

pub(crate) fn f_pdf(x: f64, d1: f64, d2: f64) -> f64 {
    if x < 0.0 {
        return 0.0;
    }
    let ln_beta = ln_gamma(d1 / 2.0) + ln_gamma(d2 / 2.0) - ln_gamma((d1 + d2) / 2.0);
    ((d1 / 2.0) * (d1 / d2).ln() + (d1 / 2.0 - 1.0) * x.ln() - ((d1 + d2) / 2.0) * (1.0 + d1 * x / d2).ln() - ln_beta).exp()
}

pub(crate) fn gamma_pdf(x: f64, alpha: f64, beta: f64) -> f64 {
    if x < 0.0 {
        return 0.0;
    }
    if x == 0.0 {
        return if alpha < 1.0 {
            f64::INFINITY
        } else if alpha == 1.0 {
            1.0 / beta
        } else {
            0.0
        };
    }
    ((alpha - 1.0) * x.ln() - x / beta - ln_gamma(alpha) - alpha * beta.ln()).exp()
}

pub(crate) fn beta_pdf(x: f64, a: f64, b: f64) -> f64 {
    let ln_beta = ln_gamma(a) + ln_gamma(b) - ln_gamma(a + b);
    ((a - 1.0) * x.ln() + (b - 1.0) * (1.0 - x).ln() - ln_beta).exp()
}

/// The Gauss error function, through the incomplete gamma function.
pub(crate) fn erf(x: f64) -> f64 {
    if x < 0.0 {
        return -erf(-x);
    }
    if x < 2.5 {
        erf_series(x)
    } else {
        1.0 - erfc_fraction(x)
    }
}

pub(crate) fn erfc(x: f64) -> f64 {
    if x < 0.0 {
        return 2.0 - erfc(-x);
    }
    if x < 0.5 {
        1.0 - erf_series(x)
    } else {
        erfc_fraction(x)
    }
}

/// erf by its series of positive terms,
/// 2/sqrt(pi) e^-x^2 (x + 2x^3/3 + 4x^5/15 + ...), which loses nothing to
/// cancellation.
fn erf_series(x: f64) -> f64 {
    let square = x * x;
    let mut term = x;
    let mut sum = x;
    let mut n = 0.0;
    for _ in 0..500 {
        n += 1.0;
        term *= 2.0 * square / (2.0 * n + 1.0);
        sum += term;
        if term < sum * 1e-17 {
            break;
        }
    }
    2.0 / std::f64::consts::PI.sqrt() * (-square).exp() * sum
}

/// erfc by its continued fraction, e^-x^2/sqrt(pi) / (x + 1/2 / (x + 1 / (x + ...))),
/// worked from the front by Lentz's method.
fn erfc_fraction(x: f64) -> f64 {
    let tiny = 1e-300;
    let mut f = x;
    let mut c = x;
    let mut d = 0.0f64;
    for k in 1..5_000 {
        let a = k as f64 / 2.0;
        d = x + a * d;
        if d.abs() < tiny {
            d = tiny;
        }
        c = x + a / c;
        if c.abs() < tiny {
            c = tiny;
        }
        d = 1.0 / d;
        let step = c * d;
        f *= step;
        if (step - 1.0).abs() < 1e-17 {
            break;
        }
    }
    (-x * x).exp() / std::f64::consts::PI.sqrt() / f
}

/// Γ(x) for any x that is not zero or a negative whole number.
pub(crate) fn gamma(x: f64) -> f64 {
    if x < 0.5 {
        return std::f64::consts::PI / ((std::f64::consts::PI * x).sin() * gamma(1.0 - x));
    }
    // A whole number is a factorial, exact as far as a double holds one.
    if x.fract() == 0.0 && x <= 171.0 {
        let mut product = 1.0;
        let mut k = 2.0;
        while k < x {
            product *= k;
            k += 1.0;
        }
        return product;
    }
    ln_gamma(x).exp()
}
