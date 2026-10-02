// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! BESSELJ, BESSELY, BESSELI and BESSELK. Excel works them with the
//! classic polynomial and rational approximations (good to about eight
//! digits, not fifteen) and carries them to higher orders by recurrence --
//! downward (Miller) for J and I, upward for Y and K. Measured over 112
//! cases: these give Excel's answers to the last digit for every Y, I and K
//! and for J below x = 8.
//!
//! The approximations are reproduced with their published eight-digit
//! constants (2/pi as 0.636619772, Euler's gamma as 0.5772156649) because
//! that is what Excel evaluates; the exact library constants would move the
//! last digits away from Excel's.
#![allow(clippy::approx_constant)]

const ACC: f64 = 40.0;
const BIG: f64 = 1e10;
const SMALL: f64 = 1e-10;

fn j0(x: f64) -> f64 {
    let ax = x.abs();
    if ax < 8.0 {
        let y = x * x;
        let a1 = 57568490574.0 + y * (-13362590354.0 + y * (651619640.7 + y * (-11214424.18 + y * (77392.33017 + y * (-184.9052456)))));
        let a2 = 57568490411.0 + y * (1029532985.0 + y * (9494680.718 + y * (59272.64853 + y * (267.8532712 + y))));
        return a1 / a2;
    }
    let z = 8.0 / ax;
    let y = z * z;
    let xx = ax - 0.785398164;
    let a1 = 1.0 + y * (-0.1098628627e-2 + y * (0.2734510407e-4 + y * (-0.2073370639e-5 + y * 0.2093887211e-6)));
    let a2 = -0.1562499995e-1 + y * (0.1430488765e-3 + y * (-0.6911147651e-5 + y * (0.7621095161e-6 - y * 0.934935152e-7)));
    (0.636619772 / ax).sqrt() * (xx.cos() * a1 - z * xx.sin() * a2)
}

fn j1(x: f64) -> f64 {
    let ax = x.abs();
    if ax < 8.0 {
        let y = x * x;
        let a1 = x * (72362614232.0 + y * (-7895059235.0 + y * (242396853.1 + y * (-2972611.439 + y * (15704.48260 + y * (-30.16036606))))));
        let a2 = 144725228442.0 + y * (2300535178.0 + y * (18583304.74 + y * (99447.43394 + y * (376.9991397 + y))));
        return a1 / a2;
    }
    let z = 8.0 / ax;
    let y = z * z;
    let xx = ax - 2.356194491;
    let a1 = 1.0 + y * (0.183105e-2 + y * (-0.3516396496e-4 + y * (0.2457520174e-5 + y * (-0.240337019e-6))));
    let a2 = 0.04687499995 + y * (-0.2002690873e-3 + y * (0.8449199096e-5 + y * (-0.88228987e-6 + y * 0.105787412e-6)));
    let answer = (0.636619772 / ax).sqrt() * (xx.cos() * a1 - z * xx.sin() * a2);
    if x < 0.0 {
        -answer
    } else {
        answer
    }
}

fn y0(x: f64) -> f64 {
    if x < 8.0 {
        let y = x * x;
        let a1 = -2957821389.0 + y * (7062834065.0 + y * (-512359803.6 + y * (10879881.29 + y * (-86327.92757 + y * 228.4622733))));
        let a2 = 40076544269.0 + y * (745249964.8 + y * (7189466.438 + y * (47447.26470 + y * (226.1030244 + y))));
        return a1 / a2 + 0.636619772 * j0(x) * x.ln();
    }
    let z = 8.0 / x;
    let y = z * z;
    let xx = x - 0.785398164;
    let a1 = 1.0 + y * (-0.1098628627e-2 + y * (0.2734510407e-4 + y * (-0.2073370639e-5 + y * 0.2093887211e-6)));
    let a2 = -0.1562499995e-1 + y * (0.1430488765e-3 + y * (-0.6911147651e-5 + y * (0.7621095161e-6 + y * (-0.934945152e-7))));
    (0.636619772 / x).sqrt() * (xx.sin() * a1 + z * xx.cos() * a2)
}

fn y1(x: f64) -> f64 {
    if x < 8.0 {
        let y = x * x;
        let a1 = x
            * (-0.4900604943e13
                + y * (0.1275274390e13 + y * (-0.5153438139e11 + y * (0.7349264551e9 + y * (-0.4237922726e7 + y * 0.8511937935e4)))));
        let a2 = 0.2499580570e14
            + y * (0.4244419664e12 + y * (0.3733650367e10 + y * (0.2245904002e8 + y * (0.1020426050e6 + y * (0.3549632885e3 + y)))));
        return a1 / a2 + 0.636619772 * (j1(x) * x.ln() - 1.0 / x);
    }
    let z = 8.0 / x;
    let y = z * z;
    let xx = x - 2.356194491;
    let a1 = 1.0 + y * (0.183105e-2 + y * (-0.3516396496e-4 + y * (0.2457520174e-5 + y * (-0.240337019e-6))));
    let a2 = 0.04687499995 + y * (-0.2002690873e-3 + y * (0.8449199096e-5 + y * (-0.88228987e-6 + y * 0.105787412e-6)));
    (0.636619772 / x).sqrt() * (xx.sin() * a1 + z * xx.cos() * a2)
}

fn i0(x: f64) -> f64 {
    let ax = x.abs();
    if ax < 3.75 {
        let y = (x / 3.75).powi(2);
        return 1.0 + y * (3.5156229 + y * (3.0899424 + y * (1.2067492 + y * (0.2659732 + y * (0.360768e-1 + y * 0.45813e-2)))));
    }
    let y = 3.75 / ax;
    (ax.exp() / ax.sqrt())
        * (0.39894228
            + y * (0.1328592e-1
                + y * (0.225319e-2
                    + y * (-0.157565e-2 + y * (0.916281e-2 + y * (-0.2057706e-1 + y * (0.2635537e-1 + y * (-0.1647633e-1 + y * 0.392377e-2))))))))
}

fn i1(x: f64) -> f64 {
    let ax = x.abs();
    let answer = if ax < 3.75 {
        let y = (x / 3.75).powi(2);
        ax * (0.5 + y * (0.87890594 + y * (0.51498869 + y * (0.15084934 + y * (0.2658733e-1 + y * (0.301532e-2 + y * 0.32411e-3))))))
    } else {
        let y = 3.75 / ax;
        let inner = 0.2282967e-1 + y * (-0.2895312e-1 + y * (0.1787654e-1 - y * 0.420059e-2));
        let outer = 0.39894228 + y * (-0.3988024e-1 + y * (-0.362018e-2 + y * (0.163801e-2 + y * (-0.1031555e-1 + y * inner))));
        outer * (ax.exp() / ax.sqrt())
    };
    if x < 0.0 {
        -answer
    } else {
        answer
    }
}

fn k0(x: f64) -> f64 {
    if x <= 2.0 {
        let y = x * x / 4.0;
        return -(x / 2.0).ln() * i0(x)
            + (-0.57721566 + y * (0.42278420 + y * (0.23069756 + y * (0.3488590e-1 + y * (0.262698e-2 + y * (0.10750e-3 + y * 0.74e-5))))));
    }
    let y = 2.0 / x;
    ((-x).exp() / x.sqrt())
        * (1.25331414 + y * (-0.7832358e-1 + y * (0.2189568e-1 + y * (-0.1062446e-1 + y * (0.587872e-2 + y * (-0.251540e-2 + y * 0.53208e-3))))))
}

fn k1(x: f64) -> f64 {
    if x <= 2.0 {
        let y = x * x / 4.0;
        return (x / 2.0).ln() * i1(x)
            + (1.0 / x)
                * (1.0 + y * (0.15443144 + y * (-0.67278579 + y * (-0.18156897 + y * (-0.1919402e-1 + y * (-0.110404e-2 + y * (-0.4686e-4)))))));
    }
    let y = 2.0 / x;
    ((-x).exp() / x.sqrt())
        * (1.25331414 + y * (0.23498619 + y * (-0.3655620e-1 + y * (0.1504268e-1 + y * (-0.780353e-2 + y * (0.325614e-2 + y * (-0.68245e-3)))))))
}

pub(crate) fn bessel_j(n: i64, x: f64) -> f64 {
    match n {
        0 => return j0(x),
        1 => return j1(x),
        _ => {}
    }
    let ax = x.abs();
    if ax == 0.0 {
        return 0.0;
    }
    let answer = if ax > n as f64 {
        let tox = 2.0 / ax;
        let (mut below, mut here) = (j0(ax), j1(ax));
        for j in 1..n {
            let above = j as f64 * tox * here - below;
            below = here;
            here = above;
        }
        here
    } else {
        let tox = 2.0 / ax;
        let m = 2 * ((n + (ACC * n as f64).sqrt() as i64) / 2);
        let mut adding = false;
        let (mut above, mut answer, mut sum, mut here) = (0.0f64, 0.0f64, 0.0f64, 1.0f64);
        for j in (1..=m).rev() {
            let below = j as f64 * tox * here - above;
            above = here;
            here = below;
            if here.abs() > BIG {
                here *= SMALL;
                above *= SMALL;
                answer *= SMALL;
                sum *= SMALL;
            }
            if adding {
                sum += here;
            }
            adding = !adding;
            if j == n {
                answer = above;
            }
        }
        answer / (2.0 * sum - here)
    };
    if x < 0.0 && n % 2 == 1 {
        -answer
    } else {
        answer
    }
}

pub(crate) fn bessel_y(n: i64, x: f64) -> f64 {
    match n {
        0 => return y0(x),
        1 => return y1(x),
        _ => {}
    }
    // Worked as 2j·Y/x: measured, BESSELY(1.5,5) -37.1903084392881 comes
    // out that way (j·(2/x)·Y gives …880), where K's digits want j·(2/x)·K.
    let (mut below, mut here) = (y0(x), y1(x));
    for j in 1..n {
        let above = 2.0 * j as f64 * here / x - below;
        below = here;
        here = above;
    }
    here
}

pub(crate) fn bessel_k(n: i64, x: f64) -> f64 {
    match n {
        0 => return k0(x),
        1 => return k1(x),
        _ => {}
    }
    let tox = 2.0 / x;
    let (mut below, mut here) = (k0(x), k1(x));
    for j in 1..n {
        let above = below + j as f64 * tox * here;
        below = here;
        here = above;
    }
    here
}

pub(crate) fn bessel_i(n: i64, x: f64) -> f64 {
    match n {
        0 => return i0(x),
        1 => return i1(x),
        _ => {}
    }
    if x == 0.0 {
        return 0.0;
    }
    let tox = 2.0 / x.abs();
    let (mut above, mut answer, mut here) = (0.0f64, 0.0f64, 1.0f64);
    let start = 2 * (n + (ACC * n as f64).sqrt() as i64);
    for j in (1..=start).rev() {
        let below = above + j as f64 * tox * here;
        above = here;
        here = below;
        if here.abs() > BIG {
            answer *= SMALL;
            here *= SMALL;
            above *= SMALL;
        }
        if j == n {
            answer = above;
        }
    }
    answer *= i0(x) / here;
    if x < 0.0 && n % 2 == 1 {
        -answer
    } else {
        answer
    }
}
