// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Where a chart's value axis begins, ends and is ticked.
//!
//! A chart part often pins none of the three, and a graph drawn from zero
//! where Excel began at 280 is wrong in every pixel, so this works them out
//! the way Excel does. Measured through COM by `tools/metrics/_xlsx_chart_ends.py`
//! and `_xlsx_chart_unit.py`: 18 series for the ends, 270 axes for the spacing.

/// The ends of a value axis and the distance between its ticks.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Scale {
    pub low: f64,
    pub high: f64,
    pub unit: f64,
}

impl Scale {
    /// Where a value falls between the ends, 0 at the foot and 1 at the top.
    pub fn at(&self, value: f64) -> f64 {
        if self.high <= self.low {
            0.0
        } else {
            (value - self.low) / (self.high - self.low)
        }
    }
}

/// The nice numbers Excel steps a tick by: one, two or five, times a power of
/// ten. Two and a half never appears.
fn rounder(wanted: f64) -> f64 {
    if !(wanted > 0.0) {
        return 1.0;
    }
    let power = wanted.log10().floor();
    for step in [1.0, 2.0, 5.0, 10.0] {
        let held = step * 10_f64.powf(power);
        // A hair of slack: a span divided by its own tick count lands a
        // whisker over the step it came from.
        if held >= wanted * (1.0 - 1e-9) {
            return held;
        }
    }
    10_f64.powf(power + 1.0)
}

/// The most intervals a plot that tall will take.
///
/// Excel leaves every label about 13.75pt of the axis and never draws more
/// than ten intervals however tall the plot is (measured to 578pt). A label
/// larger than the default 10pt takes proportionally more room; one smaller
/// takes no less than a 10pt label would.
fn intervals(plot_points: f64, label_points: f32) -> f64 {
    let room = if label_points <= 10.0 {
        13.75
    } else {
        1.65 * label_points as f64
    };
    (plot_points / room).floor().clamp(1.0, 10.0)
}

/// The scale Excel gives a value axis, from what the chart pins and what it
/// leaves out.
///
/// `stated` is the axis's own `<c:min>`, `<c:max>` and `<c:majorUnit>`. What
/// is missing is worked out from the numbers plotted:
///
/// * the axis starts at zero when nothing plotted is negative and the
///   smallest number is no more than five sixths of the largest — otherwise a
///   twentieth of the spread below the smallest, or half of it below when the
///   whole series sits well above zero;
/// * it ends a twentieth of the spread above the largest, or at zero when
///   nothing plotted is positive;
/// * both ends are then pushed out to a multiple of the tick.
pub fn scale(
    numbers: &[f64],
    stated: (Option<f64>, Option<f64>, Option<f64>),
    plot_points: f64,
    label_points: f32,
) -> Scale {
    let (min, max, unit) = stated;
    let low_seen = numbers.iter().cloned().fold(f64::INFINITY, f64::min);
    let high_seen = numbers.iter().cloned().fold(f64::NEG_INFINITY, f64::max);
    if !low_seen.is_finite() || !high_seen.is_finite() {
        return Scale {
            low: min.unwrap_or(0.0),
            high: max.unwrap_or(1.0),
            unit: unit.unwrap_or(0.1),
        };
    }

    // Nothing but zeroes gets an axis of its own.
    if low_seen == 0.0 && high_seen == 0.0 {
        return Scale {
            low: min.unwrap_or(0.0),
            high: max.unwrap_or(1.0),
            unit: unit.unwrap_or(0.1),
        };
    }

    // The foot, before it is pushed out to a tick.
    let foot = || {
        if high_seen <= 0.0 {
            // Everything is negative: the axis runs up to zero.
            low_seen - 0.05 * (high_seen - low_seen).abs().max(high_seen.abs())
        } else if low_seen < 0.0 {
            low_seen - 0.05 * (high_seen - low_seen)
        } else if low_seen <= 5.0 / 6.0 * high_seen || low_seen == high_seen {
            0.0
        } else {
            // A series that sits well clear of zero gets half its own spread
            // of room below it, not a twentieth.
            low_seen - 0.5 * (high_seen - low_seen)
        }
    };
    let head = |low: f64| {
        if high_seen <= 0.0 {
            0.0
        } else {
            high_seen + 0.05 * (high_seen - low)
        }
    };

    let rough_low = min.unwrap_or_else(foot);
    let rough_high = max.unwrap_or_else(|| head(rough_low));
    let unit = unit.unwrap_or_else(|| {
        let span = (rough_high - rough_low).abs();
        rounder(span / intervals(plot_points, label_points))
    });

    // The ends are only pushed out where the chart left them open.
    let low = match min {
        Some(pinned) => pinned,
        None => (rough_low / unit).floor() * unit,
    };
    let high = match max {
        Some(pinned) => pinned,
        None => (rough_high / unit).ceil() * unit,
    };
    Scale { low, high, unit: unit.max(f64::MIN_POSITIVE) }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// Every series COM was asked about, with the axis Excel drew for it.
    /// The plot was 263.93pt tall, or 277.79 where the categories moved off
    /// the foot of a chart that crosses zero.
    const ENDS: &[(&[f64], f64, f64, f64, f64)] = &[
        (&[1.0, 2.0, 3.0], 0.0, 3.5, 0.5, 263.93),
        (&[10.0, 20.0, 30.0], 0.0, 35.0, 5.0, 263.93),
        (&[90.0, 130.0, 340.0], 0.0, 400.0, 50.0, 263.93),
        (&[300.0, 320.0, 340.0], 280.0, 350.0, 10.0, 263.93),
        (&[330.0, 335.0, 340.0], 324.0, 342.0, 2.0, 263.93),
        (&[0.0, 50.0, 100.0], 0.0, 120.0, 20.0, 263.93),
        (&[5.0, 5.0, 5.0], 0.0, 6.0, 1.0, 263.93),
        (&[-10.0, 0.0, 10.0], -15.0, 15.0, 5.0, 277.79),
        (&[-100.0, -50.0, -20.0], -120.0, 0.0, 20.0, 277.79),
        (&[0.1, 0.2, 0.35], 0.0, 0.4, 0.05, 263.93),
        (&[1000.0, 2000.0, 12000.0], 0.0, 14000.0, 2000.0, 263.93),
        (&[95.0, 96.0, 97.0], 94.0, 97.5, 0.5, 263.93),
        (&[50.0, 60.0, 70.0], 0.0, 80.0, 10.0, 263.93),
        (&[1.0, 100.0, 10000.0], 0.0, 12000.0, 2000.0, 263.93),
        (&[12.0, 12.0, 13.0], 11.4, 13.2, 0.2, 263.93),
        (&[-5.0, 20.0, 60.0], -10.0, 70.0, 10.0, 277.79),
        (&[200.0, 240.0, 260.0], 0.0, 300.0, 50.0, 263.93),
        (&[0.0, 0.0, 0.0], 0.0, 1.0, 0.1, 263.93),
    ];

    #[test]
    fn an_axis_left_open_ends_where_excel_ends_it() {
        for (numbers, low, high, unit, plot) in ENDS {
            let found = scale(numbers, (None, None, None), *plot, 10.0);
            assert!(
                (found.low - low).abs() < 1e-6
                    && (found.high - high).abs() < 1e-6
                    && (found.unit - unit).abs() < 1e-9,
                "{numbers:?}: wanted {low}..{high} by {unit}, got {found:?}"
            );
        }
    }

    /// The spacing COM reported for a plot of a given height, over the ranges
    /// where the choice is not obvious. `_xlsx_chart_unit.py` measured 270 of
    /// these; the ones kept here are the boundaries the rule turns on.
    const UNITS: &[(f64, f64, f64, f64)] = &[
        // A short plot: three intervals at most.
        (0.0, 350.0, 53.93, 200.0),
        (0.0, 8.0, 53.93, 5.0),
        (0.0, 3.0, 53.93, 1.0),
        (0.0, 60.0, 53.93, 20.0),
        (0.0, 50.0, 53.93, 20.0),
        (0.0, 0.35, 53.93, 0.2),
        (100.0, 350.0, 53.93, 100.0),
        // Ten is as many as Excel will draw, however tall the plot.
        (0.0, 100.0, 143.93, 10.0),
        (0.0, 1.0, 143.93, 0.1),
        (0.0, 12.0, 143.93, 2.0),
        (0.0, 100.0, 563.93, 10.0),
        (0.0, 350.0, 263.93, 50.0),
        (0.0, 175.0, 113.93, 50.0),
        (0.0, 8.0, 113.93, 1.0),
        (0.0, 35.0, 113.93, 5.0),
    ];

    #[test]
    fn the_ticks_stand_where_excel_stands_them() {
        for (low, high, plot, unit) in UNITS {
            let found = scale(&[*low, *high], (Some(*low), Some(*high), None), *plot, 10.0);
            assert!(
                (found.unit - unit).abs() < 1e-9,
                "{low}..{high} on {plot}pt: wanted {unit}, got {}",
                found.unit
            );
        }
    }

    #[test]
    fn a_larger_label_leaves_room_for_fewer_ticks() {
        // Measured: 20pt labels on a 137.83pt plot take four intervals where
        // 10pt labels take ten.
        let big = scale(&[0.0, 100.0], (Some(0.0), Some(100.0), None), 137.83, 20.0);
        assert_eq!(big.unit, 50.0);
        let small = scale(&[0.0, 100.0], (Some(0.0), Some(100.0), None), 137.83, 10.0);
        assert_eq!(small.unit, 10.0);
    }

    #[test]
    fn what_the_chart_pins_is_left_where_it_is() {
        // The corpus's charts state the top of the axis and nothing else.
        let found = scale(&[90.0, 130.0, 340.7], (None, Some(350.0), None), 277.0, 10.0);
        assert_eq!(found.low, 0.0);
        assert_eq!(found.high, 350.0);
        assert_eq!(found.unit, 50.0);
        assert!((found.at(175.0) - 0.5).abs() < 1e-9);
    }
}
