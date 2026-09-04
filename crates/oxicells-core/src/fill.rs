// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! How a line of cells carries on when it is pulled out: Excel's fill handle
//! and `Range.AutoFill`, as one law.
//!
//! The law was measured against Excel in 2026-08 (the editor's fill handle)
//! and 2026-09 (`AutoFill` from VBA), and the whole of it is about how the
//! selection is DIVIDED. The line is cut into runs of neighbours of the same
//! fine kind -- a number, a member of one of Excel's built-in lists, a text
//! ending in digits with a given prefix, plain text, a formula -- and each
//! run carries on by itself, so `1, a` pulled down gives `1, a, 2, a, 3`.
//!
//! Within a run: numbers of two or more follow a least-squares line (1, 2, 4
//! continues 5.33, not 6); a single number in a block counts up by one, and
//! a single number ALONE is copied; a date alone moves a day and a time alone
//! an hour, told apart by the format's letters; dates a whole number of
//! months apart, all on the same day of the month or all at their months'
//! ends, run by calendar months; list members walk the list at their stride
//! and wrap; numbered text counts its digits up, keeping leading zeros;
//! plain text repeats; a formula's relative references move with it.

use crate::translate_formula_references;

/// What the fill needs to know of one cell of the line.
#[derive(Debug, Clone, Default)]
pub struct Seed {
    pub number: Option<f64>,
    pub text: String,
    pub formula: Option<String>,
    pub number_format: Option<String>,
}

/// `Range.AutoFill`'s `Type`, and what the fill handle does without one.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum How {
    /// xlFillDefault: the law above.
    Default,
    /// xlFillCopy: the line repeated as it is.
    Copy,
    /// xlFillSeries: the law, except that a lone number counts up too.
    Series,
    /// xlFillDays / xlFillWeekdays / xlFillMonths / xlFillYears: dates move
    /// by that unit; anything that is not a date is copied.
    Days,
    Weekdays,
    Months,
    Years,
    /// xlLinearTrend: numbers follow their fitted line, whatever they are.
    LinearTrend,
    /// xlGrowthTrend: numbers follow their fitted geometric ratio.
    GrowthTrend,
}

/// Which way the line runs, which is the axis a formula's references move
/// along.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Along {
    Rows,
    Columns,
}

/// What one cell beyond the line receives.
#[derive(Debug, Clone, PartialEq)]
pub enum Filled {
    Empty,
    Number(f64),
    Text(String),
    Formula(String),
}

/// One step of the fill: what the cell gets, and which cell of the line it
/// is patterned on -- whose formats it wears. Measured: `1` bold over `2`
/// plain pulled down gives 3 bold, 4 plain, 5 bold.
#[derive(Debug, Clone, PartialEq)]
pub struct Step {
    pub filled: Filled,
    pub seat: usize,
}

/// The lists Excel continues rather than repeats, read out of Excel itself
/// with `Application.GetCustomListContents`. This install carries eleven; a
/// Japanese Excel knows the weekdays and months twice over, the old month
/// names, the zodiac and the ten stems.
const LISTS: &[&[&str]] = &[
    &["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"],
    &["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"],
    &["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
    &[
        "January", "February", "March", "April", "May", "June", "July", "August", "September",
        "October", "November", "December",
    ],
    &["日", "月", "火", "水", "木", "金", "土"],
    &["日曜日", "月曜日", "火曜日", "水曜日", "木曜日", "金曜日", "土曜日"],
    &["1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月"],
    &["第1四半期", "第2四半期", "第3四半期", "第4四半期"],
    &[
        "睦月", "如月", "弥生", "卯月", "皐月", "水無月", "文月", "葉月", "長月", "神無月", "霜月",
        "師走",
    ],
    &["子", "丑", "寅", "卯", "辰", "巳", "午", "未", "申", "酉", "戌", "亥"],
    &["甲", "乙", "丙", "丁", "戊", "己", "庚", "辛", "壬", "癸"],
];

/// The kind of a cell, finely enough to say whether it belongs in the same
/// run as its neighbour.
#[derive(Debug, Clone, PartialEq)]
enum Kind {
    Number,
    List(usize),
    Counted(String),
    Text,
    Formula,
}

fn kind_of(seed: &Seed) -> Kind {
    if seed.formula.is_some() {
        return Kind::Formula;
    }
    if seed.number.is_some() {
        return Kind::Number;
    }
    if let Some((list, _)) = in_list(&seed.text) {
        return Kind::List(list);
    }
    match tail_digits(&seed.text) {
        Some((before, _)) => Kind::Counted(before.to_string()),
        None => Kind::Text,
    }
}

/// The list holding `text`, and where in it.
fn in_list(text: &str) -> Option<(usize, usize)> {
    LISTS.iter().enumerate().find_map(|(which, list)| {
        list.iter().position(|one| *one == text).map(|at| (which, at))
    })
}

/// Text ending in digits, split before them: `Item 12` is `("Item ", "12")`.
/// The digits have to be at the END: `2026年度` has none there.
fn tail_digits(text: &str) -> Option<(&str, &str)> {
    let cut = text.trim_end_matches(|c: char| c.is_ascii_digit()).len();
    if cut == text.len() {
        return None;
    }
    Some((&text[..cut], &text[cut..]))
}

fn wrapped(index: i64, length: usize) -> usize {
    let length = length as i64;
    (((index % length) + length) % length) as usize
}

/// The least-squares line through `values` at 0, 1, 2, …: Excel fits a line
/// rather than taking the last gap, so 1, 2, 4 continues 5.33, not 6.
fn fit_line(values: &[f64]) -> (f64, f64) {
    let n = values.len() as f64;
    if values.len() == 1 {
        return (0.0, values[0]);
    }
    let (mut sx, mut sy, mut sxy, mut sxx) = (0.0, 0.0, 0.0, 0.0);
    for (i, y) in values.iter().enumerate() {
        let x = i as f64;
        sx += x;
        sy += y;
        sxy += x * y;
        sxx += x * x;
    }
    let under = n * sxx - sx * sx;
    let slope = if under == 0.0 { 0.0 } else { (n * sxy - sx * sy) / under };
    (slope, (sy - slope * sx) / n)
}

/// What a number's format says it IS, as the step it moves by when filled on
/// its own: a `y` or a `d` outside quotes makes it a date (a day), an `h` or
/// an `s` without them a time (an hour). An `m` alone is no help.
fn unit_of(number_format: Option<&str>) -> Option<f64> {
    let format = number_format?;
    let mut bare = String::new();
    let mut quoted = false;
    let mut escaped = false;
    for c in format.chars() {
        if escaped {
            escaped = false;
            continue;
        }
        match c {
            '"' => quoted = !quoted,
            '\\' if !quoted => escaped = true,
            c if !quoted => bare.push(c.to_ascii_lowercase()),
            _ => {}
        }
    }
    if bare.contains('y') || bare.contains('d') {
        Some(1.0)
    } else if bare.contains('h') || bare.contains('s') {
        Some(1.0 / 24.0)
    } else {
        None
    }
}

// Excel counts days from the last day of 1899, so that 1 January 1900 is day
// one (and day 60 is the 29 February 1900 that never was, which nothing here
// reaches).
fn civil_from_days(days: i64) -> (i64, u32, u32) {
    // Days since 1970-01-01 -> proleptic Gregorian, by Howard Hinnant's way.
    let z = days + 719_468;
    let era = if z >= 0 { z } else { z - 146_096 } / 146_097;
    let doe = z - era * 146_097;
    let yoe = (doe - doe / 1460 + doe / 36_524 - doe / 146_096) / 365;
    let y = yoe + era * 400;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let d = (doy - (153 * mp + 2) / 5 + 1) as u32;
    let m = if mp < 10 { mp + 3 } else { mp - 9 } as u32;
    (if m <= 2 { y + 1 } else { y }, m, d)
}

fn days_from_civil(y: i64, m: u32, d: u32) -> i64 {
    let y = if m <= 2 { y - 1 } else { y };
    let era = if y >= 0 { y } else { y - 399 } / 400;
    let yoe = y - era * 400;
    let mp = if m > 2 { m - 3 } else { m + 9 } as i64;
    let doy = (153 * mp + 2) / 5 + i64::from(d) - 1;
    let doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
    era * 146_097 + doe - 719_468
}

/// The serial's date as (year, month, day).
fn date_of(serial: f64) -> (i64, u32, u32) {
    // Serial 25569 is 1970-01-01.
    civil_from_days(serial as i64 - 25_569)
}

fn serial_of(year: i64, month: u32, day: u32) -> f64 {
    (days_from_civil(year, month, day) + 25_569) as f64
}

fn month_length(year: i64, month: u32) -> u32 {
    let (next_year, next_month) = if month == 12 { (year + 1, 1) } else { (year, month + 1) };
    (days_from_civil(next_year, next_month, 1) - days_from_civil(year, month, 1)) as u32
}

/// A calendar-month series: from which month, by how many, on which day.
#[derive(Debug, Clone, Copy)]
struct MonthPlan {
    from: i64,
    step: i64,
    day: u32,
}

/// The calendar-month series running through `serials`, or None.
///
/// The dates must all fall on the same day of the month, or all be the last
/// day of theirs -- a 29-day gap that happens to end on 28 February is not a
/// month -- and the months between them must be evenly spaced, by some number
/// that is not zero.
fn month_series(serials: &[f64]) -> Option<MonthPlan> {
    if serials.len() < 2 || serials.iter().any(|one| one.fract() != 0.0) {
        return None;
    }
    let dates: Vec<(i64, u32, u32)> = serials.iter().map(|one| date_of(*one)).collect();
    let same_day = dates.iter().all(|one| one.2 == dates[0].2);
    let all_ends = dates.iter().all(|(y, m, d)| *d == month_length(*y, *m));
    if !same_day && !all_ends {
        return None;
    }
    let months: Vec<i64> = dates.iter().map(|(y, m, _)| y * 12 + i64::from(*m) - 1).collect();
    let step = months[1] - months[0];
    if step == 0 || months.windows(2).any(|pair| pair[1] - pair[0] != step) {
        return None;
    }
    Some(MonthPlan { from: months[0], step, day: dates[0].2 })
}

/// The serial a month series holds at `index`, counting from its first date:
/// the first date's day of the month, as far as this month allows.
fn month_at(plan: MonthPlan, index: i64) -> f64 {
    let reach = plan.from + plan.step * index;
    let year = reach.div_euclid(12);
    let month = reach.rem_euclid(12) as u32 + 1;
    serial_of(year, month, plan.day.min(month_length(year, month)))
}

/// Round to Excel's fifteen significant digits, so that a fitted line does
/// not show binary tails.
fn fifteen(value: f64) -> f64 {
    if value == 0.0 || !value.is_finite() {
        return value;
    }
    format!("{value:.14e}").parse().unwrap_or(value)
}

/// How one run carries on, worked out once.
#[derive(Debug, Clone)]
enum Plan {
    Line { slope: f64, base: f64 },
    Growth { ratio: f64, base: f64 },
    Months(MonthPlan),
    Weekdays { from: f64, step: i64 },
    List { list: usize, at: usize, stride: i64 },
    Counted { before: String, width: usize, padded: bool, slope: f64, base: f64 },
    Repeat(Vec<Filled>),
    Formula(Vec<Option<String>>),
}

struct Run {
    kind: Kind,
    at: usize,
    len: usize,
}

fn runs_of(line: &[Seed]) -> Vec<Run> {
    let mut runs: Vec<Run> = Vec::new();
    for (at, seed) in line.iter().enumerate() {
        let kind = kind_of(seed);
        match runs.last_mut() {
            Some(last) if last.kind == kind => last.len += 1,
            _ => runs.push(Run { kind, at, len: 1 }),
        }
    }
    runs
}

fn plan_run(run: &Run, line: &[Seed], alone: bool, how: How) -> Plan {
    let cells = &line[run.at..run.at + run.len];
    let repeat = || {
        Plan::Repeat(
            cells
                .iter()
                .map(|seed| match seed.number {
                    Some(number) => Filled::Number(number),
                    None if seed.text.is_empty() => Filled::Empty,
                    None => Filled::Text(seed.text.clone()),
                })
                .collect(),
        )
    };
    if how == How::Copy {
        return match run.kind {
            Kind::Formula => Plan::Formula(cells.iter().map(|seed| seed.formula.clone()).collect()),
            _ => repeat(),
        };
    }
    match &run.kind {
        Kind::Formula => Plan::Formula(cells.iter().map(|seed| seed.formula.clone()).collect()),
        Kind::Number => {
            let values: Vec<f64> = cells.iter().map(|seed| seed.number.unwrap_or(0.0)).collect();
            let unit = unit_of(cells[0].number_format.as_deref());
            let is_date = unit == Some(1.0);
            match how {
                How::Days | How::Weekdays | How::Months | How::Years if !is_date => repeat(),
                How::Days => {
                    let (slope, base) = fit_line(&values);
                    let slope = if values.len() == 1 { 1.0 } else { slope };
                    Plan::Line { slope, base }
                }
                How::Weekdays => {
                    let (slope, _) = fit_line(&values);
                    let step = if values.len() == 1 { 1 } else { (slope.round() as i64).max(1) };
                    Plan::Weekdays { from: values[values.len() - 1], step }
                }
                How::Months | How::Years => {
                    let by = if how == How::Years { 12 } else { 1 };
                    let plan = month_series(&values).unwrap_or_else(|| {
                        let (year, month, day) = date_of(values[0]);
                        let step = if values.len() == 1 {
                            by
                        } else {
                            let (y2, m2, _) = date_of(values[1]);
                            ((y2 * 12 + i64::from(m2)) - (year * 12 + i64::from(month))).max(by)
                        };
                        MonthPlan { from: year * 12 + i64::from(month) - 1, step, day }
                    });
                    Plan::Months(plan)
                }
                How::GrowthTrend => {
                    if values.len() == 1 || values.iter().any(|one| *one <= 0.0) {
                        return repeat();
                    }
                    let logs: Vec<f64> = values.iter().map(|one| one.ln()).collect();
                    let (slope, base) = fit_line(&logs);
                    Plan::Growth { ratio: slope.exp(), base: base.exp() }
                }
                How::LinearTrend => {
                    let (slope, base) = fit_line(&values);
                    Plan::Line { slope, base }
                }
                How::Default | How::Series | How::Copy => {
                    // Dates a whole number of months apart are a calendar
                    // series rather than a straight line: 31 Jan and 28 Feb
                    // continue 31 Mar under one rule and 28 Mar under the other.
                    if values.len() > 1 && is_date {
                        if let Some(plan) = month_series(&values) {
                            return Plan::Months(plan);
                        }
                    }
                    let (slope, base) = fit_line(&values);
                    let slope = if values.len() == 1 {
                        // A single number in a block still counts up, by one
                        // -- measured on [2, 'a'], which gives 3 and 4 -- and
                        // alone it is copied, unless it is a date or a time,
                        // which count up by their own unit, or the fill was
                        // asked for a series.
                        match unit {
                            Some(unit) => unit,
                            None if alone && how == How::Default => 0.0,
                            None => 1.0,
                        }
                    } else {
                        slope
                    };
                    Plan::Line { slope, base }
                }
            }
        }
        Kind::List(list) => {
            let found: Vec<usize> = cells.iter().filter_map(|seed| in_list(&seed.text).map(|(_, at)| at)).collect();
            let length = LISTS[*list].len();
            let first = found[0];
            let stride = if found.len() > 1 { wrapped(found[1] as i64 - first as i64, length) as i64 } else { 1 };
            let even = found
                .iter()
                .enumerate()
                .all(|(i, at)| wrapped(*at as i64 - first as i64, length) as i64 == i as i64 * stride);
            if even {
                Plan::List { list: *list, at: first, stride }
            } else {
                repeat()
            }
        }
        Kind::Counted(before) => {
            let parts: Vec<&str> = cells.iter().filter_map(|seed| tail_digits(&seed.text).map(|(_, digits)| digits)).collect();
            let numbers: Vec<f64> = parts.iter().map(|digits| digits.parse::<f64>().unwrap_or(0.0)).collect();
            let (slope, base) = fit_line(&numbers);
            // Numbered text always counts up, even on its own: "Item 1"
            // pulled down gives Item 2.
            let slope = if numbers.len() == 1 { 1.0 } else { slope };
            Plan::Counted {
                before: before.clone(),
                width: parts[0].len(),
                padded: parts[0].starts_with('0'),
                slope,
                base,
            }
        }
        Kind::Text => repeat(),
    }
}

/// What a planned run holds at `index`, counting from its own first cell.
fn run_value(plan: &Plan, index: i64, len: usize, delta: i64, along: Along) -> Filled {
    match plan {
        Plan::Line { slope, base } => Filled::Number(fifteen(base + slope * index as f64)),
        Plan::Growth { ratio, base } => Filled::Number(fifteen(base * ratio.powi(index as i32))),
        Plan::Months(months) => Filled::Number(month_at(*months, index)),
        Plan::Weekdays { from, step } => {
            // Walk from the run's last date, a weekday at a time.
            let beyond = index - (len as i64 - 1);
            let mut day = *from as i64;
            let mut left = beyond * step;
            let forward = left >= 0;
            while left != 0 {
                day += if forward { 1 } else { -1 };
                // Serial 1 was a Sunday in Excel's calendar (day 7 % 7 == 0 is
                // Saturday): weekday = serial % 7, with 0 Saturday, 1 Sunday.
                let weekday = day.rem_euclid(7);
                if weekday != 0 && weekday != 1 {
                    left += if forward { -1 } else { 1 };
                }
            }
            Filled::Number(day as f64)
        }
        Plan::List { list, at, stride } => {
            let held = LISTS[*list];
            Filled::Text(held[wrapped(*at as i64 + index * stride, held.len())].to_string())
        }
        Plan::Counted { before, width, padded, slope, base } => {
            let next = (base + slope * index as f64).round() as i64;
            let digits = next.to_string();
            let digits = if *padded && digits.len() < *width {
                format!("{}{}", "0".repeat(width - digits.len()), digits)
            } else {
                digits
            };
            Filled::Text(format!("{before}{digits}"))
        }
        Plan::Repeat(values) => values[wrapped(index, values.len())].clone(),
        Plan::Formula(formulas) => match &formulas[wrapped(index, formulas.len())] {
            None => Filled::Empty,
            Some(formula) => {
                let (rows, columns) = match along {
                    Along::Rows => (delta, 0),
                    Along::Columns => (0, delta),
                };
                let text = if formula.starts_with('=') { formula.clone() } else { format!("={formula}") };
                // A formula the engine cannot read is left exactly as it was
                // rather than moved by guesswork.
                Filled::Formula(translate_formula_references(&text, rows, columns).unwrap_or(text))
            }
        },
    }
}

/// Everything one line puts into the `steps` cells beyond it, forwards or
/// backwards. Position `q` counts on from the line -- the cell just past it
/// is `q = line.len()`, and pulling backwards runs `q` negative -- and which
/// seat of the line that lands on, and which turn round it is, together say
/// which run answers and at what index.
pub fn continue_line(line: &[Seed], steps: usize, forwards: bool, how: How, along: Along) -> Vec<Step> {
    if line.is_empty() {
        return Vec::new();
    }
    let runs = runs_of(line);
    let plans: Vec<Plan> = runs.iter().map(|run| plan_run(run, line, line.len() == 1, how)).collect();
    let length = line.len() as i64;
    let mut out = Vec::with_capacity(steps);
    for step in 1..=steps as i64 {
        let q = if forwards { length - 1 + step } else { -step };
        let seat = wrapped(q, line.len());
        let turn = q.div_euclid(length);
        let which = runs.iter().position(|run| seat >= run.at && seat < run.at + run.len).unwrap_or(0);
        let run = &runs[which];
        let index = turn * run.len as i64 + (seat as i64 - run.at as i64);
        let source = run.at as i64 + wrapped(index, run.len) as i64;
        let filled = run_value(&plans[which], index, run.len, q - source, along);
        out.push(Step { filled, seat });
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    fn number(value: f64) -> Seed {
        Seed { number: Some(value), ..Default::default() }
    }

    fn dated(value: f64) -> Seed {
        Seed { number: Some(value), number_format: Some("m/d/yyyy".to_string()), ..Default::default() }
    }

    fn text(value: &str) -> Seed {
        Seed { text: value.to_string(), ..Default::default() }
    }

    fn values(steps: &[Step]) -> Vec<Filled> {
        steps.iter().map(|step| step.filled.clone()).collect()
    }

    fn down(line: &[Seed], steps: usize, how: How) -> Vec<Filled> {
        values(&continue_line(line, steps, true, how, Along::Rows))
    }

    /// Measured from VBA (fill.vba, 2026-09-05).
    #[test]
    fn numbers_follow_the_law() {
        // A lone number is copied; asked for a series, it counts.
        assert_eq!(down(&[number(5.0)], 2, How::Default), vec![Filled::Number(5.0), Filled::Number(5.0)]);
        assert_eq!(down(&[number(5.0)], 2, How::Series), vec![Filled::Number(6.0), Filled::Number(7.0)]);
        // Two count on; three follow their fitted line.
        assert_eq!(down(&[number(1.0), number(2.0)], 2, How::Default), vec![Filled::Number(3.0), Filled::Number(4.0)]);
        assert_eq!(
            down(&[number(1.0), number(2.0), number(4.0)], 2, How::Default),
            vec![Filled::Number(5.33333333333333), Filled::Number(6.83333333333333)]
        );
        // Copy repeats; growth multiplies; a linear trend is the line.
        assert_eq!(
            down(&[number(1.0), number(2.0)], 3, How::Copy),
            vec![Filled::Number(1.0), Filled::Number(2.0), Filled::Number(1.0)]
        );
        assert_eq!(down(&[number(2.0), number(4.0)], 2, How::GrowthTrend), vec![Filled::Number(8.0), Filled::Number(16.0)]);
        assert_eq!(down(&[number(1.0), number(4.0)], 2, How::LinearTrend), vec![Filled::Number(7.0), Filled::Number(10.0)]);
        // Pulled upwards, the line runs back: 10, 8 above gives 12, 14, 16.
        let up = continue_line(&[number(10.0), number(8.0)], 3, false, How::Default, Along::Rows);
        assert_eq!(values(&up), vec![Filled::Number(12.0), Filled::Number(14.0), Filled::Number(16.0)]);
        // Days on plain numbers is a copy.
        assert_eq!(down(&[number(1.0), number(2.0)], 2, How::Days), vec![Filled::Number(1.0), Filled::Number(2.0)]);
    }

    #[test]
    fn text_lists_and_counts_follow_the_law() {
        assert_eq!(down(&[text("Item 1")], 2, How::Default), vec![Filled::Text("Item 2".into()), Filled::Text("Item 3".into())]);
        assert_eq!(down(&[text("A001")], 1, How::Default), vec![Filled::Text("A002".into())]);
        assert_eq!(down(&[text("Mon")], 2, How::Default), vec![Filled::Text("Tue".into()), Filled::Text("Wed".into())]);
        assert_eq!(down(&[text("Mon"), text("Wed")], 3, How::Default), vec![Filled::Text("Fri".into()), Filled::Text("Sun".into()), Filled::Text("Tue".into())]);
        assert_eq!(down(&[text("a")], 2, How::Default), vec![Filled::Text("a".into()), Filled::Text("a".into())]);
        // Runs carry on by themselves: 1, a gives 2, a, 3.
        let mixed = continue_line(&[number(1.0), text("a")], 3, true, How::Default, Along::Rows);
        assert_eq!(values(&mixed), vec![Filled::Number(2.0), Filled::Text("a".into()), Filled::Number(3.0)]);
        assert_eq!(mixed.iter().map(|step| step.seat).collect::<Vec<_>>(), vec![0, 1, 0]);
        // A blank in the pattern stays a blank.
        assert_eq!(down(&[text("x"), Seed::default()], 2, How::Default), vec![Filled::Text("x".into()), Filled::Empty]);
    }

    #[test]
    fn dates_follow_the_calendar() {
        let jan31 = 45_322.0;
        // A lone date moves a day; a lone time an hour.
        assert_eq!(down(&[dated(jan31)], 1, How::Default), vec![Filled::Number(jan31 + 1.0)]);
        let ten_thirty = Seed { number: Some(0.4375), number_format: Some("h:mm".to_string()), ..Default::default() };
        assert_eq!(down(&[ten_thirty], 1, How::Default), vec![Filled::Number(fifteen(0.4375 + 1.0 / 24.0))]);
        // Months keep the day as far as each month allows: 31 Jan -> 29 Feb -> 31 Mar.
        assert_eq!(
            down(&[dated(jan31)], 2, How::Months),
            vec![Filled::Number(serial_of(2024, 2, 29)), Filled::Number(serial_of(2024, 3, 31))]
        );
        // Years from 29 Feb 2024: 28 Feb 2025, 28 Feb 2026.
        assert_eq!(
            down(&[dated(serial_of(2024, 2, 29))], 2, How::Years),
            vec![Filled::Number(serial_of(2025, 2, 28)), Filled::Number(serial_of(2026, 2, 28))]
        );
        // Weekdays skip the weekend: Fri 5 Jan 2024 -> Mon 8, Tue 9, Wed 10.
        assert_eq!(
            down(&[dated(serial_of(2024, 1, 5))], 3, How::Weekdays),
            vec![Filled::Number(serial_of(2024, 1, 8)), Filled::Number(serial_of(2024, 1, 9)), Filled::Number(serial_of(2024, 1, 10))]
        );
        // 31 Jan, 28 Feb (2023, when February ends on the 28th) is a month
        // series; 29 Jan, 28 Feb is a 30-day line, though it lands on a
        // month's end.
        assert_eq!(
            down(&[dated(serial_of(2023, 1, 31)), dated(serial_of(2023, 2, 28))], 2, How::Default),
            vec![Filled::Number(serial_of(2023, 3, 31)), Filled::Number(serial_of(2023, 4, 30))]
        );
        assert_eq!(
            down(&[dated(serial_of(2023, 1, 29)), dated(serial_of(2023, 2, 28))], 1, How::Default),
            vec![Filled::Number(serial_of(2023, 3, 30))]
        );
    }

    #[test]
    fn formulas_move_with_their_cell() {
        let seed = Seed { formula: Some("A1*2".to_string()), ..Default::default() };
        assert_eq!(down(&[seed.clone()], 2, How::Default), vec![Filled::Formula("=A2*2".into()), Filled::Formula("=A3*2".into())]);
        let right = continue_line(&[seed], 1, true, How::Default, Along::Columns);
        assert_eq!(values(&right), vec![Filled::Formula("=B1*2".into())]);
    }
}
