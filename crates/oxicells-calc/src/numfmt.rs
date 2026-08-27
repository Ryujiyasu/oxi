// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Turning a cell's value into the text a sheet shows for it.
//!
//! Every expectation in the tests below is what Excel 16 put in `Range.Text`
//! for that value under that format.

use crate::datetime::{date_from_serial, weekday_sunday_one};

/// Renders `value` under `format`, the way a worksheet shows it.
///
/// A format may hold up to four sections, separated by semicolons: what to show
/// for a positive number, a negative one, zero, and text. With one section it
/// covers everything, and a negative number is shown with a minus sign; with
/// two or more, the negative section states its own sign, which is why
/// `#,##0;(#,##0)` shows `(1,235)` rather than `(-1,235)`.
pub fn format_number(value: f64, format: &str) -> String {
    if format.is_empty() || format.eq_ignore_ascii_case("general") {
        return general(value);
    }

    let sections: Vec<&str> = split_sections(format);
    let (section, signed) = if value < 0.0 && sections.len() > 1 {
        // The negative section carries its own sign, so the value loses it.
        (sections[1], false)
    } else if value == 0.0 && sections.len() > 2 {
        (sections[2], false)
    } else {
        (sections[0], true)
    };

    let magnitude = if signed { value } else { value.abs() };
    if looks_like_a_date(section) {
        return format_datetime(magnitude, section);
    }
    format_numeric(magnitude, section)
}

/// Splits on semicolons that are not inside quotes.
fn split_sections(format: &str) -> Vec<&str> {
    let mut sections = Vec::new();
    let mut quoted = false;
    let mut start = 0;
    for (at, character) in format.char_indices() {
        match character {
            '"' => quoted = !quoted,
            ';' if !quoted => {
                sections.push(&format[start..at]);
                start = at + 1;
            }
            _ => {}
        }
    }
    sections.push(&format[start..]);
    sections
}

/// A format is a date format when it names a date or time part outside quotes.
fn looks_like_a_date(format: &str) -> bool {
    let mut quoted = false;
    let mut marked_month = false;
    let mut characters = format.chars().peekable();
    while let Some(character) = characters.next() {
        // A format code is a format code in either case. The formatter has
        // always lowercased before reading one; this test did not, so a format
        // spelled in capitals was taken for a number format and printed its
        // own letters back.
        match character.to_ascii_lowercase() {
            '"' => quoted = !quoted,
            _ if quoted => {}
            // The character after one of these belongs to the directive.
            '_' | '\\' | '*' => {
                characters.next();
            }
            // `[Red]` holds a d, `[$-411]` holds neither, and neither of them
            // is a date part. Reading the d in Red as a day turned every
            // negative number in an accounting format into a date.
            '[' => {
                // A group of one repeated h, m or s is a lump of elapsed time
                // — `[h]:mm` is 36:00 for a day and a half — and it is the
                // only date part a format can consist of entirely. Skipping
                // the group without looking left `[s]` reading as a number and
                // showing 2 where Excel shows 129600.
                let mut inside = String::new();
                for held in characters.by_ref() {
                    if held == ']' {
                        break;
                    }
                    inside.push(held.to_ascii_lowercase());
                }
                let unit = inside.chars().next().unwrap_or(' ');
                if matches!(unit, 'h' | 'm' | 's') && inside.chars().all(|held| held == unit) {
                    return true;
                }
            }
            'y' | 'd' | 'h' | 's' => return true,
            // `m` is a month beside the others and minutes beside an `h`, and
            // on its own — `"mmmm"`, the month's name — it is still a date.
            // It cannot simply be added to the line above: `m` is also an
            // ordinary letter, and a number format is free to contain one.
            // What a number format is never free to contain is a month code
            // AND no digit at all, so the two are told apart by that.
            'm' => marked_month = true,
            '0' | '#' | '?' => return false,
            _ => {}
        }
    }
    marked_month
}

/// What `General` shows: the shortest text that reads back as the same number.
fn general(value: f64) -> String {
    if value == value.trunc() && value.abs() < 1e15 {
        return format!("{}", value as i64);
    }
    let text = format!("{value}");
    text
}

fn format_numeric(value: f64, format: &str) -> String {
    let mut decimals = 0usize;
    let mut grouped = false;
    let mut percent = false;
    let mut scientific = false;
    let mut scale = 0u32;
    let mut integer_places = 0usize;

    // Read the shape of the format before rendering anything with it.
    let body: Vec<char> = format.chars().collect();
    let mut seen_point = false;
    let mut quoted = false;
    let mut at = 0;
    while at < body.len() {
        let character = body[at];
        match character {
            '"' => quoted = !quoted,
            _ if quoted => {}
            // `_x` keeps the width of x, `\x` shows x, `*x` fills with x:
            // in each the character after belongs to the directive, not to
            // the number.
            '_' | '\\' | '*' => at += 1,
            // A bracketed part names a colour, a condition or a locale.
            '[' => {
                while at < body.len() && body[at] != ']' {
                    at += 1;
                }
            }
            '.' => seen_point = true,
            '0' | '#' | '?' => {
                if seen_point {
                    decimals += 1;
                } else if character == '0' {
                    integer_places += 1;
                }
            }
            '%' => percent = true,
            'E' | 'e' if at + 1 < body.len() && matches!(body[at + 1], '+' | '-') => {
                scientific = true;
            }
            ',' => {
                // A comma among the digits groups them; one after the last
                // digit divides by a thousand for each comma.
                let trailing = body[at + 1..]
                    .iter()
                    .all(|held| !matches!(held, '0' | '#' | '?' | '.'));
                if trailing {
                    scale += 1;
                } else {
                    grouped = true;
                }
            }
            _ => {}
        }
        at += 1;
    }

    let mut number = value;
    if percent {
        number *= 100.0;
    }
    for _ in 0..scale {
        number /= 1000.0;
    }

    if scientific {
        // In 0.00E+00 the places after the point belong to the mantissa; the
        // ones after E are the exponent's width, counted separately.
        let mantissa_places = format
            .split(['E', 'e'])
            .next()
            .map(|head| head.rsplit('.').next().unwrap_or("").chars()
                .filter(|held| matches!(held, '0' | '#' | '?')).count())
            .filter(|_| format.contains('.'))
            .unwrap_or(0);
        let rendered = format!("{:.*E}", mantissa_places, number);
        // Rust writes E3; Excel writes E+03.
        return match rendered.split_once('E') {
            Some((mantissa, exponent)) => {
                let power: i32 = exponent.parse().unwrap_or(0);
                format!("{mantissa}E{}{:02}", if power < 0 { '-' } else { '+' }, power.abs())
            }
            None => rendered,
        };
    }

    let negative = number < 0.0;
    // Excel sends a half away from zero; Rust's formatting sends it to the
    // even neighbour, so the rounding is done before the text is made.
    let scale = 10f64.powi(decimals as i32);
    let settled = (number.abs() * scale).round() / scale;
    let rounded = format!("{:.*}", decimals, settled);
    let (whole, fraction) = match rounded.split_once('.') {
        Some((whole, fraction)) => (whole.to_string(), Some(fraction.to_string())),
        None => (rounded.clone(), None),
    };

    let mut whole = whole;
    while whole.len() < integer_places {
        whole.insert(0, '0');
    }
    if grouped {
        whole = group_thousands(&whole);
    }

    let mut rendered = String::new();
    if negative {
        rendered.push('-');
    }
    rendered.push_str(&whole);
    if let Some(fraction) = fraction {
        rendered.push('.');
        rendered.push_str(&fraction);
    }

    // Whatever the format says around the number comes along: currency signs,
    // quoted words, percent signs.
    decorate(&rendered, format)
}

fn group_thousands(digits: &str) -> String {
    let mut grouped = String::new();
    for (at, digit) in digits.chars().enumerate() {
        if at > 0 && (digits.len() - at).is_multiple_of(3) {
            grouped.push(',');
        }
        grouped.push(digit);
    }
    grouped
}

/// Puts the number into the literal text the format wraps it in.
fn decorate(number: &str, format: &str) -> String {
    let mut before = String::new();
    let mut after = String::new();
    let mut seen_digit = false;
    let mut quoted = false;
    let mut characters = format.chars().peekable();
    while let Some(character) = characters.next() {
        match character {
            '"' => quoted = !quoted,
            _ if quoted => {
                if seen_digit {
                    after.push(character);
                } else {
                    before.push(character);
                }
            }
            '0' | '#' | '?' | '.' | ',' => seen_digit = true,
            '%' => {
                if seen_digit {
                    after.push('%');
                } else {
                    before.push('%');
                }
            }
            '\\' => {
                if let Some(escaped) = characters.next() {
                    if seen_digit {
                        after.push(escaped);
                    } else {
                        before.push(escaped);
                    }
                }
            }
            // `_x` asks for the width of x and shows nothing there. Excel's
            // own `Range.Text` gives a space, which is what a sheet shows.
            '_' => {
                characters.next();
                if seen_digit {
                    after.push(' ');
                } else {
                    before.push(' ');
                }
            }
            // `*x` fills the rest of the cell with x. What that comes to
            // depends on the width of the cell, so nothing is put here.
            '*' => {
                characters.next();
            }
            // `[Red]`, `[$-411]`, `[>100]`: a colour, a locale, a condition.
            // None of them is text.
            '[' => {
                for held in characters.by_ref() {
                    if held == ']' {
                        break;
                    }
                }
            }
            _ => {
                if seen_digit {
                    after.push(character);
                } else {
                    before.push(character);
                }
            }
        }
    }
    format!("{before}{number}{after}")
}

fn format_datetime(serial: f64, format: &str) -> String {
    let whole = serial.trunc() as i64;
    let Ok(date) = date_from_serial(whole) else {
        return general(serial);
    };
    let (year, month, day) = (date.year, date.month, date.day);
    // The fraction of a day is the time, rounded to the nearest second the way
    // Excel shows it.
    let seconds_of_day = ((serial - serial.trunc()) * 86_400.0).round() as i64;
    let hour = seconds_of_day / 3600;
    let minute = (seconds_of_day % 3600) / 60;
    let second = seconds_of_day % 60;

    const DAYS: [&str; 7] = [
        "Sunday",
        "Monday",
        "Tuesday",
        "Wednesday",
        "Thursday",
        "Friday",
        "Saturday",
    ];
    const MONTHS: [&str; 12] = [
        "January",
        "February",
        "March",
        "April",
        "May",
        "June",
        "July",
        "August",
        "September",
        "October",
        "November",
        "December",
    ];

    // An `AM/PM` in the format puts the clock on twelve hours and prints the
    // marker where it stands. Without one every hour runs to twenty-four.
    let marker = am_pm_marker(format);
    let (hour, meridiem) = match &marker {
        Some(written) => (
            // Midnight and noon are both twelve o'clock.
            match hour % 12 {
                0 => 12,
                other => other,
            },
            if hour < 12 {
                written.before.as_str()
            } else {
                written.after.as_str()
            },
        ),
        None => (hour, ""),
    };
    // weekday_sunday_one counts Sunday as one; these tables start at zero.
    let weekday = (weekday_sunday_one(whole) - 1).clamp(0, 6) as usize;
    // Elapsed time counts from the epoch rather than from midnight, so a
    // `[h]` is 1087128 where an `h` is 0. Excel rounds the whole serial to
    // the nearest second once, and every elapsed part reads off that.
    let total_seconds = (serial * 86_400.0).round() as i64;
    let (era_latin, era_short, era_full, era_year) = era(year, month, day);

    let body: Vec<char> = format.chars().collect();
    let mut rendered = String::new();
    let mut at = 0;
    let mut quoted = false;
    while at < body.len() {
        let character = body[at];
        if character == '"' {
            quoted = !quoted;
            at += 1;
            continue;
        }
        if quoted {
            rendered.push(character);
            at += 1;
            continue;
        }
        // A backslash shows the next character as itself: `yyyy\-mm` is a date
        // with hyphens in it, not a date with backslashes in it.
        if character == '\\' {
            if let Some(next) = body.get(at + 1) {
                rendered.push(*next);
            }
            at += 2;
            continue;
        }
        // `_)` reserves the width of a `)` without drawing it, and `*` fills
        // the rest of the cell with a character. Neither is text.
        if character == '_' {
            rendered.push(' ');
            at += 2;
            continue;
        }
        if character == '*' {
            at += 2;
            continue;
        }
        // A bracket group is either a lump of elapsed time — `[h]`, `[mm]`,
        // `[s]` — or something for the renderer rather than the reader, like
        // the locale tag `[$-411]` or the colour `[Red]`.
        if character == '[' {
            let close = body[at..]
                .iter()
                .position(|held| *held == ']')
                .map(|found| at + found);
            let Some(close) = close else {
                at += 1;
                continue;
            };
            let inside: String = body[at + 1..close].iter().collect();
            let lower = inside.to_ascii_lowercase();
            let unit = lower.chars().next().unwrap_or(' ');
            if !lower.is_empty() && lower.chars().all(|held| held == unit) {
                let elapsed = match unit {
                    'h' => Some(total_seconds / 3600),
                    'm' => Some(total_seconds / 60),
                    's' => Some(total_seconds),
                    _ => None,
                };
                if let Some(elapsed) = elapsed {
                    rendered.push_str(&format!("{:0width$}", elapsed, width = lower.len()));
                    at = close + 1;
                    continue;
                }
            }
            at = close + 1;
            continue;
        }
        let body_at = at;
        let run = body[at..]
            .iter()
            .take_while(|held| held.eq_ignore_ascii_case(&character))
            .count();
        let lower = character.to_ascii_lowercase();
        match lower {
            'y' => {
                if run >= 4 {
                    rendered.push_str(&format!("{year:04}"));
                } else {
                    rendered.push_str(&format!("{:02}", year % 100));
                }
            }
            'd' => match run {
                1 => rendered.push_str(&day.to_string()),
                2 => rendered.push_str(&format!("{day:02}")),
                3 => rendered.push_str(&DAYS[weekday][..3]),
                _ => rendered.push_str(DAYS[weekday]),
            },
            'h' => {
                if run >= 2 {
                    rendered.push_str(&format!("{hour:02}"));
                } else {
                    rendered.push_str(&hour.to_string());
                }
            }
            's' => {
                if run >= 2 {
                    rendered.push_str(&format!("{second:02}"));
                } else {
                    rendered.push_str(&second.to_string());
                }
            }
            'm' => {
                // An m after an hour, or before seconds, means minutes.
                let minutes = previous_was_hour(&body, at) || next_is_second(&body, at + run);
                let value = if minutes { minute } else { month };
                match (minutes, run) {
                    (false, 3) => rendered.push_str(&MONTHS[(month - 1) as usize][..3]),
                    (false, n) if n >= 4 => rendered.push_str(MONTHS[(month - 1) as usize]),
                    (_, 1) => rendered.push_str(&value.to_string()),
                    _ => rendered.push_str(&format!("{value:02}")),
                }
            }
            // The marker itself, printed as AM or PM whatever case it was
            // written in.
            'a' | 'p'
                if marker
                    .as_ref()
                    .is_some_and(|written| written.at == body_at) =>
            {
                rendered.push_str(meridiem);
                at += marker.as_ref().map_or(0, |written| written.len);
                continue;
            }
            'e' => rendered.push_str(&era_year.to_string()),
            'g' => rendered.push_str(match run {
                1 => era_latin,
                2 => era_short,
                _ => era_full,
            }),
            _ => {
                rendered.push(character);
                at += 1;
                continue;
            }
        }
        at += run;
    }
    rendered
}

/// The Japanese era a date falls in: its name three ways, and which year of it
/// this is.
///
/// The four boundaries are the days the era changed, and each was measured by
/// asking Excel for the day before and the day of — a boundary in the wrong
/// place shows up as the two sides agreeing. Nothing before Meiji is reachable,
/// since Excel's own day one is 1900.
fn era(year: i64, month: i64, day: i64) -> (&'static str, &'static str, &'static str, i64) {
    let ymd = (year, month, day);
    let (latin, short, full, from) = if ymd >= (2019, 5, 1) {
        ("R", "令", "令和", 2019)
    } else if ymd >= (1989, 1, 8) {
        ("H", "平", "平成", 1989)
    } else if ymd >= (1926, 12, 25) {
        ("S", "昭", "昭和", 1926)
    } else if ymd >= (1912, 7, 30) {
        ("T", "大", "大正", 1912)
    } else {
        ("M", "明", "明治", 1868)
    };
    (latin, short, full, year - from + 1)
}

/// The marker that puts a clock on twelve hours, as written.
struct Meridiem {
    /// Where it starts in the format, counted in characters.
    at: usize,
    /// How many characters it takes up.
    len: usize,
    /// What to print before noon, and after — copied out of the format as they
    /// stand, since that is what Excel prints.
    before: String,
    after: String,
}

/// The `AM/PM` or `A/P` in a format, if it has one.
///
/// Both halves are kept exactly as they were typed: `AM/pm` prints `AM` in the
/// morning and `pm` in the afternoon, so there is no rule about capitals to
/// apply — only text to copy. Anything else starting with an a or a p is
/// ordinary text.
fn am_pm_marker(format: &str) -> Option<Meridiem> {
    let body: Vec<char> = format.chars().collect();
    let mut quoted = false;
    for at in 0..body.len() {
        if body[at] == '"' {
            quoted = !quoted;
            continue;
        }
        if quoted {
            continue;
        }
        for (long, split) in [(5, 2), (3, 1)] {
            if at + long > body.len() {
                continue;
            }
            let held: String = body[at..at + long].iter().collect();
            let spelling = if long == 5 { "am/pm" } else { "a/p" };
            if !held.eq_ignore_ascii_case(spelling) {
                continue;
            }
            return Some(Meridiem {
                at,
                len: long,
                before: held[..split].to_string(),
                after: held[split + 1..].to_string(),
            });
        }
    }
    None
}

fn previous_was_hour(body: &[char], at: usize) -> bool {
    body[..at]
        .iter()
        .rev()
        .find(|held| held.is_ascii_alphabetic())
        .is_some_and(|held| held.eq_ignore_ascii_case(&'h'))
}

fn next_is_second(body: &[char], at: usize) -> bool {
    body[at..]
        .iter()
        .find(|held| held.is_ascii_alphabetic())
        .is_some_and(|held| held.eq_ignore_ascii_case(&'s'))
}

#[cfg(test)]
mod tests {
    use super::format_number;

    /// Every expectation is what Excel 16 put in `Range.Text` for that value
    /// under that format.
    /// A format's spacing, fill and bracket parts are instructions to the
    /// renderer, not text to show. `#,##0.0_);(#,##0.0)` is what the machinery
    /// statistics are written with, and printing the `_)` leaves every number
    /// on the sheet with a stray bracket.
    /// `m` is the one code that means two things, and it means a third thing
    /// again when nothing else is there to tell them apart.
    ///
    /// In `h:mm` it is minutes, in `d-mmm` it is a month, and in `mmmm` on its
    /// own — the month's name, which is a perfectly ordinary way to label a
    /// column — there is nothing beside it at all. Reading `m` as "not a date"
    /// left `TEXT(45297,"mmmm")` printing `mmmm45297`: the format was taken for
    /// a number format, so its letters were passed through as literal text and
    /// the serial was appended as though it were the number.
    ///
    /// What a number format never contains is a month code and no digit at
    /// all, so a digit placeholder anywhere is what rules the format out.
    /// Every one of these is what Excel 16 put in `Range.Text` for that value
    /// under that format, read off a column wide enough to show it — a narrow
    /// column answers `#######` however right the format is.
    ///
    /// Three things live here that the date formatter had never met, and all
    /// three arrived together because allowing `m` to mark a date sent six of
    /// the corpus formats down this path for the first time:
    ///
    ///   `[h]` and `[m]` and `[s]` are elapsed time, counted from the epoch
    ///   rather than from midnight, so a day and a half is 36 hours and not 12.
    ///
    ///   `e` is which year of the Japanese era it is and `g` the era's name,
    ///   as its Latin initial, one kanji, or in full. The four boundaries were
    ///   each measured on the day before and the day of the change.
    ///
    ///   A backslash shows the next character as itself, so `yyyy\-mm\-dd`
    ///   is a date with hyphens rather than one with backslashes.
    #[test]
    fn brackets_eras_and_escapes_are_what_excel_shows() {
        assert_eq!(format_number(45297.0, "mmmm"), "January");
        assert_eq!(format_number(1.5, "mmmm"), "January");
        assert_eq!(format_number(0.25, "mmmm"), "January");
        assert_eq!(format_number(45297.75, "mmmm"), "January");
        assert_eq!(format_number(45297.0, "mmm"), "Jan");
        assert_eq!(format_number(1.5, "mmm"), "Jan");
        assert_eq!(format_number(0.25, "mmm"), "Jan");
        assert_eq!(format_number(45297.75, "mmm"), "Jan");
        assert_eq!(format_number(45297.0, "mm"), "01");
        assert_eq!(format_number(1.5, "mm"), "01");
        assert_eq!(format_number(0.25, "mm"), "01");
        assert_eq!(format_number(45297.75, "mm"), "01");
        assert_eq!(format_number(45297.0, "[$-411]\"(\"e\"年\"m\"月分)\""), "(6年1月分)");
        assert_eq!(format_number(1.5, "[$-411]\"(\"e\"年\"m\"月分)\""), "(33年1月分)");
        assert_eq!(format_number(0.25, "[$-411]\"(\"e\"年\"m\"月分)\""), "(33年1月分)");
        assert_eq!(format_number(45297.75, "[$-411]\"(\"e\"年\"m\"月分)\""), "(6年1月分)");
        assert_eq!(format_number(45297.0, "\"（\"[$-411]e\"年\"m\"月）\""), "（6年1月）");
        assert_eq!(format_number(1.5, "\"（\"[$-411]e\"年\"m\"月）\""), "（33年1月）");
        assert_eq!(format_number(0.25, "\"（\"[$-411]e\"年\"m\"月）\""), "（33年1月）");
        assert_eq!(format_number(45297.75, "\"（\"[$-411]e\"年\"m\"月）\""), "（6年1月）");
        assert_eq!(format_number(45297.0, "[h]:mm"), "1087128:00");
        assert_eq!(format_number(1.5, "[h]:mm"), "36:00");
        assert_eq!(format_number(0.25, "[h]:mm"), "6:00");
        assert_eq!(format_number(45297.75, "[h]:mm"), "1087146:00");
        assert_eq!(format_number(45297.0, "[h]:mm;@"), "1087128:00");
        assert_eq!(format_number(1.5, "[h]:mm;@"), "36:00");
        assert_eq!(format_number(0.25, "[h]:mm;@"), "6:00");
        assert_eq!(format_number(45297.75, "[h]:mm;@"), "1087146:00");
        assert_eq!(format_number(45297.0, "[mm]:ss"), "65227680:00");
        assert_eq!(format_number(1.5, "[mm]:ss"), "2160:00");
        assert_eq!(format_number(0.25, "[mm]:ss"), "360:00");
        assert_eq!(format_number(45297.75, "[mm]:ss"), "65228760:00");
        assert_eq!(format_number(45297.0, "[s]"), "3913660800");
        assert_eq!(format_number(1.5, "[s]"), "129600");
        assert_eq!(format_number(0.25, "[s]"), "21600");
        assert_eq!(format_number(45297.75, "[s]"), "3913725600");
        assert_eq!(format_number(45297.0, "[m]"), "65227680");
        assert_eq!(format_number(1.5, "[m]"), "2160");
        assert_eq!(format_number(0.25, "[m]"), "360");
        assert_eq!(format_number(45297.75, "[m]"), "65228760");
        assert_eq!(format_number(45297.0, "\"(\"[$-409]mmm\\,\\ yyyy\")\""), "(Jan, 2024)");
        assert_eq!(format_number(1.5, "\"(\"[$-409]mmm\\,\\ yyyy\")\""), "(Jan, 1900)");
        assert_eq!(format_number(0.25, "\"(\"[$-409]mmm\\,\\ yyyy\")\""), "(Jan, 1900)");
        assert_eq!(format_number(45297.75, "\"(\"[$-409]mmm\\,\\ yyyy\")\""), "(Jan, 2024)");
        assert_eq!(format_number(45297.0, "yyyy\\-mm\\-dd"), "2024-01-06");
        assert_eq!(format_number(1.5, "yyyy\\-mm\\-dd"), "1900-01-01");
        assert_eq!(format_number(0.25, "yyyy\\-mm\\-dd"), "1900-01-00");
        assert_eq!(format_number(45297.75, "yyyy\\-mm\\-dd"), "2024-01-06");
        assert_eq!(format_number(45297.0, "[$-409]d\\-mmm"), "6-Jan");
        assert_eq!(format_number(1.5, "[$-409]d\\-mmm"), "1-Jan");
        assert_eq!(format_number(0.25, "[$-409]d\\-mmm"), "0-Jan");
        assert_eq!(format_number(45297.75, "[$-409]d\\-mmm"), "6-Jan");
        assert_eq!(format_number(45297.0, "ggge\"年\"m\"月\"d\"日\""), "令和6年1月6日");
        assert_eq!(format_number(1.5, "ggge\"年\"m\"月\"d\"日\""), "明治33年1月1日");
        assert_eq!(format_number(0.25, "ggge\"年\"m\"月\"d\"日\""), "明治33年1月0日");
        assert_eq!(format_number(45297.75, "ggge\"年\"m\"月\"d\"日\""), "令和6年1月6日");
        assert_eq!(format_number(45297.0, "ge.m.d"), "R6.1.6");
        assert_eq!(format_number(1.5, "ge.m.d"), "M33.1.1");
        assert_eq!(format_number(0.25, "ge.m.d"), "M33.1.0");
        assert_eq!(format_number(45297.75, "ge.m.d"), "R6.1.6");
        assert_eq!(format_number(45297.0, "yyyy\"年\"m\"月\""), "2024年1月");
        assert_eq!(format_number(1.5, "yyyy\"年\"m\"月\""), "1900年1月");
        assert_eq!(format_number(0.25, "yyyy\"年\"m\"月\""), "1900年1月");
        assert_eq!(format_number(45297.75, "yyyy\"年\"m\"月\""), "2024年1月");
    }

    /// A format code is a format code in either case.
    ///
    /// The formatter has always lowercased before reading one, but the test
    /// that decides whether a format IS a date only looked at lower-case
    /// letters. So `TEXT(F4,"DD")` fell through to the number path, where the
    /// letters are literal text and the serial is printed after them —
    /// `DD42298`.
    #[test]
    fn a_format_code_does_not_care_about_capitals() {
        assert_eq!(format_number(42298.0, "DD"), "21");
        assert_eq!(format_number(42298.0, "dd"), "21");
        assert_eq!(format_number(42298.0, "MMM YY"), "Oct 15");
        assert_eq!(format_number(42298.0, "mmm yy"), "Oct 15");
        assert_eq!(format_number(42298.0, "YYYY-MM-DD"), "2015-10-21");
        assert_eq!(format_number(42298.0, "HH:MM:SS"), "00:00:00");
    }

    /// An `AM/PM` puts the clock on twelve hours, and the half of the marker
    /// that applies is printed EXACTLY as it was typed.
    ///
    /// Every line is Excel 16's. The two mixed-case ones are what settle the
    /// rule: with the halves written differently, each output takes the case of
    /// its own half — so there is nothing about capitals to decide, only text
    /// to copy. Guessing "always AM or PM" would have passed the first four
    /// lines and been wrong.
    #[test]
    fn the_half_of_the_marker_that_applies_prints_as_it_was_typed() {
        assert_eq!(format_number(0.5, "H:MM AM/PM"), "12:00 PM");
        assert_eq!(format_number(0.5, "h AM/PM"), "12 PM");
        assert_eq!(format_number(0.75, "h:mm AM/PM"), "6:00 PM");
        assert_eq!(format_number(0.25, "h:mm AM/PM"), "6:00 AM");
        // Short form: one letter, not two.
        assert_eq!(format_number(0.75, "h:mm A/P"), "6:00 P");
        assert_eq!(format_number(0.25, "h:mm A/P"), "6:00 A");
        assert_eq!(format_number(0.75, "h:mm a/p"), "6:00 p");
        // Written small, printed small.
        assert_eq!(format_number(0.75, "h:mm am/pm"), "6:00 pm");
        assert_eq!(format_number(0.25, "h:mm am/pm"), "6:00 am");
        assert_eq!(format_number(0.75, "h:mm Am/Pm"), "6:00 Pm");
        // Mixed: each half its own.
        assert_eq!(format_number(0.75, "h:mm AM/pm"), "6:00 pm");
        assert_eq!(format_number(0.25, "h:mm AM/pm"), "6:00 AM");
        assert_eq!(format_number(0.25, "h:mm am/PM"), "6:00 am");
        // Without a marker the clock runs to twenty-four.
        assert_eq!(format_number(0.75, "h:mm"), "18:00");
    }

    #[test]
    fn a_month_on_its_own_is_still_a_date() {
        assert_eq!(format_number(45297.0, "mmmm"), "January");
        assert_eq!(format_number(45297.0, "mmm"), "Jan");
        assert_eq!(format_number(45297.0, "mm"), "01");
        // Beside the codes that were already recognised, unchanged.
        assert_eq!(format_number(45297.0, "yyyy-mm-dd"), "2024-01-06");
        // A number format with an `m` in its currency text is not a date, and
        // it is the digits that say so.
        assert_eq!(format_number(1234.5, "0.00"), "1234.50");
        assert_eq!(format_number(1234.5, "#,##0"), "1,235");
    }

    #[test]
    fn spacing_and_colour_are_not_text() {
        assert_eq!(format_number(105.3, "#,##0.0_);(#,##0.0)"), "105.3 ");
        assert_eq!(format_number(24493.0, "#,##0 ;[Red](#,##0)"), "24,493 ");
        assert_eq!(format_number(-24493.0, "#,##0 ;[Red](#,##0)"), "(24,493)");
        assert_eq!(format_number(5.0, "[Blue]0"), "5");
        // A fill takes the width of the cell, which text cannot say.
        assert_eq!(format_number(7.0, "0*-"), "7");
        // An escaped character is still shown.
        assert_eq!(format_number(7.0, r"0\%"), "7%");
    }

    #[test]
    fn numbers_render_the_way_excel_shows_them() {
        for (value, format, shown) in [
            (1234.5, "General", "1234.5"),
            (1234.5, "0", "1235"),
            (1234.5, "0.00", "1234.50"),
            (1234.5, "#,##0", "1,235"),
            (1234.5, "#,##0.00", "1,234.50"),
            (0.25, "0%", "25%"),
            (0.25, "0.00%", "25.00%"),
            (1234.5, "0.00E+00", "1.23E+03"),
            (-1234.5, "#,##0.00", "-1,234.50"),
            (-1234.5, "0", "-1235"),
            (0.0, "0.00", "0.00"),
            (0.125, "0.000", "0.125"),
            (12.0, "00000", "00012"),
            (1234.5, "$#,##0.00", "$1,234.50"),
            (1234567.0, "#,##0,", "1,235"),
        ] {
            assert_eq!(format_number(value, format), shown, "{value} as {format}");
        }
    }

    /// A half goes away from zero, not to the even neighbour.
    #[test]
    fn a_half_rounds_away_from_zero() {
        for (value, shown) in [(0.5, "1"), (1.5, "2"), (2.5, "3"), (-0.5, "-1")] {
            assert_eq!(format_number(value, "0"), shown, "{value}");
        }
    }

    /// With more than one section the negative one states its own sign.
    #[test]
    fn a_second_section_takes_the_negatives() {
        assert_eq!(format_number(1234.5, "#,##0;(#,##0)"), "1,235");
        assert_eq!(format_number(-1234.5, "#,##0;(#,##0)"), "(1,235)");
    }

    #[test]
    fn dates_render_the_way_excel_shows_them() {
        for (value, format, shown) in [
            (45000.0, "yyyy-mm-dd", "2023-03-15"),
            (45000.0, "mm-dd-yy", "03-15-23"),
            (45000.0, "yyyy\"Y\"m\"M\"d\"D\"", "2023Y3M15D"),
            (45000.5, "m/d/yy h:mm", "3/15/23 12:00"),
            (45000.75, "h:mm:ss", "18:00:00"),
            (0.5, "h:mm", "12:00"),
            (45000.0, "dddd", "Wednesday"),
        ] {
            assert_eq!(format_number(value, format), shown, "{value} as {format}");
        }
    }
}

#[cfg(test)]
mod format_codes_from_the_wild {
    use super::format_number;

    /// Both codes are lifted from a government workbook, after the XML
    /// unescaping that turns `&quot;` back into a quotation mark.
    #[test]
    fn a_negative_section_can_carry_its_own_marker() {
        assert_eq!(format_number(1.5, "0.0;\"▲\"0.0"), "1.5");
        assert_eq!(format_number(-1.5, "0.0;\"▲\"0.0"), "▲1.5");
        assert_eq!(format_number(0.0, "0.0;\"▲\"0.0"), "0.0");
    }

    #[test]
    fn a_backslash_makes_the_next_character_literal() {
        assert_eq!(format_number(1.5, r"\(0.0\)"), "(1.5)");
        assert_eq!(format_number(-1.5, r#"\(0.0\);"（▲"0.0\)"#), "（▲1.5)");
    }
}
