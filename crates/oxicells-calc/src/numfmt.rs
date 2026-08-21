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
    let mut characters = format.chars().peekable();
    while let Some(character) = characters.next() {
        match character {
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
                for held in characters.by_ref() {
                    if held == ']' {
                        break;
                    }
                }
            }
            'y' | 'd' | 'h' | 's' => return true,
            _ => {}
        }
    }
    false
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

    // Excel's twelve-hour clock only appears with AM/PM, which this does not
    // model; every hour here is on the twenty-four hour clock.
    // weekday_sunday_one counts Sunday as one; these tables start at zero.
    let weekday = (weekday_sunday_one(whole) - 1).clamp(0, 6) as usize;
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
