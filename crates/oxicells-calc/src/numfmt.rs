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
    // A number under the text format `@` shows as General: measured, 12.5 in
    // a cell formatted `@` reads 12.5.
    if format.is_empty() || format.eq_ignore_ascii_case("general") || format == "@" {
        return general(value);
    }

    let sections: Vec<&str> = split_sections(format);
    // A section may say which numbers it is for: measured, under
    // `[<1000]0;#,##0,"K"` 1500 is `2K`. With conditions the first section
    // whose condition holds is used, and one without a condition takes the
    // rest; a number shown that way keeps its sign.
    let conditions: Vec<Option<(String, f64)>> = sections.iter().map(|one| section_condition(one)).collect();
    let (section, signed) = if conditions.iter().any(Option::is_some) {
        let numeric = &sections[..sections.len().min(3)];
        let chosen = numeric
            .iter()
            .zip(&conditions)
            .find(|(_, condition)| match condition {
                Some((op, bound)) => condition_holds(op, value, *bound),
                None => true,
            })
            .map(|(one, _)| *one);
        match chosen {
            Some(one) => (one, true),
            None => return "#".repeat(1),
        }
    } else if value < 0.0 && sections.len() > 1 {
        // The negative section carries its own sign, so the value loses it.
        (sections[1], false)
    } else if value == 0.0 && sections.len() > 2 {
        (sections[2], false)
    } else {
        (sections[0], true)
    };

    // A section with no place for a digit shows only its own words:
    // measured, `0;"neg";"zero"` shows -5 as `neg` and 0 as `zero`.
    if !has_digit_place(section) && !looks_like_a_date(section) && !section.eq_ignore_ascii_case("general") {
        let words = literal_text(section);
        // A number that came by a condition keeps its sign too: measured,
        // under `[>100]"big";"small"` -0.5 shows -small.
        let conditional = conditions.iter().any(Option::is_some);
        return if signed && value < 0.0 && !words.is_empty() && (sections.len() == 1 || conditional) {
            format!("-{words}")
        } else {
            words
        };
    }
    if section.trim().eq_ignore_ascii_case("general") {
        return general(if signed { value } else { value.abs() });
    }

    let magnitude = if signed { value } else { value.abs() };
    if looks_like_a_date(section) {
        return format_datetime(magnitude, section);
    }
    if let Some(shape) = fraction_shape(section) {
        return format_fraction(magnitude, &shape);
    }
    format_numeric(magnitude, section)
}

/// Text under a format: its fourth section, or a lone section holding `@`,
/// with the text where the `@` is. Measured: under `0;0;0;"txt:"@` "abc" is
/// `txt:abc`; a format with no text section shows text as it is.
pub fn format_text(text: &str, format: &str) -> String {
    let sections = split_sections(format);
    let section = match sections.len() {
        4.. => sections[3],
        1 if sections[0].contains('@') => sections[0],
        _ => return text.to_string(),
    };
    let mut out = String::new();
    let mut quoted = false;
    let mut characters = section.chars();
    while let Some(character) = characters.next() {
        match character {
            '"' => quoted = !quoted,
            _ if quoted => out.push(character),
            '\\' => out.extend(characters.next()),
            '_' => {
                characters.next();
                out.push(' ');
            }
            '*' => {
                characters.next();
            }
            '[' => {
                for held in characters.by_ref() {
                    if held == ']' {
                        break;
                    }
                }
            }
            '@' => out.push_str(text),
            other => out.push(other),
        }
    }
    out
}

/// Whether a section has anywhere for a digit to go.
fn has_digit_place(section: &str) -> bool {
    let mut quoted = false;
    let mut escaped = false;
    let mut bracket = false;
    for character in section.chars() {
        if escaped {
            escaped = false;
            continue;
        }
        match character {
            '"' => quoted = !quoted,
            _ if quoted => {}
            '\\' | '_' | '*' => escaped = true,
            '[' => bracket = true,
            ']' => bracket = false,
            _ if bracket => {}
            '0' | '#' | '?' => return true,
            _ => {}
        }
    }
    false
}

/// A section's condition, `[<1000]` or `[>=5]`, as the operator and bound.
fn section_condition(section: &str) -> Option<(String, f64)> {
    let mut rest = section;
    while let Some(open) = rest.find('[') {
        let close = rest[open..].find(']')? + open;
        let inside = &rest[open + 1..close];
        for op in ["<=", ">=", "<>", "<", ">", "="] {
            if let Some(bound) = inside.strip_prefix(op) {
                if let Ok(bound) = bound.trim().parse::<f64>() {
                    return Some((op.to_string(), bound));
                }
            }
        }
        rest = &rest[close + 1..];
    }
    None
}

fn condition_holds(op: &str, value: f64, bound: f64) -> bool {
    match op {
        "<" => value < bound,
        "<=" => value <= bound,
        ">" => value > bound,
        ">=" => value >= bound,
        "=" => value == bound,
        _ => value != bound,
    }
}

/// The parts of a fraction format: `# ?/?` is a whole part, the text between,
/// a numerator one place wide and a denominator one place wide; `?/8` has no
/// whole part and a fixed denominator.
struct FractionShape {
    prefix: String,
    /// The whole part's placeholders, empty for an improper fraction.
    whole: String,
    between: String,
    numerator_width: usize,
    denominator: Denominator,
    suffix: String,
}

enum Denominator {
    /// Up to this many digits: `?` is 9, `??` is 99.
    Free(usize),
    Fixed(u64),
}

/// A format is a fraction format when a slash outside quotes has a digit
/// placeholder on its left and a placeholder or a number on its right.
fn fraction_shape(format: &str) -> Option<FractionShape> {
    let body: Vec<char> = format.chars().collect();
    let placeholder = |held: char| matches!(held, '0' | '#' | '?');
    let mut quoted = false;
    let mut slash = None;
    let mut at = 0;
    while at < body.len() {
        match body[at] {
            '"' => quoted = !quoted,
            _ if quoted => {}
            '_' | '\\' | '*' => at += 1,
            '[' => {
                while at < body.len() && body[at] != ']' {
                    at += 1;
                }
            }
            '/' => {
                slash = Some(at);
                break;
            }
            _ => {}
        }
        at += 1;
    }
    let slash = slash?;
    // The numerator: the run of placeholders ending at the slash.
    let mut start = slash;
    while start > 0 && placeholder(body[start - 1]) {
        start -= 1;
    }
    let numerator_width = slash - start;
    if numerator_width == 0 {
        return None;
    }
    // The text between the whole part and the numerator, then the whole part.
    let mut cut = start;
    while cut > 0 && !placeholder(body[cut - 1]) {
        cut -= 1;
    }
    let between: String = body[cut..start].iter().collect();
    let mut whole_start = cut;
    while whole_start > 0 && placeholder(body[whole_start - 1]) {
        whole_start -= 1;
    }
    let whole: String = body[whole_start..cut].iter().collect();
    let prefix: String = body[..whole_start].iter().collect();
    // The denominator: placeholders, or the digits of a fixed one.
    let mut end = slash + 1;
    let denominator = if end < body.len() && body[end].is_ascii_digit() {
        let mut digits = String::new();
        while end < body.len() && body[end].is_ascii_digit() {
            digits.push(body[end]);
            end += 1;
        }
        Denominator::Fixed(digits.parse().ok().filter(|held| *held > 0)?)
    } else {
        while end < body.len() && placeholder(body[end]) {
            end += 1;
        }
        if end == slash + 1 {
            return None;
        }
        Denominator::Free(end - slash - 1)
    };
    let suffix: String = body[end..].iter().collect();
    Some(FractionShape {
        prefix: literal_text(&prefix),
        whole,
        between: literal_text(&between),
        numerator_width,
        denominator,
        suffix: literal_text(&suffix),
    })
}

/// The text a stretch of format shows for itself: quotes gone, `\x` and `_x`
/// read as one character, brackets dropped.
fn literal_text(format: &str) -> String {
    let mut text = String::new();
    let mut quoted = false;
    let mut characters = format.chars();
    while let Some(character) = characters.next() {
        match character {
            '"' => quoted = !quoted,
            _ if quoted => text.push(character),
            '\\' => text.extend(characters.next()),
            '_' => {
                characters.next();
                text.push(' ');
            }
            '*' => {
                characters.next();
            }
            '[' => {
                for held in characters.by_ref() {
                    if held == ']' {
                        break;
                    }
                }
            }
            other => text.push(other),
        }
    }
    text
}

/// A number under a fraction format, as Excel shows it.
///
/// Measured in `Range.Text`: under `# ?/?` 1.5 is `1 1/2`, 0.75 is ` 3/4` --
/// the `#` shows nothing for the whole part and the space between stays --
/// 3.14159265 is `3 1/7`, and 2, 0.96 and 0.05 are `2    `, `1    ` and
/// `0    `: a fraction that comes to nothing leaves its width in spaces, and a
/// whole part that would then be all there is shows its zero. `# ??/??` allows
/// a denominator to 99 and pads the numerator on the left and the denominator
/// on the right: 1.5 is `1  1/2 ` and 12.3456 is `12 28/81`. `?/?` is
/// improper, 1.5 being `3/2` and 12.3456 `37/3`. A fixed denominator is not
/// reduced: `# ?/8` shows 1.5 as `1 4/8` and 0.75 as ` 6/8`, and a numerator
/// wider than its place is shown whole, ` 12/16`. Which fraction stands for a
/// value is `nearest_fraction`'s business, and it is not always the nearest.
fn format_fraction(value: f64, shape: &FractionShape) -> String {
    let negative = value < 0.0;
    let magnitude = value.abs();
    let mixed = !shape.whole.is_empty();
    let (mut whole, part) = if mixed {
        (magnitude.trunc(), magnitude - magnitude.trunc())
    } else {
        (0.0, magnitude)
    };
    let (mut numerator, denominator) = match shape.denominator {
        Denominator::Fixed(held) => ((part * held as f64 + 0.5).floor() as u64, held),
        Denominator::Free(width) => {
            let most = 10u64.pow(width as u32) - 1;
            nearest_fraction(part, most)
        }
    };
    if mixed && numerator == denominator {
        whole += 1.0;
        numerator = 0;
    }
    let blank = mixed && numerator == 0;

    let mut text = String::new();
    text.push_str(&shape.prefix);
    if negative {
        text.push('-');
    }
    if mixed {
        let whole = whole as u64;
        let zero_places = shape.whole.chars().filter(|held| *held == '0').count();
        let shown = if whole == 0 && zero_places == 0 {
            // `#` shows no zero -- unless the zero is all there is to show.
            if blank { "0".to_string() } else { String::new() }
        } else {
            format!("{whole:0>zero_places$}")
        };
        text.push_str(&shown);
        text.push_str(&shape.between);
    }
    let denominator_width = match shape.denominator {
        Denominator::Fixed(held) => held.to_string().len(),
        Denominator::Free(width) => width,
    };
    if blank {
        for _ in 0..shape.numerator_width + 1 + denominator_width {
            text.push(' ');
        }
    } else {
        let top = numerator.to_string();
        let bottom = denominator.to_string();
        text.push_str(&format!("{top:>width$}", width = shape.numerator_width));
        text.push('/');
        text.push_str(&format!("{bottom:<width$}", width = denominator_width));
    }
    text.push_str(&shape.suffix);
    text
}

/// The fraction Excel picks for a value when the denominator may run to
/// `most`: the last CONVERGENT of the value's continued fraction whose
/// denominator fits, not the nearest fraction that fits.
///
/// The two differ, and Excel is measured on the side of the convergents:
/// 1.0625 under `# ?/?` shows `1    `, though 1/9 is nearer to 1/16 than 0
/// is; and 12.3456 under `# ??/??` shows 28/81, which is the convergent
/// before 47/136, where the nearest fraction under 100 would be found by a
/// search that Excel does not make.
fn nearest_fraction(value: f64, most: u64) -> (u64, u64) {
    let (mut previous, mut current) = ((1u64, 0u64), (value.floor() as u64, 1u64));
    let mut rest = value - value.floor();
    for _ in 0..64 {
        if rest < 1e-9 {
            break;
        }
        let inverted = 1.0 / rest;
        let term = inverted.floor();
        rest = inverted - term;
        let term = term as u64;
        let Some(next_denominator) = term
            .checked_mul(current.1)
            .and_then(|held| held.checked_add(previous.1))
        else {
            break;
        };
        if next_denominator > most {
            break;
        }
        let next_numerator = term * current.0 + previous.0;
        previous = current;
        current = (next_numerator, next_denominator);
    }
    current
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
pub fn looks_like_a_date(format: &str) -> bool {
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
            // The Japanese era parts: `g` the era's name, `e` its year, `r`
            // its year in two digits -- and a run of three or more `a`s is the
            // weekday. Measured: `ggge年` shows 令和6年, `r` shows 06, `aaa`
            // shows 金. An `e` before a sign is scientific notation, not an
            // era; `General` is not a format made of a g and an e.
            'g' | 'r' => marked_month = true,
            'e' if !matches!(characters.peek(), Some('+') | Some('-')) => marked_month = true,
            'a' => {
                let mut run = 1;
                while characters.peek().is_some_and(|held| held.eq_ignore_ascii_case(&'a')) {
                    characters.next();
                    run += 1;
                }
                if run >= 3 {
                    return true;
                }
            }
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
    marked_month && !format.trim().eq_ignore_ascii_case("general")
}

/// What `General` shows: the shortest text that reads back as the same number.
fn general(value: f64) -> String {
    // General fits a number in eleven characters, the sign aside: the whole
    // digits, a point and as many decimals as are left, trailing zeros off;
    // what cannot be shown that way -- twelve whole digits or more, or a
    // fraction that rounds to nothing -- goes to exponent form with five
    // decimals at most. Measured with TEXT(x,"General"): 12345678901.5 is
    // 12345678902, 0.123456789012 0.123456789, 1234567890.12 1234567890,
    // 0.0000123 0.0000123, 1E-10 1E-10, 1E+15 1E+15.
    if value == 0.0 || !value.is_finite() {
        return "0".to_string();
    }
    let magnitude = value.abs();
    let sign = if value < 0.0 { "-" } else { "" };
    let whole_digits = if magnitude < 1.0 { 1 } else { magnitude.log10().floor() as i32 + 1 };
    if whole_digits <= 11 {
        let decimals = (11 - whole_digits - 1).max(0);
        let rounded = round_half_away(magnitude, decimals);
        let rounded_digits = if rounded < 1.0 { 1 } else { rounded.log10().floor() as i32 + 1 };
        if rounded != 0.0 && rounded_digits <= 11 {
            let written = format!("{rounded:.*}", decimals as usize);
            let written = if written.contains('.') {
                written.trim_end_matches('0').trim_end_matches('.').to_string()
            } else {
                written
            };
            return format!("{sign}{written}");
        }
    }
    let exponent = magnitude.log10().floor() as i32;
    let mut mantissa = round_half_away(magnitude / 10f64.powi(exponent), 5);
    let mut exponent = exponent;
    if mantissa >= 10.0 {
        mantissa /= 10.0;
        exponent += 1;
    }
    let written = format!("{mantissa:.5}");
    let written = written.trim_end_matches('0').trim_end_matches('.');
    let mark = if exponent < 0 { '-' } else { '+' };
    format!("{sign}{written}E{mark}{:02}", exponent.abs())
}

/// A digit place in a format: `0` shows a digit or a zero, `#` a digit or
/// nothing, `?` a digit or a space.
#[derive(Clone, Copy, PartialEq)]
enum Place {
    Zero,
    Hash,
    Query,
}

fn place_of(character: char) -> Option<Place> {
    match character {
        '0' => Some(Place::Zero),
        '#' => Some(Place::Hash),
        '?' => Some(Place::Query),
        _ => None,
    }
}

/// What stands in a place that has no digit of its own.
fn empty_place(place: Place) -> &'static str {
    match place {
        Place::Zero => "0",
        Place::Hash => "",
        Place::Query => " ",
    }
}

/// Where a character of the format sits: before the point, after it, or in
/// the exponent.
#[derive(Clone, Copy, PartialEq)]
enum Part {
    Whole,
    Fraction,
    Exponent,
}

/// A number under a numeric format section, as Excel shows it. Measured in
/// `Range.Text`: `0.##` shows 12 as `12.`, `??.??` 3.14159 as ` 3.14`,
/// `0.0E+0` 0.000123 as `1.2E-4`, `#,##0,,` 1234567 as `1`; the places are
/// filled from the right, what the number has beyond them going to the first
/// one; `#` leaves nothing and `?` a space where there is no digit, and the
/// fraction's trailing zeros go the same way.
fn format_numeric(value: f64, format: &str) -> String {
    let body: Vec<char> = format.chars().collect();

    // The shape: which places there are in each part, and what else the
    // format asks of the number.
    let mut whole_places: Vec<Place> = Vec::new();
    let mut fraction_places: Vec<Place> = Vec::new();
    let mut exponent_places: Vec<Place> = Vec::new();
    let mut percent = 0i32;
    let mut grouped = false;
    let mut scale = 0i32;
    let mut scientific: Option<bool> = None;
    let mut part = Part::Whole;
    let mut quoted = false;
    let mut at = 0;
    while at < body.len() {
        let character = body[at];
        if quoted {
            if character == '"' {
                quoted = false;
            }
            at += 1;
            continue;
        }
        match character {
            '"' => quoted = true,
            '\\' | '_' | '*' => at += 1,
            '[' => {
                while at < body.len() && body[at] != ']' {
                    at += 1;
                }
            }
            '.' if part == Part::Whole => part = Part::Fraction,
            '%' => percent += 1,
            'E' | 'e' if scientific.is_none() && matches!(body.get(at + 1), Some('+' | '-')) => {
                scientific = Some(body[at + 1] == '+');
                part = Part::Exponent;
                at += 1;
            }
            // After the fraction's last place a comma still divides by a
            // thousand: measured, TEXT(123456789,"0.0,,""M""") is 123.5M.
            ',' if part == Part::Fraction => {
                if !fraction_places.is_empty()
                    && !body[at + 1..].iter().take_while(|held| !matches!(held, ';')).any(|held| place_of(*held).is_some())
                {
                    scale += 1;
                }
            }
            ',' if part == Part::Whole => {
                // Among the places it groups; after the last place before
                // the point it divides by a thousand.
                let more_places = body[at + 1..]
                    .iter()
                    .take_while(|held| !matches!(held, '.' | 'E' | 'e' | ';'))
                    .any(|held| place_of(*held).is_some());
                if more_places && !whole_places.is_empty() {
                    grouped = true;
                } else if !whole_places.is_empty() {
                    scale += 1;
                }
            }
            other => {
                if let Some(place) = place_of(other) {
                    match part {
                        Part::Whole => whole_places.push(place),
                        Part::Fraction => fraction_places.push(place),
                        Part::Exponent => exponent_places.push(place),
                    }
                }
            }
        }
        at += 1;
    }

    let negative = value < 0.0;
    let mut number = value.abs() * 100f64.powi(percent) / 1000f64.powi(scale);

    // Scientific: the mantissa has as many whole digits as the format has
    // places for, and the exponent comes out of that.
    let mut exponent = 0i32;
    if scientific.is_some() && number != 0.0 {
        let places = whole_places.len().max(1) as i32;
        let magnitude = number.log10().floor() as i32;
        exponent = if places > 1 && whole_places.first() == Some(&Place::Hash) {
            magnitude.div_euclid(places) * places
        } else {
            magnitude - (places - 1)
        };
        number /= 10f64.powi(exponent);
        // Rounding may carry the mantissa over into another digit.
        let decimals = fraction_places.len() as i32;
        let settled = round_half_away(number, decimals);
        if settled >= 10f64.powi(places) {
            number = settled / 10.0;
            exponent += 1;
        }
    }

    let settled = round_half_away(number, fraction_places.len() as i32);
    let rendered = format!("{:.*}", fraction_places.len(), settled);
    let (whole_digits, fraction_digits) = match rendered.split_once('.') {
        Some((whole, fraction)) => (whole.to_string(), fraction.to_string()),
        None => (rendered.clone(), String::new()),
    };
    let whole_digits = if whole_digits == "0" { String::new() } else { whole_digits };

    // The whole part, place by place from the right; what does not fit goes
    // to the first place.
    let count = whole_places.len();
    let digits: Vec<char> = whole_digits.chars().collect();
    let mut whole_slots: Vec<String> = Vec::with_capacity(count);
    for (index, place) in whole_places.iter().enumerate() {
        let from_right = count - 1 - index;
        let mut slot = String::new();
        if index == 0 && digits.len() > count {
            slot.extend(&digits[..digits.len() - count]);
        }
        if from_right < digits.len() {
            slot.push(digits[digits.len() - 1 - from_right]);
        } else {
            slot.push_str(empty_place(*place));
        }
        whole_slots.push(slot);
    }
    if grouped && count > 0 {
        // Grouping runs over the digits the whole part shows, however they
        // are spread; they are gathered into the first place.
        let shown: String = whole_slots.concat();
        let lead: String = shown.chars().take_while(|held| !held.is_ascii_digit()).collect();
        let figures: String = shown.chars().filter(char::is_ascii_digit).collect();
        whole_slots = vec![String::new(); count];
        whole_slots[0] = format!("{lead}{}", group_thousands(&figures));
    }

    // The fraction's trailing zeros: gone under `#`, a space under `?`.
    let mut fraction_slots: Vec<String> = fraction_digits.chars().map(String::from).collect();
    for index in (0..fraction_slots.len()).rev() {
        if fraction_slots[index] != "0" || fraction_places[index] == Place::Zero {
            break;
        }
        fraction_slots[index] = empty_place(fraction_places[index]).to_string();
    }

    // The exponent, at least as many digits as it has `0`s.
    let exponent_digits = {
        let zeros = exponent_places.iter().filter(|place| **place == Place::Zero).count();
        format!("{:0>width$}", exponent.abs(), width = zeros.max(1))
    };

    // Now the format again, left to right, putting it all in its place.
    let mut out = String::new();
    // ...unless nothing is left of it once rounded: measured,
    // TEXT(-0.5,"#,##0,") is 0.
    let nothing_left = settled == 0.0 && scientific.is_none();
    if negative && !nothing_left {
        out.push('-');
    }
    let (mut whole_at, mut fraction_at) = (0usize, 0usize);
    let mut part = Part::Whole;
    let mut quoted = false;
    let mut at = 0;
    while at < body.len() {
        let character = body[at];
        if quoted {
            if character == '"' {
                quoted = false;
            } else {
                out.push(character);
            }
            at += 1;
            continue;
        }
        match character {
            '"' => quoted = true,
            '\\' => {
                if let Some(next) = body.get(at + 1) {
                    out.push(*next);
                }
                at += 1;
            }
            // `_x` keeps the width of x: Excel's own `Range.Text` gives a
            // space. `*x` fills what the cell has left, which depends on its
            // width, so nothing is put here.
            '_' => {
                out.push(' ');
                at += 1;
            }
            '*' => at += 1,
            '[' => {
                // `[$€-407]` shows its currency; a colour, a condition or a
                // bare locale shows nothing.
                let start = at;
                while at < body.len() && body[at] != ']' {
                    at += 1;
                }
                let inside: String = body[start + 1..at.min(body.len())].iter().collect();
                if let Some(currency) = inside.strip_prefix('$') {
                    out.push_str(currency.split('-').next().unwrap_or(""));
                }
            }
            '.' if part == Part::Whole => {
                part = Part::Fraction;
                out.push('.');
            }
            ',' if part == Part::Whole || part == Part::Fraction => {}
            'E' | 'e' if part != Part::Exponent && scientific.is_some() && matches!(body.get(at + 1), Some('+' | '-')) => {
                part = Part::Exponent;
                out.push(character);
                if exponent < 0 {
                    out.push('-');
                } else if scientific == Some(true) {
                    out.push('+');
                }
                out.push_str(&exponent_digits);
                at += 1;
            }
            other => match (place_of(other), part) {
                (Some(_), Part::Whole) => {
                    if let Some(slot) = whole_slots.get(whole_at) {
                        out.push_str(slot);
                    }
                    whole_at += 1;
                }
                (Some(_), Part::Fraction) => {
                    if let Some(slot) = fraction_slots.get(fraction_at) {
                        out.push_str(slot);
                    }
                    fraction_at += 1;
                }
                (Some(_), Part::Exponent) => {}
                (None, _) => out.push(other),
            },
        }
        at += 1;
    }
    out
}

/// Half away from zero, as Excel rounds for display; Rust's own formatting
/// sends a half to the even neighbour.
fn round_half_away(number: f64, decimals: i32) -> f64 {
    let scale = 10f64.powi(decimals);
    (number * scale).round() / scale
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


/// How many decimals the seconds of a date format carry: the zeros after
/// `s.` or `ss.`, three at most.
fn second_places(format: &str) -> u32 {
    let lower = format.to_ascii_lowercase();
    let body: Vec<char> = lower.chars().collect();
    let mut quoted = false;
    for at in 0..body.len() {
        match body[at] {
            '"' => quoted = !quoted,
            's' if !quoted && body.get(at + 1) == Some(&'.') => {
                let zeros = body[at + 2..].iter().take_while(|held| **held == '0').count();
                return zeros.min(3) as u32;
            }
            _ => {}
        }
    }
    0
}

fn format_datetime(serial: f64, format: &str) -> String {
    let whole = serial.trunc() as i64;
    let Ok(date) = date_from_serial(whole) else {
        return general(serial);
    };
    let (year, month, day) = (date.year, date.month, date.day);
    // The fraction of a day is the time, rounded to the nearest second the way
    // Excel shows it.
    //
    // Unless the seconds carry decimals: `mm:ss.0` rounds to the tenth, and
    // measured, 0.0123 of a day is 17:42.7 there, where `mm:ss` says 17:43.
    let places = second_places(format);
    let scale = 10_i64.pow(places);
    let ticks = ((serial - serial.trunc()) * 86_400.0 * scale as f64).round() as i64;
    let seconds_of_day = ticks / scale;
    let fraction = ticks % scale;
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
    // The Japanese weekday, which `aaa` and `aaaa` always show, and which
    // `ddd` and `dddd` show under the Japanese locale tag: measured,
    // `[$-411]dddd` is 金曜日 where `dddd` is Friday.
    const JA_DAYS: [&str; 7] = ["日", "月", "火", "水", "木", "金", "土"];
    let japanese = format.to_ascii_lowercase().contains("[$-411]");
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
                3 if japanese => rendered.push_str(JA_DAYS[weekday]),
                3 => rendered.push_str(&DAYS[weekday][..3]),
                _ if japanese => {
                    rendered.push_str(JA_DAYS[weekday]);
                    rendered.push_str("曜日");
                }
                _ => rendered.push_str(DAYS[weekday]),
            },
            'a' if run >= 3 => {
                rendered.push_str(JA_DAYS[weekday]);
                if run >= 4 {
                    rendered.push_str("曜日");
                }
            }
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
                let zeros = body[at + run..]
                    .iter()
                    .skip(1)
                    .take_while(|held| **held == '0')
                    .count();
                if places > 0 && body.get(at + run) == Some(&'.') && zeros > 0 {
                    rendered.push('.');
                    rendered.push_str(&format!("{fraction:0width$}", width = places as usize));
                    at += run + 1 + zeros;
                    continue;
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
            'r' => rendered.push_str(&format!("{era_year:02}")),
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
    use super::{format_number, format_text};

    /// Places, sections and text, each as Excel 16 shows it in `Range.Text`.
    #[test]
    fn places_sections_and_text_as_excel_shows_them() {
        assert_eq!(format_number(12.0, "0.##"), "12.");
        assert_eq!(format_number(3.14159, "??.??"), " 3.14");
        assert_eq!(format_number(0.000123, "0.0E+0"), "1.2E-4");
        assert_eq!(format_number(12345.678, "0.00E+00"), "1.23E+04");
        assert_eq!(format_number(1234567.0, "#,##0,,"), "1");
        assert_eq!(format_number(1500.0, "[<1000]0;#,##0,\"K\""), "2K");
        assert_eq!(format_number(-5.0, "0;\"neg\";\"zero\""), "neg");
        assert_eq!(format_number(0.0, "0;\"neg\";\"zero\""), "zero");
        assert_eq!(format_number(1.23456789, "0.####"), "1.2346");
        assert_eq!(format_text("abc", "0;0;0;\"txt:\"@"), "txt:abc");
        assert_eq!(format_text("abc", "0.00"), "abc");
    }

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
        // Measured in Excel on 5 January 2024 (a Friday): the era parts and
        // the Japanese weekday stand as date formats on their own, without a
        // locale tag or a y/m/d beside them.
        assert_eq!(format_number(45296.0, "aaa"), "金");
        assert_eq!(format_number(45296.0, "aaaa"), "金曜日");
        assert_eq!(format_number(45296.0, "dddd"), "Friday");
        assert_eq!(format_number(45296.0, "[$-411]dddd"), "金曜日");
        assert_eq!(format_number(45296.0, "[$-411]ddd"), "金");
        assert_eq!(format_number(45296.0, "ggge\"年\""), "令和6年");
        assert_eq!(format_number(45296.0, "ge.m.d"), "R6.1.5");
        assert_eq!(format_number(45296.0, "r"), "06");
        assert_eq!(format_number(45296.0, "e\"年\"m\"月\""), "6年1月");
        // And what only looks like one of them is not.
        assert_eq!(format_number(12345.0, "0.00E+00"), "1.23E+04");
        assert_eq!(format_number(12345.0, "General"), "12345");
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
mod fractions {
    use super::format_number;

    /// Every expectation is what Excel 16 put in `Range.Text` for that value
    /// under that format, in a column wide enough to show it all.
    #[test]
    fn fractions_are_shown_as_excel_shows_them() {
        // The last is pi to eight places, which is what was typed into Excel
        // and not the constant.
        #[allow(clippy::approx_constant)]
        let values = [
            1.5, 0.75, 2.0, 0.333333, 1.0625, -1.5, 0.05, 0.96, 12.3456, 0.0, 0.5, 100.125,
            3.14159265,
        ];
        let table: [(&str, [&str; 13]); 9] = [
            (
                "# ?/?",
                [
                    "1 1/2", " 3/4", "2    ", " 1/3", "1    ", "-1 1/2", "0    ", "1    ",
                    "12 1/3", "0    ", " 1/2", "100 1/8", "3 1/7",
                ],
            ),
            (
                "# ??/??",
                [
                    "1  1/2 ", "  3/4 ", "2      ", "  1/3 ", "1  1/16", "-1  1/2 ", "  1/20",
                    " 24/25", "12 28/81", "0      ", "  1/2 ", "100  1/8 ", "3  1/7 ",
                ],
            ),
            (
                "?/?",
                [
                    "3/2", "3/4", "2/1", "1/3", "1/1", "-3/2", "0/1", "1/1", "37/3", "0/1", "1/2",
                    "801/8", "22/7",
                ],
            ),
            (
                "# ?/8",
                [
                    "1 4/8", " 6/8", "2    ", " 3/8", "1 1/8", "-1 4/8", "0    ", "1    ",
                    "12 3/8", "0    ", " 4/8", "100 1/8", "3 1/8",
                ],
            ),
            (
                "# ?/16",
                [
                    "1 8/16", " 12/16", "2     ", " 5/16", "1 1/16", "-1 8/16", " 1/16",
                    " 15/16", "12 6/16", "0     ", " 8/16", "100 2/16", "3 2/16",
                ],
            ),
            (
                "# ???/???",
                [
                    "1   1/2  ", "   3/4  ", "2        ", "   1/3  ", "1   1/16 ", "-1   1/2  ",
                    "   1/20 ", "  24/25 ", "12 216/625", "0        ", "   1/2  ", "100   1/8  ",
                    "3  16/113",
                ],
            ),
            (
                "0 ?/?",
                [
                    "1 1/2", "0 3/4", "2    ", "0 1/3", "1    ", "-1 1/2", "0    ", "1    ",
                    "12 1/3", "0    ", "0 1/2", "100 1/8", "3 1/7",
                ],
            ),
            (
                "# ?/2",
                [
                    "1 1/2", "1    ", "2    ", " 1/2", "1    ", "-1 1/2", "0    ", "1    ",
                    "12 1/2", "0    ", " 1/2", "100    ", "3    ",
                ],
            ),
            (
                "??/??",
                [
                    " 3/2 ", " 3/4 ", " 2/1 ", " 1/3 ", "17/16", "- 3/2 ", " 1/20", "24/25",
                    "1000/81", " 0/1 ", " 1/2 ", "801/8 ", "22/7 ",
                ],
            ),
        ];
        for (format, shown) in table {
            for (value, expected) in values.iter().zip(shown) {
                assert_eq!(format_number(*value, format), expected, "{value} as {format}");
            }
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

/// How many percent signs each section of a format holds, quoted text,
/// escaped characters and bracketed parts aside.
pub fn sections_with_percents(format: &str) -> impl Iterator<Item = usize> {
    let mut counts = vec![0usize];
    let mut chars = format.chars();
    while let Some(character) = chars.next() {
        match character {
            '"' => {
                for inner in chars.by_ref() {
                    if inner == '"' {
                        break;
                    }
                }
            }
            '\\' => {
                chars.next();
            }
            '[' => {
                for inner in chars.by_ref() {
                    if inner == ']' {
                        break;
                    }
                }
            }
            ';' => counts.push(0),
            '%' => *counts.last_mut().expect("one section at least") += 1,
            _ => {}
        }
    }
    counts.into_iter()
}
