// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! English hyphenation (Liang's algorithm), for `<w:autoHyphenation/>`.
//!
//! DERIVED from Word (tools/metrics/_pb_hyphen_gen.py, 10 words × 12 widths,
//! Times New Roman 12 on Letter, break positions read off the Word PDF). Word's
//! rule has two stages and all 120 arms agree with it:
//!
//! 1. If the line's natural right gap — what is left after the last WHOLE word
//!    that fits — is within the hyphenation zone (`w:hyphenationZone`, default
//!    0.25in = 360tw = 18pt), Word does NOT hyphenate. `according` at a 15.3pt
//!    gap stays whole even though `ac-` (13pt) would have fitted.
//! 2. Otherwise Word takes the LONGEST legal prefix whose width plus the hyphen
//!    fits the remaining space, and wraps the whole word when none does:
//!    `countries` at a 27.3pt gap stays whole because `coun-` measures 28pt;
//!    `beautiful` at 21.4pt stays whole because `beau-` measures 24pt;
//!    `resource` takes its only break `re-` even though 26.2pt still remains;
//!    `through` (no legal break) stays whole at gaps up to 34.8pt.
//!
//! Word's break positions agree with the TeX en-US patterns on 8 of the 10
//! probe words (ac-cord-ing, coun-tries, au-thor-i-tar-i-an, pre-dis-posed,
//! in-for-ma-tion, beau-ti-ful, re-source, and `through` with no break at all).
//! They do NOT agree on two, where a suppressing pattern in this table forbids
//! a break Word takes:
//!   `de-mocracy`   — `.de1mo` allows it, `.de4mocr`/`4mocr` override to 4
//!   `hyphena-tion` — `he1na4` forbids it
//! so Word's own hyphenator is not this table. The divergence is recorded here
//! rather than patched away: a per-word carve-out would be exactly the
//! exception-stacking the derivation rules forbid, and the honest fix is a
//! table measured against Word. Until then the table is right where it is
//! right, and Oxi hyphenating nothing at all is wrong everywhere.
//!
//! ## Pattern data
//!
//! `data/hyph_en_us.txt` is the pattern list from `hyph_en_US.dic` (Hunspell /
//! LibreOffice `dict-en`), whose README states:
//!
//! > License: BSD-style. Unlimited copying, redistribution and modification of
//! > this file is permitted with this copyright and license information.
//! > Conversion and modifications by László Németh (nemeth at OOo).
//! > Based on the plain TeX hyphenation table
//! > (`macros/plain/base/hyphen.tex`) and the TugBoat hyphenation exceptions log
//! > (`info/digests/tugboat/tb0hyf.tex`), processed by `hyphenex.sh`.
//!
//! and the upstream notices it carries forward:
//!
//! > hyphen.tex — The Plain TeX hyphenation tables [NOT TO BE CHANGED IN ANY
//! > WAY!] Unlimited copying and redistribution of this file are permitted as
//! > long as this file is not modified.
//! >
//! > hyphenex output — Hyphenation exceptions for US English … Copyright 2007
//! > TeX Users Group. You may freely use, modify and/or distribute this file.
//!
//! BSD-style terms are compatible with this crate's MPL-2.0 (the repository
//! rule is MIT / Apache-2.0 / BSD for third-party data and code).

use std::collections::HashMap;
use std::sync::OnceLock;

/// Word's own minimums, taken from the same dictionary file
/// (`LEFTHYPHENMIN 2` / `RIGHTHYPHENMIN 3`): at least 2 characters stay on the
/// first line and at least 3 move to the next.
const LEFT_MIN: usize = 2;
const RIGHT_MIN: usize = 3;

const PATTERNS: &str = include_str!("data/hyph_en_us.txt");

struct Patterns {
    /// pattern letters (no digits, '.' kept as the word boundary marker) ->
    /// the digit vector that pattern contributes, one slot per gap.
    map: HashMap<String, Vec<u8>>,
    max_len: usize,
}

fn patterns() -> &'static Patterns {
    static P: OnceLock<Patterns> = OnceLock::new();
    P.get_or_init(|| {
        let mut map: HashMap<String, Vec<u8>> = HashMap::new();
        let mut max_len = 0usize;
        for line in PATTERNS.lines() {
            let line = line.trim();
            // `%` comments, blank lines, and the `word=with=breaks` exception
            // form (this dictionary has none, but the format allows it).
            if line.is_empty() || line.starts_with('%') || line.contains('=') {
                continue;
            }
            let mut letters = String::new();
            // values[i] applies to the gap BEFORE letters[i]; one extra slot
            // at the end for the gap after the last letter.
            let mut values: Vec<u8> = vec![0];
            for ch in line.chars() {
                if let Some(d) = ch.to_digit(10) {
                    let last = values.len() - 1;
                    values[last] = d as u8;
                } else {
                    letters.push(ch);
                    values.push(0);
                }
            }
            max_len = max_len.max(letters.chars().count());
            map.insert(letters, values);
        }
        Patterns { map, max_len }
    })
}

/// Byte offsets inside `word` after which a hyphen may be placed, ascending.
///
/// The word is matched lowercased and wrapped in the `.` boundary markers the
/// pattern file uses. Only pure-alphabetic words are hyphenated — a token
/// carrying digits or punctuation is left alone, which is also what Word does
/// with the corpus's `04_205_0104_6_1`-style codes.
pub fn break_offsets(word: &str) -> Vec<usize> {
    let chars: Vec<char> = word.chars().collect();
    if chars.len() < LEFT_MIN + RIGHT_MIN || !chars.iter().all(|c| c.is_alphabetic()) {
        return Vec::new();
    }
    let p = patterns();
    let lower: String = word.to_lowercase();
    let padded: Vec<char> = std::iter::once('.')
        .chain(lower.chars())
        .chain(std::iter::once('.'))
        .collect();
    // points[i] is the priority of the gap before padded[i].
    let mut points = vec![0u8; padded.len() + 1];
    for start in 0..padded.len() {
        let mut piece = String::new();
        for len in 1..=p.max_len.min(padded.len() - start) {
            piece.push(padded[start + len - 1]);
            if let Some(vals) = p.map.get(&piece) {
                for (k, v) in vals.iter().enumerate() {
                    let idx = start + k;
                    if idx < points.len() && *v > points[idx] {
                        points[idx] = *v;
                    }
                }
            }
        }
    }
    // points index -> position in the ORIGINAL word: padded[0] is '.', so the
    // gap before padded[i] sits after (i - 1) real characters.
    let mut out = Vec::new();
    let mut byte = 0usize;
    for (ci, ch) in chars.iter().enumerate() {
        byte += ch.len_utf8();
        let after = ci + 1; // characters kept on this line
        if after < LEFT_MIN || chars.len() - after < RIGHT_MIN {
            continue;
        }
        // odd value = a permitted break
        if points.get(after + 1).copied().unwrap_or(0) % 2 == 1 {
            out.push(byte);
        }
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    fn splits(word: &str) -> Vec<String> {
        break_offsets(word)
            .into_iter()
            .map(|b| word[..b].to_string())
            .collect()
    }

    #[test]
    fn matches_the_word_probe() {
        // Every prefix Word was OBSERVED to hyphenate at
        // (tools/metrics/_pb_hyphen_gen.py, Word PDF).
        for (word, seen) in [
            ("according", vec!["ac", "accord"]),
            ("countries", vec!["coun"]),
            ("predisposed", vec!["pre", "predis"]),
            ("information", vec!["in", "infor"]),
            ("beautiful", vec!["beau", "beauti"]),
            ("resource", vec!["re"]),
        ] {
            let got = splits(word);
            for s in seen {
                assert!(got.iter().any(|g| g == s), "{word}: {s:?} missing from {got:?}");
            }
        }
        // Word never hyphenated this one, at gaps up to 34.8pt.
        assert!(break_offsets("through").is_empty());
    }

    #[test]
    fn records_where_the_table_and_word_disagree() {
        // Word broke BOTH of these in the probe; this table forbids the first
        // break of each (see the module docs). Asserted so the divergence
        // shows up as a fact in the test suite instead of as a surprise in a
        // corpus gate — and so a future table swap has to confront it.
        assert_eq!(splits("democracy"), vec!["democ"]); // Word also takes "de-"
        assert_eq!(splits("hyphenation"), vec!["hy", "hyphen"]); // Word also takes "hyphena-"
    }

    #[test]
    fn respects_the_dictionary_minimums() {
        // LEFTHYPHENMIN 2 / RIGHTHYPHENMIN 3.
        for w in ["a", "an", "the", "item", "oxide"] {
            for b in break_offsets(w) {
                assert!(b >= LEFT_MIN && w.len() - b >= RIGHT_MIN, "{w} broke at {b}");
            }
        }
    }

    #[test]
    fn leaves_non_words_alone() {
        assert!(break_offsets("04_205_0104_6_1").is_empty());
        assert!(break_offsets("x1y2z3").is_empty());
    }
}
