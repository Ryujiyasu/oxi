// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! OMML character substitution rules (COM-verified 2026-04-18).
//!
//! Word substitutes Latin/Greek letters and a few operators to their
//! Mathematical-Italic counterparts when rendering inside `<m:r>` runs.
//!
//! All 166 rules COM-measured; see Phase 1 memory:
//! - `project_omml_italic_math_table.md`  — Latin (62 chars, h special)
//! - `project_omml_greek_table.md`         — Greek (54 chars, linear)
//! - `project_omml_operators_table.md`     — Operators (3 of 50 substituted)
//!
//! The substitution is applied BEFORE glyph lookup in Cambria Math.

/// Substitute a single character per Word's OMML rendering rules.
///
/// Returns the Math-Italic variant where applicable, or the original char
/// if no substitution applies (digits, operators, relations, etc.).
///
/// # Examples
///
/// ```
/// use oxidocs_core::font::math_substitute;
///
/// assert_eq!(math_substitute('A'), '𝐴');  // U+1D434
/// assert_eq!(math_substitute('x'), '𝑥');  // U+1D465
/// assert_eq!(math_substitute('h'), 'ℎ');  // U+210E (special — Planck constant)
/// assert_eq!(math_substitute('α'), '𝛼');  // U+1D6FC
/// assert_eq!(math_substitute('Ω'), '𝛺');  // U+1D6FA
/// assert_eq!(math_substitute('-'), '−');  // U+2212 (ASCII minus → math minus)
/// assert_eq!(math_substitute('∂'), '𝜕');  // U+1D715 (italic partial)
/// assert_eq!(math_substitute('∇'), '𝛻');  // U+1D6FB (italic nabla)
/// assert_eq!(math_substitute('0'), '0');  // digits unchanged
/// assert_eq!(math_substitute('∑'), '∑');  // operators unchanged
/// assert_eq!(math_substitute('∞'), '∞');  // infinity unchanged
/// ```
pub fn math_substitute(c: char) -> char {
    let cp = c as u32;
    match cp {
        // Latin uppercase A-Z → U+1D434..U+1D44D (Mathematical Italic Capital)
        0x0041..=0x005A => char::from_u32(cp - 0x41 + 0x1D434).unwrap_or(c),

        // Latin lowercase h → U+210E (Planck constant; U+1D455 is unassigned)
        0x0068 => '\u{210E}',

        // Latin lowercase a-g, i-z → U+1D44E..U+1D467 (Mathematical Italic Small)
        // Contiguous range minus h; the math puts h at U+210E externally.
        0x0061..=0x0067 | 0x0069..=0x007A => {
            char::from_u32(cp - 0x61 + 0x1D44E).unwrap_or(c)
        }

        // Greek lowercase α-ω (includes ς U+03C2) → U+1D6FC..U+1D714
        // Linear offset: cp + 0x1D34B
        0x03B1..=0x03C9 => char::from_u32(cp + 0x1D34B).unwrap_or(c),

        // Greek uppercase Α-Ω (includes unassigned slot U+03A2) → U+1D6E2..U+1D6FA
        // Linear offset: cp + 0x1D351
        0x0391..=0x03A9 => char::from_u32(cp + 0x1D351).unwrap_or(c),

        // Greek variant forms: non-linear, explicit assignments in Math Italic Greek
        0x03D1 => '\u{1D717}', // ϑ theta-symbol (script) → Math Italic Theta Symbol
        0x03D5 => '\u{1D719}', // ϕ phi-symbol (closed)   → Math Italic Phi Symbol
        0x03D6 => '\u{1D71B}', // ϖ pi-symbol             → Math Italic Pi Symbol
        0x03F1 => '\u{1D71A}', // ϱ rho-symbol            → Math Italic Rho Symbol
        0x03F5 => '\u{1D716}', // ϵ lunate epsilon        → Math Italic Epsilon Symbol

        // ASCII hyphen-minus → proper Minus Sign
        0x002D => '\u{2212}',

        // Partial differential ∂ → Math Italic Partial Differential
        0x2202 => '\u{1D715}',

        // Nabla ∇ → Math Italic Nabla
        0x2207 => '\u{1D6FB}',

        // Everything else unchanged: digits 0-9, all other operators,
        // relations (= ≠ ≤ ≥ ∈ ⊂ etc.), quantifiers (∀ ∃), arrows (→ ⇒),
        // blackboard bold letters (ℕ ℤ ℝ), infinity ∞, etc.
        _ => c,
    }
}

/// Substitute every char in a string.
pub fn math_substitute_str(s: &str) -> String {
    s.chars().map(math_substitute).collect()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn latin_uppercase_all() {
        // A-Z should map contiguously to U+1D434..U+1D44D
        for (i, c) in ('A'..='Z').enumerate() {
            let expected = char::from_u32(0x1D434 + i as u32).unwrap();
            assert_eq!(math_substitute(c), expected, "input {c:?}");
        }
    }

    #[test]
    fn latin_lowercase_non_h() {
        // a..g → 1D44E..1D454
        for (i, c) in ('a'..='g').enumerate() {
            let expected = char::from_u32(0x1D44E + i as u32).unwrap();
            assert_eq!(math_substitute(c), expected, "input {c:?}");
        }
        // i..z → 1D456..1D467 (note: skip U+1D455 slot)
        for (i, c) in ('i'..='z').enumerate() {
            let expected = char::from_u32(0x1D456 + i as u32).unwrap();
            assert_eq!(math_substitute(c), expected, "input {c:?}");
        }
    }

    #[test]
    fn h_is_planck_constant() {
        assert_eq!(math_substitute('h'), '\u{210E}');
        assert_eq!(math_substitute('h') as u32, 0x210E);
        // Verify the natural slot U+1D455 is NOT used
        assert_ne!(math_substitute('h') as u32, 0x1D455);
    }

    #[test]
    fn digits_unchanged() {
        for c in '0'..='9' {
            assert_eq!(math_substitute(c), c);
        }
    }

    #[test]
    fn greek_lowercase_full_range() {
        let greek_lower = ['α', 'β', 'γ', 'δ', 'ε', 'ζ', 'η', 'θ', 'ι', 'κ',
                          'λ', 'μ', 'ν', 'ξ', 'ο', 'π', 'ρ', 'ς', 'σ', 'τ',
                          'υ', 'φ', 'χ', 'ψ', 'ω'];
        for c in greek_lower {
            let cp = c as u32;
            let expected = char::from_u32(cp + 0x1D34B).unwrap();
            assert_eq!(math_substitute(c), expected, "input {c:?}");
        }
    }

    #[test]
    fn greek_uppercase_full_range() {
        // Α through Ω, skipping the unassigned slot at U+03A2
        let greek_upper = ['Α', 'Β', 'Γ', 'Δ', 'Ε', 'Ζ', 'Η', 'Θ', 'Ι', 'Κ',
                          'Λ', 'Μ', 'Ν', 'Ξ', 'Ο', 'Π', 'Ρ', 'Σ', 'Τ', 'Υ',
                          'Φ', 'Χ', 'Ψ', 'Ω'];
        for c in greek_upper {
            let cp = c as u32;
            let expected = char::from_u32(cp + 0x1D351).unwrap();
            assert_eq!(math_substitute(c), expected, "input {c:?}");
        }
    }

    #[test]
    fn greek_variants_explicit() {
        assert_eq!(math_substitute('ϑ') as u32, 0x1D717);
        assert_eq!(math_substitute('ϕ') as u32, 0x1D719);
        assert_eq!(math_substitute('ϖ') as u32, 0x1D71B);
        assert_eq!(math_substitute('ϱ') as u32, 0x1D71A);
        assert_eq!(math_substitute('ϵ') as u32, 0x1D716);
    }

    #[test]
    fn operator_substitutions() {
        assert_eq!(math_substitute('-') as u32, 0x2212); // hyphen → minus
        assert_eq!(math_substitute('∂') as u32, 0x1D715); // partial
        assert_eq!(math_substitute('∇') as u32, 0x1D6FB); // nabla
    }

    #[test]
    fn operators_unchanged() {
        // From COM measurement: 47 operators unchanged
        let unchanged = ['∑', '∏', '∫', '∮', '⋀', '⋁', '⋂', '⋃',
                        '=', '≠', '<', '>', '≤', '≥', '≈', '≡',
                        '∈', '∉', '⊂', '⊃', '⊆', '⊇',
                        '+', '×', '÷', '·', '∗', '±', '∓',
                        '∞', '∀', '∃', '∄',
                        '∅', 'ℕ', 'ℤ', 'ℚ', 'ℝ', 'ℂ',
                        '→', '⇒', '↔', '⇔', '↦'];
        for c in unchanged {
            assert_eq!(math_substitute(c), c, "input {c:?} was changed unexpectedly");
        }
    }

    #[test]
    fn math_substitute_str_example() {
        // "ax+b=h" should become italic variants + unchanged operator
        let s = math_substitute_str("ax+b=h");
        // Expected: 𝑎𝑥+𝑏=ℎ (a, x, b italic; h → Planck constant; + = unchanged)
        assert_eq!(s, "\u{1D44E}\u{1D465}+\u{1D44F}=\u{210E}");
    }

    #[test]
    fn idempotent_on_already_substituted() {
        // Once in the Math Italic range, further substitution should be no-op.
        assert_eq!(math_substitute('\u{1D434}'), '\u{1D434}'); // already Math Italic A
        assert_eq!(math_substitute('\u{210E}'), '\u{210E}');   // Planck constant
        assert_eq!(math_substitute('\u{1D6FC}'), '\u{1D6FC}'); // Math Italic alpha
    }

    #[test]
    fn japanese_unchanged() {
        // CJK characters must NOT be substituted — OMML users may include
        // Japanese/Chinese labels in math runs.
        for c in ['日', '本', '国', '語', 'ア', 'イ', 'ウ'] {
            assert_eq!(math_substitute(c), c);
        }
    }
}
