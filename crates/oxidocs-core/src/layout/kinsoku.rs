/// Japanese line-breaking rules based on JIS X 4051
///
/// Kinsoku processing prevents certain characters from appearing at the
/// start or end of a line.

/// Characters prohibited at the start of a line (行頭禁則文字)
/// Word's default (Standard Microsoft) kinsoku — verified via COM (2026-04-08).
/// Word allows small kana ぁぃっゃゅょ etc, prolonged sound mark ー, and 〜
/// at line start despite strict JIS X 4051 prohibition. Only the chars
/// below are actually blocked by Word's default kinsoku.
const LINE_START_PROHIBITED: &[char] = &[
    // Closing brackets
    '）', '〕', '］', '｝', '〉', '》', '」', '』', '】', '〙', '〗',
    ')', ']', '}',
    // Punctuation that shouldn't start a line
    '、', '。', '，', '．', '：', '；', '？', '！',
    // Mid-dot
    '・',
    // Repetition marks
    'ヽ', 'ヾ', 'ゝ', 'ゞ',
    // Period/comma variants
    '‥', '…',
];

/// Characters prohibited at the end of a line (行末禁則文字)
/// Opening brackets
const LINE_END_PROHIBITED: &[char] = &[
    '（', '〔', '［', '｛', '〈', '《', '「', '『', '【', '〘', '〖',
    '(', '[', '{',
];

/// Check if a character is prohibited at the start of a line
pub fn is_line_start_prohibited(ch: char) -> bool {
    LINE_START_PROHIBITED.contains(&ch)
}

/// Check if a character is prohibited at the end of a line
pub fn is_line_end_prohibited(ch: char) -> bool {
    LINE_END_PROHIBITED.contains(&ch)
}

/// Characters that Word allows to "hang" past the right margin (burasagari).
/// COM-confirmed (2026-04-08): Word's HangingPunctuation=true (default) hangs
/// CJK closing brackets and CJK comma/period — but NOT colon/semicolon/?/!.
/// See memory/hangable_oikomi_rule.md.
const HANGABLE_PUNCT: &[char] = &[
    // CJK comma and period (and fullwidth)
    '、', '。', '，', '．',
    // CJK closing brackets
    '）', '〕', '］', '｝', '〉', '》', '」', '』', '】', '〙', '〗',
];

/// Check if a character is allowed to hang past the right margin
/// (burasagari / hanging punctuation).
pub fn is_hangable_punct(ch: char) -> bool {
    HANGABLE_PUNCT.contains(&ch)
}

/// CJK punctuation that can be compressed from full-width to half-width (50% compression).
/// These are full-width punctuation marks where Word compresses the whitespace built into
/// the glyph for justification purposes.
const CJK_COMPRESSIBLE_PUNCTUATION: &[char] = &[
    // Ideographic comma and period
    '、', '。',
    // Fullwidth comma and period
    '，', '．',
    // CJK brackets
    '「', '」', '『', '』', '（', '）', '〔', '〕', '［', '］', '｛', '｝',
    '〈', '〉', '《', '》', '【', '】', '〘', '〙', '〖', '〗',
    // Fullwidth forms
    '（', '）',
    // Colon, semicolon fullwidth
    '：', '；',
    // Fullwidth question/exclamation
    '？', '！',
];

/// Check if a CJK punctuation character can be compressed (50% width reduction)
pub fn is_cjk_compressible(ch: char) -> bool {
    CJK_COMPRESSIBLE_PUNCTUATION.contains(&ch)
}

// =====================================================================
// Yakumono compression (約物詰め) - COM-confirmed (2026-04-08)
// See memory/yakumono_compression_spec.md for full ruleset.
//
// Word applies these rules during normal layout (NOT just justify).
// Two adjacent CJK punctuation chars: ONE of them is compressed to 50%.
//   - Closing-side punct: compressed when NEXT char is a trigger
//   - Opening-side punct: compressed when PREV char is a trigger
//                         (and the prev char is not itself compressed)
// =====================================================================

/// Closing-side punctuation that compresses (50%) when followed by a trigger.
/// These have built-in right-side spacing in the glyph that gets removed.
const YAKUMONO_CLOSING: &[char] = &[
    '）', '」', '』', '〕', '】', '》', '〙', '〗', '｝', '］',
    '、', '。', '，', '．',
];

/// Opening-side punctuation that compresses (50%) when preceded by a trigger.
/// These have built-in left-side spacing that gets removed.
const YAKUMONO_OPENING: &[char] = &[
    '（', '「', '『', '〔', '【', '《', '〘', '〖', '｛', '［',
];

/// Trigger chars: presence triggers compression of an adjacent closing/opening punct.
/// Includes all closing/opening punct PLUS special triggers (・：；) that are
/// triggers but NOT compressible themselves.
const YAKUMONO_TRIGGER: &[char] = &[
    // openers
    '（', '「', '『', '〔', '【', '《', '〘', '〖', '｛', '［',
    // closers
    '）', '」', '』', '〕', '】', '》', '〙', '〗', '｝', '］',
    // commas/periods
    '、', '。', '，', '．',
    // special triggers (themselves uncompressed): middle dot, colon, semicolon
    '・', '：', '；',
];

/// Check if a char is a closing-side compressible punct.
pub fn is_yakumono_closing(ch: char) -> bool {
    YAKUMONO_CLOSING.contains(&ch)
}

/// Check if a char is an opening-side compressible punct.
pub fn is_yakumono_opening(ch: char) -> bool {
    YAKUMONO_OPENING.contains(&ch)
}

/// Check if a char triggers yakumono compression on adjacent puncts.
pub fn is_yakumono_trigger(ch: char) -> bool {
    YAKUMONO_TRIGGER.contains(&ch)
}

/// Check if a character is CJK (Chinese, Japanese, Korean)
/// These characters can have line breaks between any two adjacent characters
/// (subject to kinsoku rules)
/// CJK ideograph or kana (NOT punctuation/symbols).
/// Used for autoSpaceDE: Word adds 2.5pt only between Latin and CJK ideographs/kana.
/// Punctuation like （、。は does NOT trigger auto-space.
pub fn is_cjk_ideograph_or_kana(ch: char) -> bool {
    matches!(ch as u32,
        // CJK Unified Ideographs
        0x4E00..=0x9FFF |
        // CJK Unified Ideographs Extension A
        0x3400..=0x4DBF |
        // CJK Compatibility Ideographs
        0xF900..=0xFAFF |
        // Hiragana (excluding 0x3000-0x303F punctuation)
        0x3041..=0x309F |
        // Katakana
        0x30A1..=0x30FF |
        // Katakana Phonetic Extensions
        0x31F0..=0x31FF |
        // CJK Unified Ideographs Extension B
        0x20000..=0x2A6DF
    )
}

pub fn is_cjk(ch: char) -> bool {
    matches!(ch as u32,
        // CJK Unified Ideographs
        0x4E00..=0x9FFF |
        // CJK Unified Ideographs Extension A
        0x3400..=0x4DBF |
        // CJK Compatibility Ideographs
        0xF900..=0xFAFF |
        // Hiragana
        0x3040..=0x309F |
        // Katakana
        0x30A0..=0x30FF |
        // Katakana Phonetic Extensions
        0x31F0..=0x31FF |
        // CJK Symbols and Punctuation
        0x3000..=0x303F |
        // Enclosed CJK Letters and Months (㊀, ㈱, etc.)
        0x3200..=0x32FF |
        // CJK Compatibility (㎡, ㎞, ㎏, etc.)
        0x3300..=0x33FF |
        // Halfwidth and Fullwidth Forms
        0xFF00..=0xFFEF |
        // General Punctuation (※, †, ‡, etc.) — Word uses East Asian font
        0x2010..=0x2044 |
        // Geometric Shapes (□, ○, ◎, ●, △, ▲, etc.)
        0x25A0..=0x25FF |
        // Miscellaneous Symbols (☆, ★, ♪, etc.)
        0x2600..=0x26FF |
        // Dingbats (✓, ✕, etc.)
        0x2700..=0x27BF |
        // Box Drawing (─, │, ┌, etc.)
        0x2500..=0x257F |
        // Block Elements (▀, ▄, █, etc.)
        0x2580..=0x259F |
        // Arrows (←, ↑, →, ↓, etc.)
        0x2190..=0x21FF |
        // Mathematical Operators (×, ÷, ±, etc.)
        0x2200..=0x22FF |
        // Latin-1 math symbols Word renders with East Asian font
        0x00D7 | // × multiplication sign
        0x00F7 | // ÷ division sign
        // CJK Unified Ideographs Extension B
        0x20000..=0x2A6DF
    )
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_line_start_prohibited() {
        assert!(is_line_start_prohibited('。'));
        assert!(is_line_start_prohibited('、'));
        assert!(is_line_start_prohibited('）'));
        assert!(is_line_start_prohibited('っ'));
        assert!(!is_line_start_prohibited('あ'));
        assert!(!is_line_start_prohibited('A'));
    }

    #[test]
    fn test_line_end_prohibited() {
        assert!(is_line_end_prohibited('（'));
        assert!(is_line_end_prohibited('「'));
        assert!(!is_line_end_prohibited('あ'));
        assert!(!is_line_end_prohibited('。'));
    }

    #[test]
    fn test_is_cjk() {
        assert!(is_cjk('漢'));
        assert!(is_cjk('あ'));
        assert!(is_cjk('ア'));
        assert!(is_cjk('、'));
        assert!(is_cjk('。'));
        assert!(!is_cjk('A'));
        assert!(!is_cjk('1'));
        assert!(!is_cjk(' '));
    }
}
