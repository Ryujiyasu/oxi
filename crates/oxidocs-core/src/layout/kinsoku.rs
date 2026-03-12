/// Japanese line-breaking rules based on JIS X 4051
///
/// Kinsoku processing prevents certain characters from appearing at the
/// start or end of a line.

/// Characters prohibited at the start of a line (行頭禁則文字)
/// Closing brackets, punctuation, small kana, etc.
const LINE_START_PROHIBITED: &[char] = &[
    // Closing brackets
    '）', '〕', '］', '｝', '〉', '》', '」', '』', '】', '〙', '〗',
    ')', ']', '}',
    // Punctuation that shouldn't start a line
    '、', '。', '，', '．', '：', '；', '？', '！',
    '・', 'ー', '～',
    // Small kana
    'ぁ', 'ぃ', 'ぅ', 'ぇ', 'ぉ', 'っ', 'ゃ', 'ゅ', 'ょ', 'ゎ',
    'ァ', 'ィ', 'ゥ', 'ェ', 'ォ', 'ッ', 'ャ', 'ュ', 'ョ', 'ヮ',
    // Prolonged sound mark
    'ヽ', 'ヾ', 'ゝ', 'ゞ',
    // Period and comma variants
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

/// Check if a character is CJK (Chinese, Japanese, Korean)
/// These characters can have line breaks between any two adjacent characters
/// (subject to kinsoku rules)
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
        // Halfwidth and Fullwidth Forms
        0xFF00..=0xFFEF |
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
