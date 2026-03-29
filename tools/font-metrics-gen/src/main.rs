use serde::Serialize;
use std::collections::BTreeMap;
use std::path::Path;
use ttf_parser::Face;

/// Compact font metrics output for a single font face.
#[derive(Serialize)]
struct FontMetricsData {
    family: String,
    units_per_em: u16,
    ascender: i16,
    descender: i16,
    line_gap: i16,
    /// OS/2 usWinAscent (used by Word for line height calculation)
    win_ascent: u16,
    /// OS/2 usWinDescent (used by Word for line height calculation)
    win_descent: u16,
    /// OS/2 sTypoAscender
    typo_ascender: i16,
    /// OS/2 sTypoDescender
    typo_descender: i16,
    /// OS/2 sTypoLineGap
    typo_line_gap: i16,
    /// Advance widths keyed by codepoint, in font units.
    /// Only stores characters we care about (ASCII, CJK punct, kana, etc.)
    widths: BTreeMap<u32, u16>,
}

/// Font files to extract metrics from: (filename, display_name, face_index).
const FONTS: &[(&str, &str, u32)] = &[
    ("calibri.ttf", "Calibri", 0),
    ("calibrib.ttf", "Calibri Bold", 0),
    ("times.ttf", "Times New Roman", 0),
    ("timesbd.ttf", "Times New Roman Bold", 0),
    ("arial.ttf", "Arial", 0),
    ("arialbd.ttf", "Arial Bold", 0),
    ("CENTURY.TTF", "Century", 0),
    ("msgothic.ttc", "MS Gothic", 0),
    ("msgothic.ttc", "MS PGothic", 1),
    ("msmincho.ttc", "MS Mincho", 0),
    ("msmincho.ttc", "MS PMincho", 1),
    ("YuGothR.ttc", "Yu Gothic Regular", 0),
    ("YuGothB.ttc", "Yu Gothic Bold", 0),
    ("yumin.ttf", "Yu Mincho Regular", 0),
    ("yumindb.ttf", "Yu Mincho Demibold", 0),
    ("cambria.ttc", "Cambria", 0),
    ("cambriab.ttf", "Cambria Bold", 0),
    ("meiryo.ttc", "Meiryo", 0),
    // OSS metric-compatible fonts
    ("Carlito-Regular.ttf", "Carlito", 0),
    ("Carlito-Bold.ttf", "Carlito Bold", 0),
    ("Caladea-Regular.ttf", "Caladea", 0),
    ("Caladea-Bold.ttf", "Caladea Bold", 0),
    ("LiberationSans-Regular.ttf", "Liberation Sans", 0),
    ("LiberationSans-Bold.ttf", "Liberation Sans Bold", 0),
    ("LiberationSerif-Regular.ttf", "Liberation Serif", 0),
    ("LiberationSerif-Bold.ttf", "Liberation Serif Bold", 0),
    ("NotoSansJP-VF.ttf", "Noto Sans JP", 0),
    ("NotoSerifJP-VF.ttf", "Noto Serif JP", 0),
];

/// Character ranges to extract widths for.
fn chars_to_measure() -> Vec<char> {
    let mut chars = Vec::new();

    // ASCII printable (U+0020..U+007E)
    for cp in 0x0020u32..=0x007E {
        if let Some(c) = char::from_u32(cp) {
            chars.push(c);
        }
    }

    // CJK Symbols and Punctuation (U+3000..U+303F)
    for cp in 0x3000u32..=0x303F {
        if let Some(c) = char::from_u32(cp) {
            chars.push(c);
        }
    }

    // Hiragana (U+3040..U+309F)
    for cp in 0x3040u32..=0x309F {
        if let Some(c) = char::from_u32(cp) {
            chars.push(c);
        }
    }

    // Katakana (U+30A0..U+30FF)
    for cp in 0x30A0u32..=0x30FF {
        if let Some(c) = char::from_u32(cp) {
            chars.push(c);
        }
    }

    // CJK Unified Ideographs: full BMP range (U+4E00..U+9FFF)
    // Covers JIS Level 1-2, Joyo Kanji, and all common kanji used in Japanese documents.
    // For monospace CJK fonts (MS Mincho/Gothic) all are 1em, but proportional fonts
    // (Meiryo, Yu Gothic) may have varying widths for rare characters.
    for cp in 0x4E00u32..=0x9FFF {
        if let Some(c) = char::from_u32(cp) {
            chars.push(c);
        }
    }

    // CJK Compatibility Ideographs (U+F900..U+FAFF) - some used in Japanese
    for cp in 0xF900u32..=0xFAFF {
        if let Some(c) = char::from_u32(cp) {
            chars.push(c);
        }
    }

    // Special symbols used in Japanese documents
    for c in "㎡㎥㎢㏄㎝㎜㎞㎏㎎㍻㍼㍽㍾①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳".chars() {
        chars.push(c);
    }

    // Fullwidth Latin (U+FF01..U+FF60)
    for cp in 0xFF01u32..=0xFF60 {
        if let Some(c) = char::from_u32(cp) {
            chars.push(c);
        }
    }

    // Halfwidth Katakana (U+FF65..U+FF9F)
    for cp in 0xFF65u32..=0xFF9F {
        if let Some(c) = char::from_u32(cp) {
            chars.push(c);
        }
    }

    // Fullwidth currency/symbols (U+FFE0..U+FFE6)
    for cp in 0xFFE0u32..=0xFFE6 {
        if let Some(c) = char::from_u32(cp) {
            chars.push(c);
        }
    }

    // Deduplicate
    chars.sort();
    chars.dedup();
    chars
}

fn extract_face(face: &Face, name: &str) -> FontMetricsData {
    let units_per_em = face.units_per_em();
    let ascender = face.ascender();
    let descender = face.descender();
    let line_gap = face.line_gap();

    // Extract OS/2 table metrics (used by Word for line height)
    let (win_ascent, win_descent, typo_ascender, typo_descender, typo_line_gap) =
        if let Some(os2) = face.tables().os2 {
            (
                os2.windows_ascender() as u16,
                os2.windows_descender().unsigned_abs(),
                os2.typographic_ascender(),
                os2.typographic_descender(),
                os2.typographic_line_gap(),
            )
        } else {
            (ascender as u16, (-descender) as u16, ascender, descender, line_gap)
        };

    let mut widths = BTreeMap::new();
    let chars = chars_to_measure();

    for ch in &chars {
        if let Some(glyph_id) = face.glyph_index(*ch) {
            let advance = face.glyph_hor_advance(glyph_id).unwrap_or(0);
            widths.insert(*ch as u32, advance);
        }
    }

    FontMetricsData {
        family: name.to_string(),
        units_per_em,
        ascender,
        descender,
        line_gap,
        win_ascent,
        win_descent,
        typo_ascender,
        typo_descender,
        typo_line_gap,
        widths,
    }
}

fn main() {
    let fonts_dir = std::env::args()
        .nth(1)
        .unwrap_or_else(|| "C:/Windows/Fonts".to_string());
    let output_dir = std::env::args()
        .nth(2)
        .unwrap_or_else(|| "output".to_string());

    std::fs::create_dir_all(&output_dir).expect("Failed to create output directory");

    let mut all_metrics: Vec<FontMetricsData> = Vec::new();

    for &(filename, display_name, face_index) in FONTS {
        let path = Path::new(&fonts_dir).join(filename);
        if !path.exists() {
            eprintln!("SKIP: {} not found at {}", filename, path.display());
            continue;
        }

        let data = std::fs::read(&path).expect("Failed to read font file");

        match Face::parse(&data, face_index) {
            Ok(face) => {
                let metrics = extract_face(&face, display_name);
                eprintln!(
                    "  {} (face {}) => {} glyphs, UPM={}, asc={}, desc={}, gap={}, winAsc={}, winDesc={}",
                    display_name,
                    face_index,
                    metrics.widths.len(),
                    metrics.units_per_em,
                    metrics.ascender,
                    metrics.descender,
                    metrics.line_gap,
                    metrics.win_ascent,
                    metrics.win_descent,
                );
                all_metrics.push(metrics);
            }
            Err(e) => eprintln!("  ERROR parsing {} face {}: {}", filename, face_index, e),
        }
    }

    // Write combined JSON
    let combined_path = Path::new(&output_dir).join("font_metrics.json");
    let json = serde_json::to_string_pretty(&all_metrics).expect("Failed to serialize");
    std::fs::write(&combined_path, &json).expect("Failed to write");
    eprintln!(
        "\nWrote {} font metrics to {}",
        all_metrics.len(),
        combined_path.display()
    );

    // Also write a compact version (no pretty-print) for embedding
    let compact_path = Path::new(&output_dir).join("font_metrics_compact.json");
    let compact = serde_json::to_string(&all_metrics).expect("Failed to serialize");
    std::fs::write(&compact_path, &compact).expect("Failed to write");
    eprintln!(
        "Compact version: {} ({} bytes)",
        compact_path.display(),
        compact.len()
    );
}
