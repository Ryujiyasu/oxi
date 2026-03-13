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

/// Font files to extract metrics from, with display names.
const FONTS: &[(&str, &str)] = &[
    ("calibri.ttf", "Calibri"),
    ("calibrib.ttf", "Calibri Bold"),
    ("times.ttf", "Times New Roman"),
    ("timesbd.ttf", "Times New Roman Bold"),
    ("arial.ttf", "Arial"),
    ("arialbd.ttf", "Arial Bold"),
    ("CENTURY.TTF", "Century"),
    ("msgothic.ttc", "MS Gothic"),
    ("msmincho.ttc", "MS Mincho"),
    ("YuGothR.ttc", "Yu Gothic Regular"),
    ("YuGothB.ttc", "Yu Gothic Bold"),
    ("yumin.ttf", "Yu Mincho Regular"),
    ("yumindb.ttf", "Yu Mincho Demibold"),
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

    // Common CJK Ideographs (sample: 一〜龍 first 200 + common chars)
    for cp in 0x4E00u32..=0x4E00 + 200 {
        if let Some(c) = char::from_u32(cp) {
            chars.push(c);
        }
    }
    // Additional common kanji
    for c in "日本語文字入力変換漢字仮名表示処理実行結果確認完了".chars() {
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

    for &(filename, display_name) in FONTS {
        let path = Path::new(&fonts_dir).join(filename);
        if !path.exists() {
            eprintln!("SKIP: {} not found at {}", filename, path.display());
            continue;
        }

        let data = std::fs::read(&path).expect("Failed to read font file");

        // Handle .ttc (TrueType Collection) - extract first face
        if filename.ends_with(".ttc") {
            let face_count = ttf_parser::fonts_in_collection(&data).unwrap_or(0);
            eprintln!(
                "TTC: {} contains {} faces, extracting first",
                filename, face_count
            );
            match Face::parse(&data, 0) {
                Ok(face) => {
                    let metrics = extract_face(&face, display_name);
                    eprintln!(
                        "  {} => {} glyphs, UPM={}, asc={}, desc={}, gap={}, winAsc={}, winDesc={}",
                        display_name,
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
                Err(e) => eprintln!("  ERROR parsing {}: {}", filename, e),
            }
        } else {
            match Face::parse(&data, 0) {
                Ok(face) => {
                    let metrics = extract_face(&face, display_name);
                    eprintln!(
                        "  {} => {} glyphs, UPM={}, asc={}, desc={}, gap={}, winAsc={}, winDesc={}",
                        display_name,
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
                Err(e) => eprintln!("  ERROR parsing {}: {}", filename, e),
            }
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
