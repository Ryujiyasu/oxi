// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Integration tests: parse `word/settings.xml` end-to-end and verify
//! the six `Document` fields populated from it after `parse_docx`.
//!
//! Parser code paths tested:
//!   - [parser/ooxml.rs:555](crates/oxidocs-core/src/parser/ooxml.rs#L555)
//!     parse_adjust_line_height_in_table (substring search)
//!   - [parser/ooxml.rs:564](crates/oxidocs-core/src/parser/ooxml.rs#L564)
//!     parse_compat_mode (compatSetting w:name="compatibilityMode")
//!   - [parser/ooxml.rs:598](crates/oxidocs-core/src/parser/ooxml.rs#L598)
//!     parse_compress_punctuation (characterSpacingControl)
//!   - [parser/ooxml.rs:628](crates/oxidocs-core/src/parser/ooxml.rs#L628)
//!     parse_compat_bool_flag (generic <w:flag/> reader)
//!   - [parser/ooxml.rs:657](crates/oxidocs-core/src/parser/ooxml.rs#L657)
//!     parse_default_tab_stop (defaultTabStop val twips → pt)
//!
//! Non-obvious behaviors pinned:
//!   - compat_mode defaults to 15 (Word 2013+) on absent
//!     compatSetting — NOT to 14. Per parser/ooxml.rs:583,592 the
//!     fallback is the modern Word value.
//!   - compress_punctuation accepts TWO distinct val strings as
//!     true ("compressPunctuation" AND
//!     "compressPunctuationAndJapaneseKana"). "doNotCompress" is
//!     false (same end state as absent).
//!   - parse_compat_bool_flag TRI-STATE: <flag/> → TRUE,
//!     <flag w:val="0"/> → FALSE, <flag w:val="false"/> → FALSE,
//!     absent → FALSE. Presence-no-val vs absent are
//!     distinguishable at the source XML level but produce the
//!     same IR boolean.
//!   - adjust_line_height_in_table uses SUBSTRING SEARCH
//!     (parser/ooxml.rs:560) NOT real XML parsing. A future "proper
//!     parse" refactor must preserve this behavior or risk
//!     regressing docs that depend on the substring shortcut.
//!
//! Fixtures live in `tools/fixtures/settings_samples/` and are
//! authored by `tools/metrics/build_settings_repro_fixtures.py`.

use std::fs;

use oxidocs_core::ir::Document;
use oxidocs_core::parse_docx;

fn fixture_path(name: &str) -> std::path::PathBuf {
    let crate_root = std::env::current_dir().unwrap();
    let workspace_root = crate_root.parent().unwrap().parent().unwrap();
    workspace_root
        .join("tools")
        .join("fixtures")
        .join("settings_samples")
        .join(name)
}

fn load(name: &str) -> Option<Document> {
    let path = fixture_path(name);
    if !path.exists() {
        eprintln!("skipping: {} not found", path.display());
        return None;
    }
    let data = fs::read(&path).expect("read fixture");
    Some(parse_docx(&data).expect("parse fixture"))
}

#[test]
fn v1_settings_all_features_on_populate_every_field() {
    let Some(doc) = load("v1_settings_all_features_on.docx") else { return };

    // compat_mode=14 (Word 2010) — explicitly DIFFERENT from the
    // default 15. A regression that ignored compatSetting would
    // show 15 here.
    assert_eq!(doc.compat_mode, 14, "compatibilityMode=14 parsed");

    // characterSpacingControl="compressPunctuation" → true.
    assert!(
        doc.compress_punctuation,
        "characterSpacingControl=compressPunctuation → compress_punctuation=true"
    );

    // The three compat-bool flags, each via self-closing presence.
    assert!(
        doc.do_not_expand_shift_return,
        "<w:doNotExpandShiftReturn/> → true"
    );
    assert!(
        doc.balance_single_byte_double_byte_width,
        "<w:balanceSingleByteDoubleByteWidth/> → true"
    );
    assert!(
        doc.adjust_line_height_in_table,
        "<w:adjustLineHeightInTable/> → true (via substring search)"
    );

    // defaultTabStop=708 → 35.4pt (val/20).
    let dts = doc
        .default_tab_stop
        .expect("defaultTabStop present must populate Some");
    assert!(
        (dts - 35.4).abs() < 0.001,
        "defaultTabStop=708 → 35.4pt (twips/20), got {}",
        dts
    );
}

#[test]
fn v1_settings_minimal_defaults_each_field() {
    let Some(doc) = load("v1_settings_minimal_defaults.docx") else { return };

    // Empty <w:settings> → ALL defaults. This is the FALLBACK
    // path for every settings parser. A regression that changed
    // any default would surface here.

    // compat_mode default = 15 (Word 2013+), NOT 14. Parser
    // explicitly favors modern Word (parser/ooxml.rs:583,592).
    assert_eq!(
        doc.compat_mode, 15,
        "absent compatSetting → default 15 (NOT 14)"
    );

    assert!(
        !doc.compress_punctuation,
        "absent characterSpacingControl → compress_punctuation=false (default)"
    );
    assert!(!doc.do_not_expand_shift_return);
    assert!(!doc.balance_single_byte_double_byte_width);
    assert!(!doc.adjust_line_height_in_table);
    assert!(
        doc.default_tab_stop.is_none(),
        "absent defaultTabStop → None (NOT a default value like 36.0)"
    );
}

#[test]
fn v1_settings_yakumono_kana_variant_also_true() {
    // The OTHER valid val string for compress_punctuation=true.
    // Pinning both variants protects the "OR" branch at
    // parser/ooxml.rs:611-612 from being narrowed to a single
    // string match in a future refactor.
    let Some(doc) = load("v1_settings_yakumono_kana_variant.docx") else { return };
    assert!(
        doc.compress_punctuation,
        "characterSpacingControl=compressPunctuationAndJapaneseKana → true \
         (the kana variant, ALSO accepted alongside compressPunctuation)"
    );
}

#[test]
fn v1_settings_yakumono_donotcompress_distinguishes_explicit_from_absent() {
    // Explicit "doNotCompress" → false. SAME end state as absent,
    // but the explicit suppression catches a future regression
    // that flips the default for the absent case.
    let Some(doc) = load("v1_settings_yakumono_donotcompress.docx") else { return };
    assert!(
        !doc.compress_punctuation,
        "characterSpacingControl=doNotCompress → false (explicit suppression)"
    );
}

#[test]
fn v1_settings_compat_flag_val_zero_suppresses_polarity() {
    // The compat-bool tri-state. <flag/> (presence-no-val) → TRUE,
    // <flag w:val="0"/> → FALSE, <flag w:val="false"/> → FALSE.
    // BOTH "0" AND "false" trigger the suppression branch
    // (parser/ooxml.rs:642).
    let Some(doc) = load("v1_settings_compat_flag_val_zero_suppresses.docx")
    else { return };

    assert!(
        !doc.do_not_expand_shift_return,
        "<w:doNotExpandShiftReturn w:val=\"0\"/> → false (val=0 SUPPRESSES \
         the presence-default true)"
    );
    assert!(
        !doc.balance_single_byte_double_byte_width,
        "<w:balanceSingleByteDoubleByteWidth w:val=\"false\"/> → false \
         (val=\"false\" also suppresses, parser accepts BOTH \"0\" AND \"false\")"
    );
}

#[test]
fn all_five_fixtures_parse_with_expected_one_page() {
    // Smoke + structural: each fixture parses without error and
    // produces exactly one page (the document.xml is identical
    // across fixtures — only settings.xml varies).
    let cases: &[&str] = &[
        "v1_settings_all_features_on.docx",
        "v1_settings_minimal_defaults.docx",
        "v1_settings_yakumono_kana_variant.docx",
        "v1_settings_yakumono_donotcompress.docx",
        "v1_settings_compat_flag_val_zero_suppresses.docx",
    ];
    for name in cases {
        let path = fixture_path(name);
        if !path.exists() {
            eprintln!("skipping {}", name);
            continue;
        }
        let data = fs::read(&path).unwrap();
        let doc = parse_docx(&data)
            .unwrap_or_else(|e| panic!("failed to parse {}: {:?}", name, e));
        assert_eq!(doc.pages.len(), 1, "{} should produce 1 page", name);
    }
}
