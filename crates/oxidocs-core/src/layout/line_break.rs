// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! `LayoutEngine::break_into_lines_with_grid` -- moved out of `layout/mod.rs` so that it is its own
//! codegen unit (see tools/metrics/split_layout_mod.py). Behaviour-preserving.

use super::*;

/// `break_into_lines_with_grid` as a method of its own type: rustc puts a method's code in the
/// codegen unit of its self type's module, so this (not the file move alone)
/// is what gives the giant its own unit. Deref keeps `self.x` meaning the engine.
pub(super) struct LineBreaker<'a>(pub(super) &'a LayoutEngine);

impl<'a> std::ops::Deref for LineBreaker<'a> {
    type Target = LayoutEngine;
    fn deref(&self) -> &LayoutEngine {
        self.0
    }
}

impl<'a> LineBreaker<'a> {
    // Word's quarter-em automatic gap follows the preceding visible glyph's
    // size. With a 10pt ideograph followed by an 8/10/12pt Latin glyph, Word
    // PDF measures the same 2.52pt leading gap in all three arms.
    fn autospace_after_style(&self, ch: char, style: &RunStyle, para: &ParagraphStyle) -> f32 {
        let size = self.resolve_font_size(style, para);
        let size = if std::env::var("OXI_S899_DISABLE").is_err()
            && style.font_size.is_some()
            && matches!(style.vertical_align,
                Some(VerticalAlign::Superscript) | Some(VerticalAlign::Subscript))
        {
            LayoutEngine::vertical_align_font_size(size)
        } else {
            size
        };
        let cs = if style.fit_text.is_some() || style.ruby_spread {
            style.character_spacing.unwrap_or(0.0)
        } else {
            snap_character_spacing(style.character_spacing.unwrap_or(0.0))
        };
        self.natural_autospace_after(ch, style, para, size, cs)
    }

    pub(super) fn break_into_lines_with_grid(
        &self,
        fragments: &[(&str, &RunStyle, Option<FieldType>, usize, usize)],
        available_width: f32,
        first_line_indent: f32,
        para_style: &ParagraphStyle,
        grid_char_pitch: Option<f32>,
        grid_char_cw_ratio: Option<f32>,
        lines_and_chars: bool,
        s476_body: bool,
        is_justified: bool,
        doc_grid_no_type: bool,
        para_has_lrpb: bool,
        // S677: when true (a w:smallCaps paragraph), flush the word accumulator on a
        // per-fragment font_size CHANGE so the small-caps size segments (full vs 0.8×)
        // survive as their own fragments instead of merging into one (the same
        // fragment-flattening that S655 fixes for w:position). false everywhere else
        // → byte-identical (no font_size-boundary flush for normal paragraphs).
        caps_size_split: bool,
        vertical: bool,
        quantized_char_grid: bool,
    ) -> Vec<Line> {
        // Helper: convert pt to twips for Word-GDI-compatible integer comparison
        let pt_to_tw = |pt: f32| -> i32 { (pt * 20.0).round() as i32 };
        let available_tw = pt_to_tw(available_width);
        // Legacy Latin lines without justification use the same exact advance
        // sum as modern layout. Word boundary sweeps across five proportional
        // and monospaced fonts distinguish this from per-word twip rounding.
        let legacy_latin_exact = self.compat_mode == 14
            && self.compat_mode_explicit
            && !is_justified
            && !vertical
            && std::env::var_os("OXI_LEGACY_EXACT_DISABLE").is_none()
            && fragments.iter().all(|f| !f.0.chars().any(kinsoku::is_cjk));
        // S1446 (2026-09-17, default ON, opt-out OXI_S1446_DISABLE): a vertical
        // character set at a PROPORTIONAL advance (ＭＳ Ｐ明朝 kana / 、。) breaks
        // on its exact width: no compression absorb and no hang past the column
        // end (COM, tests/fixtures/vhang: ＭＳ Ｐ明朝 11pt pushes 'れ。' to the next
        // column in every arm, while ＭＳ 明朝's full-width 。 hangs under compat
        // 14). The explicit OXI_VERTICAL_NATURAL_BOUNDARY keeps the checkpoint's
        // unscoped form (every character with a table advance).
        let s1446_natural_boundary_explicit = std::env::var_os("OXI_VERTICAL_NATURAL_BOUNDARY").is_some();
        let vertical_natural_boundary_enabled = vertical
            && crate::font::vertical_font_advance_on()
            && (s1446_natural_boundary_explicit || std::env::var_os("OXI_S1446_DISABLE").is_none());
        let dbg_frags = std::env::var("OXI_DBG_FRAGS").ok().filter(|pre| {
            let head: String = fragments.iter().flat_map(|f| f.0.chars()).take(pre.chars().count()).collect();
            head == *pre
        });
        if (std::env::var("OXI_DBG1318").is_ok() && fragments.first().map_or(false, |f| f.0.starts_with('\u{203B}'))) || dbg_frags.is_some() {
            eprintln!("[FRAGS] avail={:.2} first_indent={:.2} left_tw={:?} left_chars={:?} hang_chars={:?} {:?}", available_width, first_line_indent,
                para_style.indent_left, para_style.indent_left_chars, para_style.indent_hanging_chars,
                fragments.iter().take(6).map(|f| (f.0.chars().take(6).collect::<String>(), f.1.character_spacing, f.1.font_size, f.3)).collect::<Vec<_>>());
        }

        // Day 33 part 19 (2026-05-10): paragraphs containing ONLY whitespace
        // (ASCII space, tab, U+3000 fullwidth space, etc.) render as a single
        // line in Word regardless of total natural width. COM-confirmed via
        // WS_10 / WS_50 / WS_100 / WS_142 / WS_300 minimal repros (all 5
        // produce identical 1-line break boundaries with BEFORE→AFTER advance
        // = 31pt = 1 line each). MIX_50_TEXT (50 spaces + text) DOES wrap
        // normally → rule is binary at paragraph level.
        // This is the safe inverse of commit 82de3fa (reverted 2026-05-03)
        // which used a per-character "trailing U+3000 immune" flag that
        // propagated to ALL U+3000s in a paragraph, regressing d77a mid-text
        // U+3000 indentation use. The all-whitespace gate is paragraph-binary
        // and never affects mixed-content paragraphs.
        let para_all_whitespace = fragments
            .iter()
            .all(|(text, _, _, _, _)| text.chars().all(|c| c.is_whitespace() || c == '\u{3000}'))
            && fragments.iter().any(|(text, _, _, _, _)| !text.is_empty());

        let mut lines = Vec::new();
        let mut current_line = Line {
            empty_break_style: None,
            seg2_at: None,
            fragments: vec![],
            ..Default::default()
        };
        let mut current_width = first_line_indent;
        // Integer twips accumulator for line break decisions.
        // Avoids f32 rounding drift that causes ±0.1pt error over 40+ characters.
        let mut current_width_tw: i32 = pt_to_tw(first_line_indent);
        // S475 (2026-06-01): capacity-adjusted break width = Σ(natural_adv −
        // max_yakumono_compress). Runs parallel to current_width_tw; only CONSULTED
        // when s475_break is ON (else default byte-identical). The break accepts a
        // char iff this accumulator (incl the char) ≤ available_tw — greedy first-fit
        // with punct-only demand compression folded into the fit width. See
        // session471 finding + workflow wtvi6fvix.
        let mut current_capw_tw: i32 = pt_to_tw(first_line_indent);
        let mut compress_used = false; // true after compression-based overflow absorption
                                       // S885: the most recent RIGHT/CENTER tab on the current line —
                                       // (stop_rel, cw_before_jump, cw_after_jump, alignment, lines.len()).
                                       // Content under it ends AT the stop (right) / centered on it, so the
                                       // NEXT tab must resolve from the aligned end, not the provisional
                                       // left-advance (stop + content width). Line identity via lines.len().
        let mut rc_prev_tab: Option<(f32, f32, f32, TabStopAlignment, usize)> = None;
        // S243 (2026-05-24): removed dead variable `current_grid_extra`
        // (assigned/incremented in 8 sites but never read).

        // Word buffer spans across fragment boundaries so that a single word
        // split across two runs (e.g. "te" in Run1 + "st" in Run2) is kept
        // together for line-break decisions.
        let mut word = String::new();
        let mut word_width: f32 = 0.0;
        let mut word_first_width_tw: i32 = 0;
        let mut word_natural_width: f32 = 0.0; // 2-pass wrap: natural (pre-compression) width
                                               // S809 (2026-07-13, default ON, opt-out OXI_S809_DISABLE): Latin
                                               // trailing-punctuation hang (overflowPunct). LEGACY (compat<=14)
                                               // Word lets a line-end '.' or ',' hang COMPLETELY past the content
                                               // right — DERIVED via _pb_punct_gen.py (right-margin sweep, ctrl
                                               // 'x' vs '.' vs ',': the legacy dot/comma flip boundary == the word
                                               // width WITHOUT the punct, exact to 0.1pt; the compat15 variant
                                               // flips at word+punct = NO hang) + render-truth both ways:
                                               // uklocalspending (compat 14) fits 'mandated' AT the margin with
                                               // the '.' hanging to 773.0; uk_health_form (compat 15) WRAPS
                                               // 'claims.' even though it fits to 524.01 <= edge 524.06 (Word
                                               // measures the full stop). The compat<15 gate is the Latin analog
                                               // of the S568/S572 legacy CJK oikomi discriminator. Tracks the
                                               // width of the word's trailing hangable punct; the flush fit tests
                                               // credit it. Latin-doc scope (JP keeps its own CJK burasage).
                                               // S929 (2026-07-18, default ON, opt-out OXI_S929_DISABLE): an
                                               // effective pBdr RIGHT border suppresses the legacy trailing-punct
                                               // hang — the border fences the text edge, so the full word+punct
                                               // must fit inside it. DERIVED via a faithful transplant probe of
                                               // 002c1ffa65f3a566's BoxPara ((b) …for the day.) + variant matrix
                                               // (96 Word renders, right-margin flip sweep): faith (pBdr all
                                               // sides) flips at the word+punct width = NO hang, reproducing the
                                               // real doc's wrap; nobdr / bare / pBdr-minus-RIGHT all flip at the
                                               // word-sans-punct width = FULL hang (the S809 rule). The RIGHT
                                               // side alone discriminates; tabs / hanging indent / right-tab
                                               // marker are irrelevant (notab ≡ faith). An explicit
                                               // w:val="none"/"nil" right border (the S482 sentinel) is no fence.
        let s929_right_fence = para_style
            .borders
            .as_ref()
            .and_then(|b| b.right.as_ref())
            .map_or(false, |d| d.style != "none")
            && std::env::var("OXI_S929_DISABLE").is_err();
        // S1262 (2026-08-30, default ON, opt-out OXI_S1262_DISABLE): the hang
        // also applies at compatibilityMode 15 when the paragraph is JUSTIFIED,
        // and the mark set includes the closing quotes.
        // WORD TRUTH (`tools/metrics/_pb_hangpunct_{gen,read}.py`, 156 arms =
        // 13 marks x 3 filler lengths x {left, both} x {cm14, cm15}, Courier New
        // so every glyph is 6.00pt and the column boundary is exact):
        //     compat  jc      period comma rquote apos    everything else
        //       14    left    hangs                       wraps
        //       14    both    hangs                       wraps
        //       15    left    WRAPS                       wraps
        //       15    both    hangs                       wraps
        // i.e. the legacy modes hang regardless of alignment (what S809 already
        // did) and cm15 keeps the hang only for justified text. semicolon,
        // colon, bang, query, rparen, rbracket, hyphen, a letter control and a
        // no-mark control all behave as expected in every arm, so the
        // discriminator is real.
        // ★`compat_mode` READS 15 when the document declares nothing, so the
        // old `< 15` test silently excluded every compat-less document. The
        // documented idiom for "legacy" is `compat_mode <= 14 ||
        // !compat_mode_explicit` (ir/types.rs) -- `creative__009790431a821d2f`
        // declares no compatibilityMode at all, and Word hangs its periods.
        // WITNESS `creative__009790431a821d2f` (EN Phase-1 FAIL 0.9680, with no
        // tables / images / footnotes / columns to confound): Word fits
        // `... Bankers Ghana.` by hanging the period at 523.38..526.17 past the
        // 523.32 column edge; Oxi counted it, broke a word early, spent an extra
        // line and pushed 4 paragraphs onto the next page.
        let s1262 = std::env::var("OXI_S1262_DISABLE").is_err();
        let s809_legacy = self.compat_mode < 15
            || (s1262 && !self.compat_mode_explicit);
        let s809_hang = !self.doc_body_has_real_cjk
            && (s809_legacy || (s1262 && is_justified))
            && !s929_right_fence
            && std::env::var("OXI_S809_DISABLE").is_err();
        if std::env::var("OXI_DBG_HANG").is_ok() {
            eprintln!("[HANG] s809={} legacy={} just={} cm={} expl={} cjk={} fence={}",
                s809_hang, s809_legacy, is_justified, self.compat_mode,
                self.compat_mode_explicit, self.doc_body_has_real_cjk, s929_right_fence);
        }
        // Mixed CJK/Latin justified body text can place a final ASCII period
        // beyond the content edge too. Keep that terminal glyph advance separate
        // from the earlier punctuation compression pool: spending it twice would
        // accept words that Word moves to the next line. This branch deliberately
        // leaves commas and quotes under their existing Latin-only policy.
        let cjk_latin_period_hang = self.doc_body_has_real_cjk
            && self.compat_mode >= 15 && self.compat_mode_explicit
            && is_justified && s476_body && !vertical
            && grid_char_pitch.is_none() && !s929_right_fence;
        let mut word_trail_hang_w: f32 = 0.0;
        // S245 (2026-05-24): removed dead variable `word_grid_extra`
        // (assigned/incremented at 3 sites but never read after S243
        // removed `current_grid_extra`).
        let mut word_style: Option<RunStyle> = None;
        let mut word_field_type: Option<FieldType> = None;
        let mut word_run_index: usize = 0;
        let mut word_char_offset: usize = 0;

        // LATIN-WORDWRAP (default ON, opt-out OXI_LATIN_WORDWRAP_DISABLE): Western
        // word-wrap for long Latin tokens (URLs). Word treats a maximal Latin run (e.g.
        // «https://www.mhlw.go.jp/…») as ONE word — it does NOT break at the internal
        // «/»«-»«:» opportunities to fill a PARTIAL line; the whole token wraps to the
        // next line first, and breaks at those opportunities ONLY when it overflows a
        // FULL line. Oxi's is_break_after (below) flushed at every «/»«-»«:» → it greedily
        // packed the URL start onto the preceding CJK line → 1 line too few (tokyoshugyo
        // −31, the «（ＰＤＦ版のＵＲＬ：https…» blocks). When on, the is_break_after chars do
        // NOT flush — they record a break OPPORTUNITY (char-count, cum-width) so flush_word
        // can split an over-long token across lines, with the EXACT per-«/» fragmentation
        // (run_idx/char_offset via word_seg_meta) preserved so fitting tokens are
        // byte-identical to the old greedy path. KINSOKU-gated (see flush_word) so a token
        // after a line-end-prohibited opening bracket «（» is NOT word-wrapped (Word keeps
        // the bracket with the token; wrapping would orphan it — c7b923). Gate: tokyoshugyo
        // 0.9746→0.9753 (+1 para), full corpus 81/84 unchanged (only tokyoshugyo's
        // pagination moves), SSIM 0 regressed (all previously-changed word_png docs
        // byte-identical via the seg-meta + kinsoku gate). See [[char_budget_wall]].
        let latin_wordwrap = std::env::var("OXI_LATIN_WORDWRAP_DISABLE").is_err();
        // A Latin-only paragraph can break a compound at its hyphen even
        // when other paragraphs contain CJK text. Keep URL tokens together.
        let legacy_latin_hyphen = std::env::var_os("OXI_LEGACY_LATIN_HYPHEN").is_some()
            && self.compat_mode_explicit
            && self.compat_mode <= 14
            && self.doc_body_has_real_cjk
            && s476_body
            && !vertical
            && !fragments.iter().any(|f| f.0.chars().any(kinsoku::is_cjk));
        // S1100 (2026-08-08, default ON, opt-out OXI_S1100_DISABLE): an EM DASH
        // (U+2014) / EN DASH (U+2013) carries a break OPPORTUNITY AFTER it.
        // DERIVED (tools/metrics/_pb_emdash_gen.py, 76 arms = 4 shapes x 19
        // right-indent steps sweeping the dash across the margin, Word PDF):
        //   NB (NBSP-joined, the target's shape «1300 mm —400 mm»)
        //     right 3000tw  line1 = «… 1300 mm —400 mm»   whole token fits
        //     right 3100..3800 line1 = «…t over 1300 mm —»  ★break AFTER the dash
        //     right 3900+   line1 = «…ww wwww not over»    dash itself does not fit
        //   GL («1300mm—400mm», no spaces at all) 3300..3900 = «…1300mm—» — same
        //   EN (U+2013) 3200..3900 = «…1300 mm –» — same
        // Not ONE arm of the 76 breaks BEFORE the dash, so the after-only
        // opportunity explains the whole sweep (UAX #14 puts U+2014 in class B2
        // = both sides, but Word's observable behaviour needs only the after).
        // ★This is the OPPORTUNITY model (word_breaks), NOT the flush model: the
        // maximal token is kept together and split only when it overflows, which
        // is exactly what the sweep shows. The S1044 note records an
        // "after-break-CHARACTER" attempt that was reverted for breaking
        // legal__0001482d wi=1030 — that put the dash in `is_break_after`, whose
        // non-latin_wordwrap path FLUSHES at the dash.
        // policies__00148f8d p68: «…not over 1300\u{a0}mm\u{a0}—400\u{a0}mm…» is one
        // unbreakable token for Oxi (cur 4247 + word 2007 = 6254 > avail 5481) so
        // the whole thing wrapped; Word ends line 1 at «…1300 mm —» (x1 474.10 <=
        // content right 475.05). Latin-doc scope, shared with S801's own
        // classification of these dashes → JP byte-identical by construction.
        let s1100_dash_break = std::env::var("OXI_S1100_DISABLE").is_err();
        // KERNBREAK space-compression credit (2026-07-07, ★default ON,
        // opt-out OXI_KERNBREAK_DISABLE):
        // Word's JUSTIFIED Latin fit test allows squeezing word spaces below
        // natural — db9ca render-truth: 16/73 full lines have natural(em+kern)
        // width EXCEEDING the column (max +0.581/space = 22% of the 2.625
        // TNR space; min word space ≈ 0.78×natural). The overflow tests below
        // extend `available` by the accumulated per-space credit
        // (cap × space_width, OXI_KERNBREAK_CAP default 0.25); reset with the
        // line. The Latin analog of the CJK 約物 oikomi capacity credit.
        let kernbreak_para = std::env::var("OXI_KERNBREAK_DISABLE").is_err()
            // Explicit legacy compatibility uses its own proportional/monospace
            // capacity rules; enabling kerning does not enable modern space shrink.
            && !(self.compat_mode_explicit && self.compat_mode <= 14)
            && is_justified
            && para_style
                .default_run_style
                .as_ref()
                .and_then(|rs| rs.kern)
                .map_or(false, |k| k > 0.0);
        let kernbreak_cap: f32 = std::env::var("OXI_KERNBREAK_CAP")
            .ok()
            .and_then(|v| v.parse().ok())
            .unwrap_or(0.25);
        // S799 (2026-07-12, opt-out OXI_S799_DISABLE): a JUSTIFIED Latin line fits
        // its last word by compressing SPACES below natural — Word's justify
        // shrinks spaces on demand (the Latin analog of the JP 約物 oikomi).
        // Derived on ukframework's nominate bullet (jc=both, per-space w:spacing
        // baked into the docx): natural space = em 2.487 + cs ≈ 3.46pt, Word PDF
        // renders the line's spaces at 3.12 avg (~10% shrink) and fits
        // 'shareholder' where Oxi wrapped it (needed 61tw over 9 spaces ≈
        // 6.8tw/space, well under the 25% cap). This extends the KERNBREAK
        // space-compression credit (kern-active-only) to the no-kern justified
        // case. Scope: !doc_body_has_real_cjk → JP byte-identical by construction.
        let s799_space_shrink = std::env::var("OXI_S799_DISABLE").is_err()
            && is_justified
            && !self.doc_body_has_real_cjk;
        // S953 (2026-07-20, opt-out OXI_S953_DISABLE): a COMPAT-15 JUSTIFIED
        // line hangs its trailing '.'/',' past the right edge — the S809 rule's
        // modern-justified arm. DERIVED (hang_probe, Arial-12, 12-space line,
        // 2tw margin sweeps against the RENDER-measured natural 444.14):
        //   left-aligned flip  = natural EXACT (no shrink, no hang — matches
        //                        the S809 compat15 control);
        //   justified flip B* = (natural − w('.')) − quarter_cap  ⇒ the '.'
        //                        HANGS on top of the S825 quarter-space shrink.
        // Real doc 0018d5f3 (blind knife-edge): «…access sick notes.» — Word
        // compresses 12 spaces by 8.54 (≤ cap 9.99) and parks the '.' at
        // 522.12→525.45 past the 522.15 margin; Oxi missed the fit by 1.5pt
        // without the hang (DBGFLUSH needed 8879tw vs avail+cred 8849). The
        // first probe round mis-read "no hang" because a hand-summed natural
        // was −3.25 off (Oxi's own width sum matches Word within 0.19) — the
        // measured-natural lesson (S825b) again.
        // S1018 (2026-07-26, opt-out OXI_S1018_DISABLE): S953's trailing
        // '.'/',' hang does NOT apply to a DECIMAL close-paren numbered list
        // item (`1)`, `5)`, …). reports__0013aa1d p4 «5) TAS North … benchmark
        // of 85%.»: Word wraps to 3 lines, but Oxi's S953 hangs the terminal
        // '.' and over-packs `85%.` onto line 2 (natural 8543 > avail+S825 8516
        // by 27tw, closed by the full period width) → the empty para + Figure 8
        // shift a page → −1×2. S825 space-compression stays; only the S953
        // over-hang is removed. Census (compat15 + justified + resolved decimal-
        // close-paren marker): golden 0 / real_en 45-paras-in-target / JP 0.
        // The legacy compat<15 s809_hang arm is a separate rule — untouched.
        let s1018_decimal_paren_list = std::env::var("OXI_S1018_DISABLE").is_err()
            && para_style.list_marker.as_deref().map_or(false, |m| {
                m.trim().strip_suffix(')').map_or(false, |digits| {
                    !digits.is_empty() && digits.chars().all(|c| c.is_ascii_digit())
                })
            });
        // The modern justified arm of S1262 must retain the numbered-list
        // exclusion too; only the legacy arm has independent hanging rules.
        let s953_hang = (s809_hang && (s809_legacy || !s1018_decimal_paren_list))
            || (!self.doc_body_has_real_cjk
                && self.compat_mode >= 15
                && self.compat_mode_explicit
                && s799_space_shrink
                && !s929_right_fence
                && !s1018_decimal_paren_list
                && std::env::var("OXI_S953_DISABLE").is_err());
        let s809_hang = s953_hang;
        // S1630 (2026-10-01, OPT-IN OXI_S1630=1 -- held, see the end): in the
        // compat-15 justified arm only the '.' hangs; S953 measured a period
        // (Arial-12 «notes.») and took the comma along unmeasured.
        // `_pb_jline_slice_gen.py` (reports__0013bcb8's own justified line, Book
        // Antiqua 8pt, ending «…semper vel,»): Word keeps the line down to
        // 218.95pt at w:w 105 and 208.70pt at 100%, i.e. the plain S825/S1475
        // shrink with the comma COUNTED (excess 4.88 / 4.48 against the 4.73 /
        // 4.50 allowance). With the comma hanging Oxi kept «vel,» on a 218.25pt
        // column where Word wraps it, which put the right column two lines of
        // text ahead and lifted the p2 table 20pt. (The same line ending in '.'
        // does not hang either -- the period's own discriminator is open; it
        // keeps S953's behaviour here.)
        // HELD: legal__0027c9c1 (Arial 11, «…October 11,») needs the old
        // behaviour -- Word shrinks that line 7.6pt where the S825/S1475
        // allowance gives 6.42, and its comma and period variants also flip at
        // the same width. The comma hang was standing in for an allowance law
        // that is still wrong for some lines; that law is the real fix.
        let s1630_no_comma = std::env::var("OXI_S1630").as_deref() == Ok("1") && !s809_legacy;
        // S799 cap: the no-kern justified shrink is SMALLER than KERNBREAK's 0.25
        // (a blanket 0.25 over-fits — framework {−1:20}→{−1:31}); sweep knob.
        // S1475: the last-word ceiling applies to the S825 (compat-15 explicit)
        // arm only — the c14 / S933 / S1046 classes keep their own allowances.
        let s1475_on = std::env::var("OXI_S1475_DISABLE").is_err()
            && s799_space_shrink
            && self.compat_mode >= 15
            && self.compat_mode_explicit;
        let mut s1475_space_tw: i32 = 0;
        let s799_cap: f32 = std::env::var("OXI_S799_CAP")
            .ok()
            .and_then(|v| v.parse().ok())
            .unwrap_or(0.10);
        // S994 (2026-07-23, opt-out OXI_S994_DISABLE): w:wpJustification widens a
        // fully-justified line's fit budget to W_actual × (1 + 281/7200) (MS-OE376
        // §2.1.481) — fitting ~one more word per line. The effective COLUMN width
        // widens uniformly, so the credit is `available_tw × 281/7200` on EVERY line
        // (Model A / "paragraph-width"): the first-line indent is a start offset, not
        // a column-width reduction. This is both physically correct AND the empirical
        // winner over the 0011dcc/0011b198 population (Model B, which subtracts the
        // indent on line 0, under-fits the first line). OXI_WPJ_MODEL=B for the
        // post-indent alternative. The matching render inter-word compression is at
        // the justify site.
        let wpj_active =
            self.wp_justification && is_justified && std::env::var("OXI_S994_DISABLE").is_err();
        let wpj_model_a = std::env::var("OXI_WPJ_MODEL")
            .map(|v| v != "B")
            .unwrap_or(true);
        let wpj_first_indent_tw = pt_to_tw(first_line_indent);
        let wpj_credit_at = |nlines: usize| -> i32 {
            if !wpj_active {
                return 0;
            }
            let content_tw = if nlines == 0 && !wpj_model_a {
                (available_tw - wpj_first_indent_tw).max(0)
            } else {
                available_tw
            };
            ((content_tw as f32) * 281.0 / 7200.0).round() as i32
        };
        let mut latin_space_credit_tw: i32 = 0;
        let mut latin_space_credit_remainder: f32 = 0.0;
        // S995/S996 / C14JUST (2026-07-24, default ON, opt-out OXI_C14JUST_DISABLE):
        // the probe-derived compat14 legacy Latin justify rule. A justified line KEEPS
        // a candidate word (compressing inter-word SPACES to the half-em floor) iff the
        // whole line FITS within its compression CAPACITY:
        //    natural(line_body + space + candidate) − available ≤ n_spaces × half_em
        // i.e. the fit budget = latin_space_credit_tw (= Σ half-em per accumulated
        // space) with NO alt_slack gate and NO trailing-punct hang. (S996 correction:
        // S995 shipped `alt_slack>1.49×space ? capacity : 0 + hang`, which mis-predicts
        // BOTH ways — the alt_slack gate zeroes the credit for low-alt_slack candidates
        // [213 'law.'/584 'or' wrap, should KEEP] and the +hang over-credits wide ones
        // [581 'partnership.'/twin147 'programs.' keep, should WRAP]. _pb_c14floor
        // pinned: floor C≈3.556pt≈half-em, candidate-width-INDEPENDENT, period-hang=0
        // ['final.'≡'finalx' flip], double-space = 2 space-chars. _pb_c14down FALSIFIED
        // line-ordinal. The report's "same-excess reversals" [581/660, 147/105] were
        // rendered-width measurement artifacts — 660's natural overflow is −71.9pt
        // [trivial keep], NOT the 43.2pt of 581; the capacity rule has no reversal and
        // predicts all 4 S995 errors + the 2 counterexamples correctly [DELIVERABLE].)
        // SUPERSEDES S933 fs/4 + S994 wpj_credit for c14 (do NOT add them). MONOSPACE-
        // scoped at the accumulation (char i-width == M-width): the half-em floor was
        // DERIVED on Courier; proportional c14 docs keep their S933 fs/4 allowance
        // [reference__0029c1c/technical__00549a8f]. c14_space_tw>0 (monospace only)
        // gates the fit tests.
        let c14_active = std::env::var("OXI_C14JUST_DISABLE").is_err()
            && self.compat_mode == 14
            && self.compat_mode_explicit
            && is_justified
            && !self.doc_body_has_real_cjk;
        let mut c14_space_tw: i32 = 0; // natural space width (captured at accumulation)
                                       // S774 (2026-07-10, rides the TABTW/Latin scope): a RIGHT-aligned tab
                                       // pins the following segment's END at the tab stop — the segment grows
                                       // LEFTWARD into the gap the tab jumped, so the wrap check must credit
                                       // that slack (hmrc's title «[crown]⇥Starter Checklist», right tab at
                                       // 521.55: Word renders ONE line with the 24pt text right-aligned; Oxi
                                       // treated the stop as a left anchor → tab + 190pt text overflowed →
                                       // 2 lines). The tab-width post-process below already shrinks the tab
                                       // fragment by the segment width, so allowing the segment through here
                                       // renders exactly right-aligned. Reset with the other per-line
                                       // accumulators and on the next tab.
        let mut right_tab_slack_tw: i32 = 0;
        // S958 (2026-07-20): the tw position of the last CENTER tab stop on this
        // line. A center tab's following segment straddles the stop — its right
        // edge is stop + S/2, not stop + S — so the overflow track (which only
        // grows rightward) over-counts by S/2. policies__00148f8d's footer
        // disclaimer (Arial 8pt, leading <w:tab/> at a 3600tw center stop) is ONE
        // line in Word (PDF: x 172.94..429.99, centre 301.5 = the stop) but Oxi
        // broke it in two, inflating the footer stack by 9.2pt and pulling the
        // body bottom from 639.3 to 630.10 -> a 4-line paragraph widow-split.
        let mut center_tab_stop_tw: Option<(i32, i32)> = None;
        let mut word_breaks: Vec<(usize, f32)> = Vec::new();
        // S1059 (2026-08-02): cumulative width after each char of the pending
        // token, so an over-long SEGMENT can be split at the character level.
        // Kept parallel to `word` rather than pushed into `word_breaks` — a
        // bound there becomes its own LineFragment (that is what S745 accepts
        // for wordWrap=0), and per-char fragments would change DWrite shaping
        // for every Latin document.
        let mut word_char_ws: Vec<f32> = Vec::new();
        // Per-SEGMENT (run_idx, char_offset) so a split token keeps the EXACT run
        // metadata of the default per-«/» fragmentation (DWrite re-derives glyph
        // positions from run_idx/char_offset; a merged multi-run token with the first
        // run's metadata mis-renders → SSIM regress). Each entry = (start char-count
        // within `word`, run_idx, char_offset). seg_pending marks "next char starts a
        // new segment" (after a break char).
        let mut word_seg_meta: Vec<(usize, usize, usize)> = Vec::new();
        let mut word_seg_styles: Vec<(usize, RunStyle)> = Vec::new();
        let mut seg_pending = false;

        // S1026 (Origin A, 2026-07-28): total / running non-whitespace char count.
        // The c14 badness "long final token" test (Part B, r_terminal=0.05) needs to
        // know if the candidate is the paragraph's FINAL non-whitespace token — true
        // iff every non-whitespace char has been consumed (pushed to `word`) by the
        // time it is flushed. Defined BEFORE the flush_word macro so the macro's
        // definition-site hygiene can see it; incremented at the two `word.push`
        // sites; read ONLY inside the c14-scoped badness → non-c14 byte-identical.
        let s1026_total_nonws: usize = fragments
            .iter()
            .map(|(t, _, _, _, _)| t.chars().filter(|c| !c.is_whitespace()).count())
            .sum();
        let mut s1026_nonws_consumed: usize = 0;
        // ★S1026-REPLAY diagnostic (2026-07-28, OXI_DBG_REPLAY, default OFF =
        // byte-identical). Emits per-candidate break geometry at the decision point
        // so the analysis unit can join the 165-record first-divergence dataset to
        // actual geometry (REPORT_S1026_sequential_replay_dataset §6/§9). Read ONCE
        // before the macro (definition-site hygiene). Only body calls carry a para
        // index; the pass (0=first, 1=S721 retry) lets the join pick the final pass.
        let s1026_replay: bool = std::env::var("OXI_DBG_REPLAY").is_ok();
        let s1026_replay_para: Option<usize> = S1026_REPLAY_PARA.with(|c| c.get());
        let s1026_replay_pass: u8 = if S721_ORPHAN_RETRY.with(|f| f.get()) {
            1
        } else {
            0
        };
        let s1026_replay_on = s1026_replay && s1026_replay_para.is_some();
        // S1027 (2026-07-28, OPT-IN OXI_S1027=1, default OFF = byte-identical):
        // LEADING-SPACE-at-wrap-boundary. Word breaks a run of N U+0020 spaces by
        // consuming ONE at the break; the remaining N−1 LEAD the next line (a fixed
        // Courier cell, justify-invariant — REPORT_S1026_leadspace_rule_probe v50,
        // 32-arm Word probe: N=1→0 / 2→1 / 3→2 / 16→15, predecessor-independent).
        // Oxi collapses all N → the next line's curw is (N−1)·space short. Alone it
        // is pagination-neutral on the twins (the 17 corrected lines have slack) BUT
        // it CORRECTS the badness-geometry tuples: with it, the v46 "expressiveness
        // proof" collapses — the 2003.×7 WRAP lines lead (`2003.  Amended`), moving
        // their β bound 0.5333→1.4167, and the verbal collision leads (`REQUIRED.
        // (a)` → β<1.175) while partner./state./years./person/before don't (single
        // space) → ONE interval β ∈ [0.6375, 1.175) fits every measured specimen.
        // Pairs with OXI_S1026_W (width penalty) on the corrected runtime tuples.
        // Scope: c14 monospace (the 2 Courier twins; JP CJK-gated out).
        let s1027_on = std::env::var("OXI_S1027_DISABLE").is_err();
        macro_rules! wrap_and_seed {
            ($sty:expr) => {{
                if std::env::var("OXI_DBGWRAP").is_ok() {
                    let tail: Vec<String> = current_line.fragments.iter().rev().take(3).map(|f| f.text.clone()).collect();
                    eprintln!("[WRAP-SEED] tail={:?} nfrag={} cw_tw={} lines={}", tail, current_line.fragments.len(), current_width_tw, lines.len());
                }
                let n_trail = if s1027_on && c14_active && c14_space_tw > 0 {
                    current_line
                        .fragments
                        .iter()
                        .rev()
                        .take_while(|f| f.text == SPACE_STRING)
                        .count()
                } else {
                    0
                };
                if n_trail >= 2 {
                    let keep = current_line.fragments.len() - (n_trail - 1);
                    current_line.fragments.truncate(keep);
                }
                // S1342 (2026-09-06, default ON, opt-out OXI_S1342_DISABLE): 行末禁則 on
                // EVERY word-path wrap -- a line may not end in an opening bracket.
                // The OPENWRAP branch carried the bracket only on its own path;
                // reference__0ea3ec86 p23 「…事務所等（|33･299㌻）」 and p8 「…項目（|16
                // 項目）」 wrapped their digit tokens through another flush and left
                // the 「（」 at the line end (Word: 「…事務所等」 / 「（33･299㌻）、…」).
                let mut s1342_carried: Vec<LineFragment> = Vec::new();
                if std::env::var("OXI_S1342_DISABLE").is_err() {
                    while current_line.fragments.len() > 1
                        && current_line
                            .fragments
                            .last()
                            .and_then(|f| f.text.chars().last())
                            .map_or(false, kinsoku::is_line_end_prohibited)
                    {
                        let f = current_line.fragments.pop().unwrap();
                        current_width -= f.width;
                        current_width_tw -= pt_to_tw(f.width);
                        current_capw_tw -= pt_to_tw(f.width);
                        s1342_carried.push(f);
                    }
                }
                lines.push(std::mem::take(&mut current_line));
                current_width = 0.0;
                current_width_tw = 0;
                current_capw_tw = 0;
                latin_space_credit_tw = 0;
                    latin_space_credit_remainder = 0.0;
                right_tab_slack_tw = 0;
                center_tab_stop_tw = None;
                compress_used = false;
                for f in s1342_carried.into_iter().rev() {
                    current_width += f.width;
                    current_width_tw += pt_to_tw(f.width);
                    current_capw_tw += pt_to_tw(f.width);
                    current_line.fragments.push(f);
                }
                if n_trail >= 2 {
                    let sp_w = (c14_space_tw as f32) / 20.0;
                    let seed_tw = c14_space_tw * (n_trail as i32 - 1);
                    for _ in 0..(n_trail - 1) {
                        current_line.fragments.push(LineFragment {
                            auto_space_shrink: 0.0,
                            text: SPACE_STRING.to_owned(),
                            width: sp_w,
                            natural_width: sp_w,
                            style: $sty.clone(),
                            tab_alignment: None,
                            tab_position: None,
                            field_type: None,
                            run_index: 0,
                            char_offset: 0,
                        });
                    }
                    current_width += sp_w * (n_trail - 1) as f32;
                    current_width_tw += seed_tw;
                    current_capw_tw += seed_tw;
                }
            }};
        }
        // Helper: flush the accumulated word into current_line, breaking if needed.
        let dbg_flush: bool = std::env::var("OXI_DBGFLUSH").ok().map_or(false, |needle| {
            !needle.is_empty() && fragments.iter().any(|&(t, _, _, _, _)| t.contains(&needle))
        });
        // S1346 (2026-09-07): the at-default regime's elective budget for a LATIN
        // word that overflows the floor (set per fragment next to s1318_at_default_regime,
        // read inside flush_word).
        let s1346_regime_credit: std::cell::Cell<i32> = std::cell::Cell::new(0);
        let word_full_punctuation_credit = std::cell::Cell::new(false);
        // S1585 (2026-09-27, default ON, opt-out OXI_S1585_DISABLE): a JUSTIFIED
        // CJK line squeezes its autoSpace gaps (and mid-line ideographic spaces)
        // to keep its last unit. `_pb_cjk_jshrink_gen.py` (MS Mincho 10.5 +
        // Century, compat 15, docGrid lines, 1-twip right-indent sweeps, left vs
        // justified flips; Oxi's natural flips match Word within 0.1 in every arm):
        //   CJK last unit, G gaps          1.3 / 2.0 / 2.6 / 2.9 / 3.1 (G 1/2/4/6/8)
        //                                  = 1.5 x (fs/4) x G/(G+2)
        //   Latin last word, 2 mid U+3000   5F 4.8, 5FG 7.3, 5FGH 7.9, ABCDEF 7.9
        //                                  = min(0.25 x mid U+3000 + gaps,
        //                                        max(gaps, 0.31 x (word + gap)))
        //   Latin last word, no U+3000      1.3 whatever the word (= the gap term)
        //   pure CJK                        0.0
        // policies__1411889624's address line 「…西堀6番館ビル5F」 (natural 426.55,
        // Word keeps it at 421.75) is the second row.
        let s1585_on = std::env::var_os("OXI_S1585_DISABLE").is_none()
            && self.doc_body_has_real_cjk
            && is_justified
            && s476_body
            && !vertical
            && grid_char_pitch.is_none()
            && self.compat_mode >= 15
            && self.compat_mode_explicit;
        let s1585_counts = |chars: &[char]| -> (usize, usize) {
            let solid = |c: char| kinsoku::is_cjk(c) && c != '\u{3000}';
            let mut gaps = 0usize;
            for w in chars.windows(2) {
                let (a, b) = (w[0], w[1]);
                if (solid(a) && b.is_ascii_alphanumeric()) || (a.is_ascii_alphanumeric() && solid(b)) {
                    gaps += 1;
                }
            }
            let mut seen = false;
            let mut mid_sp = 0usize;
            let mut pending = 0usize;
            for &c in chars {
                if c == '\u{3000}' {
                    if seen { pending += 1; }
                } else if !c.is_whitespace() {
                    seen = true;
                    mid_sp += pending;
                    pending = 0;
                }
            }
            (gaps, mid_sp)
        };
        let s1585_gap_part = |g: usize, fs: f32| -> f32 {
            if g == 0 { 0.0 } else { 1.5 * (fs / 4.0) * g as f32 / (g as f32 + 2.0) }
        };
        // Legacy gap squeeze (proposal 2026-10-05, opt-out
        // OXI_LEGACY_GAP_SQUEEZE_DISABLE): a compat-14 compressPunctuation body
        // line keeps its overflowing last unit by squeezing its CJK<->Latin
        // autoSpace gaps, left-aligned or justified alike. Faithful slices of
        // legal__0f631d468773aee2 (MS Mincho + Century 12pt, 1-twip right-indent
        // sweeps, left and justified flips identical, PDF advances):
        //   each gap gives at most half of its fs/4 (G=2/4/6: 3.00/5.88/8.88);
        //   an overflowing CJK char is kept while the overflow is within
        //   min(fs/2, the gap floor) (12 gaps, no ')': 6.00);
        //   a following line-start-prohibited ')' rides on the floor alone
        //   (12 gaps with ')': 10.08; 10.5pt: 8.76 = 5.25 + 3.47).
        // doNotCompress (flip at natural) and compat 15 jc=left (flip at
        // natural) do not squeeze; compat-15 justified lines are S1585's.
        let legacy_gap_on = std::env::var_os("OXI_LEGACY_GAP_SQUEEZE_DISABLE").is_none()
            && self.compress_punctuation
            && self.compat_mode <= 14
            && s476_body
            && !vertical
            && !lines_and_chars
            && grid_char_pitch.is_none();
        let legacy_gap_floor = |chars: &[char], fs: f32| -> f32 {
            let is_lat = |c: char| {
                (c.is_ascii_alphabetic() && para_style.auto_space_de)
                    || (c.is_ascii_digit() && para_style.auto_space_dn)
            };
            let mut gaps = 0usize;
            let mut island_from_start = chars.first().map_or(false, |&c| is_lat(c));
            for w in chars.windows(2) {
                let (a, b) = (w[0], w[1]);
                let a_cjk = kinsoku::is_cjk_ideograph_or_kana(a);
                let b_cjk = kinsoku::is_cjk_ideograph_or_kana(b);
                if a_cjk && is_lat(b) {
                    gaps += 1;
                    island_from_start = false;
                } else if is_lat(a) && b_cjk {
                    if !island_from_start {
                        gaps += 1;
                    }
                    island_from_start = false;
                } else if !is_lat(b) {
                    island_from_start = false;
                }
            }
            gaps as f32 * fs / 8.0
        };
        let ideographic_closing_spacing = s476_body && is_justified && self.compress_punctuation
            && grid_char_pitch.is_none() && !vertical;
        let mut ideographic_closing_lines = std::collections::BTreeSet::new();
        let ideographic_space_capacity = |line: &Line| -> f32 {
            line.fragments.iter()
                .skip_while(|f| f.text.chars().all(char::is_whitespace))
                .filter(|f| !f.text.is_empty() && f.text.chars().all(|c| c == '\u{3000}'))
                .map(|f| (f.width - f.natural_width * 0.25).max(0.0)).sum()
        };
        macro_rules! flush_word {
            ($style:expr) => {
                if !word.is_empty() {
                    // S1346: the regime's elective half-cell for a Latin word when
                    // the line already holds a compressible mark.
                    let s1346_credit_tw: i32 = {
                        // Legacy no-character-grid Word admits the first Latin
                        // glyph at natural width before considering a whole word
                        // against punctuation capacity. Controlled Word sweeps
                        // distinguish 5pt Mincho letters from 2.22pt Arial i;
                        // modern justified Word instead admits the complete word.
                        // The first-glyph admission test is an elective word
                        // wrapping rule. A single line-start-prohibited closing
                        // character instead retains the kinsoku allowance.
                        // Word's faithful paragraph controls keep the final ')'
                        // with the preceding Japanese sentence on the same line.
                        let legacy_closing = word.chars().count() == 1
                            && word.chars().next().is_some_and(kinsoku::is_line_start_prohibited);
                        let legacy_first_fits = !word_full_punctuation_credit.get()
                            || self.compat_mode >= 15
                            || legacy_closing
                            || current_width_tw + word_first_width_tw <= available_tw;
                        let c = if legacy_first_fits { s1346_regime_credit.get() } else { 0 };
                        let latin = c > 0 && word.chars().all(|ch| (ch as u32) < 0x2E80 || ch as u32 == 0xFFE5);
                        // elective halves on the line: a lone mark gives one, a run of k
                        // adjacent marks k - 1 (the first member of a pair is natural)
                        let mut halves = 0i32;
                        if latin {
                            let mut run = 0i32;
                            for ch in current_line.fragments.iter().flat_map(|f| f.text.chars()) {
                                let m = matches!(ch, '、' | '。' | '，' | '．' | '・' | '：' | '；')
                                    || kinsoku::is_yakumono_opening(ch)
                                    || kinsoku::is_yakumono_closing(ch);
                                if m {
                                    run += 1;
                                } else {
                                    halves += if run == 1 { 1 } else { (run - 1).max(0) };
                                    run = 0;
                                }
                            }
                            halves += if run == 1 { 1 } else { (run - 1).max(0) };
                            // S1346: a line-initial lone opening bracket offers no blank
                            let first_two: Vec<char> = current_line.fragments.iter().flat_map(|f| f.text.chars()).take(2).collect();
                            if first_two.first().map_or(false, |&c| kinsoku::is_yakumono_opening(c))
                                && !first_two.get(1).map_or(false, |&ch| {
                                    matches!(ch, '、' | '。' | '，' | '．' | '・' | '：' | '；')
                                        || kinsoku::is_yakumono_opening(ch)
                                        || kinsoku::is_yakumono_closing(ch)
                                })
                            {
                                halves = (halves - 1).max(0);
                            }
                        }
                        // S1346: after an opening bracket the word may spend every blank
                        // (「…各島支庁（303」 1.5 from 、、（); otherwise half a cell
                        let open_before = latin
                            && current_line
                                .fragments
                                .last()
                                .and_then(|f| f.text.chars().last())
                                .map_or(false, kinsoku::is_yakumono_opening);
                        if c > 0 {
                            (if open_before || word_full_punctuation_credit.get() { halves * c } else if halves > 0 { c } else { 0 }) + 2
                        } else {
                            0
                        }
                    };
                    let s1346_credit_tw = s1346_credit_tw.max(if s1585_on
                        && word.chars().all(|c| c.is_ascii_alphanumeric())
                    {
                        let fs = $style.font_size.unwrap_or(self.default_font_size);
                        let line_chars: Vec<char> = current_line.fragments.iter().flat_map(|f| f.text.chars()).collect();
                        let (g0, mid_sp) = s1585_counts(&line_chars);
                        let gap_before = line_chars.last().map_or(false, |&c| kinsoku::is_cjk(c) && c != '\u{3000}');
                        let g = g0 + gap_before as usize;
                        let gap_part = s1585_gap_part(g, fs);
                        let sp_cap = 0.25 * mid_sp as f32 * fs + gap_part;
                        let last = 0.31 * (word_width + if gap_before { fs / 4.0 } else { 0.0 });
                        pt_to_tw(sp_cap.min(gap_part.max(last)))
                    } else { 0 });
                    let s1346_credit_tw = s1346_credit_tw.max(if legacy_gap_on
                        && IN_TABLE_LAYOUT.with(|c| c.get()) == 0
                        && word.chars().count() == 1
                        && word.chars().next().is_some_and(kinsoku::is_line_start_prohibited)
                    {
                        let fs = $style.font_size.unwrap_or(self.default_font_size);
                        let line_chars: Vec<char> = current_line.fragments.iter()
                            .flat_map(|f| f.text.chars()).chain(word.chars()).collect();
                        pt_to_tw(legacy_gap_floor(&line_chars, fs))
                    } else { 0 });
                    if dbg_flush {
                        eprintln!("[DBGFLUSH] word={:?} w_tw={} cur_tw={} curw_f={:.4} ww_f={:.4} avail={} spcred={} tabslack={} line_n={} just={} s799={} s1346+1585={} s1585_on={}",
                            word, pt_to_tw(word_width), current_width_tw, current_width, word_width, available_tw,
                            latin_space_credit_tw, right_tab_slack_tw, lines.len(),
                            is_justified, s799_space_shrink, s1346_credit_tw, s1585_on);
                    }
                    let ws = word_style.take().unwrap_or_else(|| $style.clone());
                    let wft = word_field_type.take();
                    // COM-confirmed (2026-04-14): charGrid extra does NOT affect line
                    // break. Word wraps based on natural char widths (fontSize for
                    // fullwidth, smaller for halfwidth). Grid extra only affects
                    // character positioning within the line, not line break count.
                    // S1061 (2026-08-02, held) → S1061b (2026-08-15, SHIPPED default-ON
                    // for EXPLICIT compat15, opt-out OXI_S1061_DISABLE, force-everywhere
                    // OXI_S1061=1 kept for probing): the fit test rounds the PENDING WORD
                    // to twips and adds it to an accumulator that was itself rounded per
                    // word, so a line measures up to ~2tw narrower than its true advance
                    // and Oxi keeps a word Word wraps. Word's rule is "the line content
                    // EXCLUDING the trailing space fits (text_right - ind_right)" at
                    // EXACT design-metric sums — scratchpad/rightedge (42 arms) +
                    // policies__003ccc95 'owners' knife-edge (exact 8560.23 > budget
                    // 8560 → Word wraps; per-word-rounded 8558 kept it → the doc's -2
                    // page drift). ★compat-SCOPED, which dissolves the old HELD blocker:
                    // Word's compat15 layout runs on exact design metrics (hmtx/upm
                    // sums), while compat14 legacy layout runs on QUANTIZED advances —
                    // Courier New 12pt breaks at 7.2pt/char exactly (= 600/1000em),
                    // NOT the font's true 1229/2048 = 7.20117 (dcc PDF advances are
                    // flat 7.2; capacity-exact lines 9936=9936 KEEP, which exact sums
                    // would wrap). The legacy per-word rounding ≡ the quantized model
                    // on the c14 twins (144tw/char for words < 22 chars), so c14 keeps
                    // it and the S1028 windows stay calibrated — no re-derivation
                    // needed. Blast radius (S1125-era A/B): 5 compat15 docs improve
                    // (correspondence__001b7c 0.9857→1.0000 FAIL→PASS, 001aad pcd+1→0,
                    // creative__0158c +0.06, reference__0052ba +0.003, policies pcd
                    // −2→−1), the 2 compat14 legal twins are the only regressions —
                    // exactly the scope split.
                    // ★CEIL, not round: Word wraps on ANY exact excess. ukhealthform
                    // 'claims.' measures 9063.398tw against budget 9063 — Word wraps,
                    // but pt_to_tw's round() collapsed it to 9063 and kept it (the JP-
                    // gate PASS→FAIL); 'owners' 8560.23 vs 8560 is the same shape. The
                    // 0.01tw epsilon only absorbs f32 accumulation noise (≤ ~1e-3tw);
                    // an exactly-full line (integer-tw content, the c14 shape) stays
                    // KEEP by construction.
                    // ★JUSTIFIED lines keep the legacy rounded accumulator (opt-out
                    // OXI_S1061_JC_DISABLE to test the blanket form). Same shape as the
                    // compat14 exclusion: a justified line's fit budget carries the
                    // S799/S994/S1028 space-compression credits, and those were
                    // calibrated ON the rounded accumulator, which drifts ~+0.5tw per
                    // word above the true sum (db9ca line 2: rounded 8232 vs exact
                    // 8223 at 8 words). Making the content exact without re-deriving
                    // the credits lets a justified line keep a word Word wraps —
                    // db9ca__20241122 «(This also applies if “This Content” is» keeps
                    // `is` where Word's own PDF wraps it (SSIM −0.0037). The two
                    // specimens that DEMAND exact (policies__003ccc95 'owners',
                    // ukhealthform 'claims.') are both LEFT-aligned, where no credit
                    // exists and the budget is the bare content width.
                    // The historical compat14 exclusion above concerned
                    // justified capacity, whose space credits remain separate.
                    // Non-justified Latin lines fit on the unrounded sum too.
                    let s1061b = (((self.compat_mode >= 15 && self.compat_mode_explicit)
                        || legacy_latin_exact)
                        && std::env::var("OXI_S1061_DISABLE").is_err())
                        || std::env::var("OXI_S1061").is_ok();
                    let word_width_tw = if s1061b {
                        ((current_width + word_width) * 20.0 - 0.01).ceil() as i32
                            - current_width_tw
                    } else {
                        pt_to_tw(word_width)
                    };
                    // Fit uses the unrounded compression capacity. Keep the
                    // emitted glyph advance independent of that fit-only credit.
                    let word_fit_width_tw = if s1061b {
                        ((current_width + word_width - latin_space_credit_remainder) * 20.0
                            - if std::env::var("OXI_SPACE_CREDIT_FIXED_DISABLE").is_err() { 0.0 } else { 0.01 }).ceil() as i32
                            - current_width_tw
                    } else { word_width_tw };
                    // S1022 (2026-07-27, char-budget badness breaker) — SHARED decision.
                    // The c14 monospace break is GREEDY oikomi (Word packs aggressively —
                    // 750/760 lines are oikomi; the global slack²-DP over-packs 0.96→0.35).
                    // Word wraps despite half-em capacity ONLY when the line has GENUINE
                    // ROOM (natural content < available) AND the compression to fit the
                    // next word is "worse" than the slack of wrapping: r·comp² > slack²,
                    // r = w_comp/w_slack (default 0.55, joint plateau [0.45,0.60] AFTER
                    // Origin B; calibration pair p1[fit]/(6)vary[wrap] pins r ∈ (0.04,
                    // 1.045)). KEY: when slack ≤ 0 the
                    // line is already FULL/over at natural (the oikomi region) → NEVER wrap
                    // by badness. Paired with the NBSP compression credit.
                    // ★Origin B (2026-07-27): computed HERE, before the latin_wordwrap
                    // branch, so a token with an INTERNAL Latin break opportunity (e.g. the
                    // ':' in «for:») does NOT bypass the badness — the whole-token wrap in
                    // that branch ORs this decision (it previously used capacity-only, so
                    // legal__0011b198 para 147 «…required for:» packed onto a full line Word
                    // wraps). false ⇒ byte-identical; non-c14 docs are always false.
                    // S1026 (Origin A, 2026-07-28, HELD OPT-IN OXI_S1026=1, default OFF
                    // = byte-identical). REPORT_S1022_originA_bounded_probe (Part B) +
                    // REPORT_S1026_compensation_residual (the dcc compensation set). Part
                    // B (long-final-token r=0.05) is the RIGHT model — it removes 2 real
                    // Oxi over-counts (partner./11.055.) but EXPOSES a pre-existing dcc
                    // compensation set. Stages A2 (§6.1 first-line 1-3char, slack ∈
                    // [-3sp,0]), A3 (§6.3 later-line ≥2 shallow 2-char), R4 (§7 4-char
                    // positive-slack r=0.40) fix that set: with all 4 the FULL model is
                    // dcc 0.9825→0.9839 (each partial subset is WORSE — the stages
                    // compensate). BUT 0011b198 has its OWN compensation set the dcc-
                    // derived stages don't cover → b198 0.9917→0.9876 in EVERY variant.
                    // §11 gate #5 (don't ship a stage that makes either twin worse) →
                    // held until b198's compensation residuals are co-derived. Tunes:
                    // OXI_S1026_RT (Part B r), OXI_S1026_R4 (R4 r). s1026_final_token =
                    // is this word the paragraph's FINAL non-whitespace token? (Part B).
                    let s1026_on = std::env::var("OXI_S1026").ok().as_deref() == Some("1");
                    let s1026_final_token = s1026_total_nonws > 0
                        && s1026_nonws_consumed == s1026_total_nonws;
                    let s1022_badness_wrap = c14_active && c14_space_tw > 0
                        && std::env::var("OXI_S1022_DISABLE").is_err()
                        && !current_line.fragments.is_empty()
                        && {
                            let comp =
                                (current_width_tw + word_width_tw - available_tw) as f32;
                            if comp <= 0.0 {
                                false // fits at natural width: no compression, keep
                            } else {
                                // slack if we WRAP (trailing inter-word space is dropped
                                // at the line end, add it back).
                                let slack = (available_tw - current_width_tw
                                    + c14_space_tw) as f32;
                                // ordinary accumulated compression capacity — read by
                                // BOTH Origin A parts (does the candidate still fit?).
                                let cap_fits = current_width_tw + word_fit_width_tw
                                    <= available_tw + latin_space_credit_tw;
                                if std::env::var("OXI_S1028_DISABLE").is_err() {
                                    // ★S1028 UNIFIED SLACK-WINDOW rule (2026-07-28, OPT-IN,
                                    // pairs with OXI_S1027 leading-space correction). The
                                    // full-doc Word-truth census (scratchpad/s1028_census.py:
                                    // every trace candidate labeled by Word's PDF line-start
                                    // offsets — the non-ws offset space is arm-independent)
                                    // shows the baseline's misses concentrate in the
                                    // NEGATIVE-slack region the old rule froze ("slack<=0
                                    // never wraps"): Word WRAPS short candidates down to
                                    // slack=-432 (any line ordinal — A2's line-0/2-char
                                    // limits were needless narrowings), KEEPS at slack<=-564
                                    // (A2's negative controls) and KEEPS at slack>=288
                                    // (wrapping would leave a >=1-cell hole; Word prefers
                                    // compressing the candidate in over stretching). WRAP
                                    // window: T_LO <= slack <= T_HI, defaults -510 (between
                                    // -432 and -564) and 216 (between 156 and 288).
                                    // Capacity stays the main fit test. first-divergence
                                    // count: baseline 184 paras -> 139 (positive-only form).
                                    let t_hi: f32 = std::env::var("OXI_S1028_T").ok()
                                        .and_then(|v| v.parse().ok()).unwrap_or(216.0);
                                    let t_lo: f32 = std::env::var("OXI_S1028_LO").ok()
                                        .and_then(|v| v.parse().ok()).unwrap_or(-510.0);
                                    slack <= t_hi && slack >= t_lo
                                } else if slack <= 0.0 {
                                    // ★S1026 Stage A2 (Origin A, §6.1): first-line SHORT-
                                    // candidate (1..3 monospace chars) shallow oikomi, slack
                                    // ∈ [-3 c14 spaces, 0]. Word WRAPS such a candidate on the
                                    // FIRST line when it is ≤3 spaces into the post-trailing-
                                    // space oikomi region and the accumulated capacity still
                                    // fits — the M=0 hole between the positive-slack badness
                                    // and capacity overflow. 81/81 Word-correct on both twins;
                                    // GEOMETRIC (to/by/or/is/of/be/up/on/not…), not a lexical
                                    // case. The -3*space lower bound is LOAD-BEARING (negative
                                    // controls dcc 341 / b198 252 «a» at slack -564/-712 Word
                                    // KEEPS). line-0 ONLY (§6.2: an any-line blanket is
                                    // falsified 110/118 → wraps ≠ keeps). Widens the original
                                    // exact-2char/-1space Part A (§6). ★A3 (§6.3): a LATER
                                    // line (ordinal ≥2) shallow 2-char oikomi, slack ∈
                                    // [-c14_space, 0). 7/7 Word-correct (fixes p353 «be»);
                                    // bounded, NOT permission to widen all later-line slack.
                                    // Otherwise the existing slack≤0 hard-fit (false).
                                    (s1026_on
                                        && lines.is_empty()
                                        && word_width_tw >= c14_space_tw
                                        && word_width_tw <= 3 * c14_space_tw
                                        && slack <= 0.0
                                        && slack >= -3.0 * (c14_space_tw as f32)
                                        && cap_fits)
                                    || (s1026_on
                                        && lines.len() >= 2
                                        && word_width_tw == 2 * c14_space_tw
                                        && slack < 0.0
                                        && slack >= -(c14_space_tw as f32)
                                        && cap_fits)
                                } else if std::env::var("OXI_S1026_W").ok().as_deref()
                                    == Some("1") {
                                    // ★S1026 WIDTH-PENALTY badness (REPORT_S1026_broad_per_
                                    // line_badness §7, HELD OPT-IN OXI_S1026_W=1). The scoped
                                    // B3/B4/B5 model was FALSIFIED; with proper origin
                                    // alignment (compare only the FIRST r-sensitive boundary
                                    // where Word / r=0.55 / r=0.28 agree before it) the two-
                                    // twin surface is EXACT: width ≥5 chars = 13/13 Word KEEP,
                                    // width ==2 = 11/11 WRAP (24/24, all line ordinals). The
                                    // current r=0.55 test assigns NO cost to ejecting a LONG
                                    // candidate; add it: wrap when 0.55·comp² > slack² +
                                    // β·long_excess², long_excess = max(0, width − 2·space).
                                    // 2-char → long_excess=0 = EXACTLY r=0.55 (all 11 short
                                    // wraps preserved); 5/7/11-char get a wrap penalty → KEEP.
                                    // SUBSUMES R4 + Part B (both positive-slack r adjustments)
                                    // → they are bypassed here (no double-application). β >
                                    // 0.4617 (measured lower bound on the 13 KEEP origins);
                                    // 0.50 is the TRIAL value, NOT a shipping constant — the
                                    // §9 length probe must supply the upper bound. NO line
                                    // ordinal (§10); negative-slack path (A2/A3) unchanged.
                                    let long_excess =
                                        (word_width_tw - 2 * c14_space_tw).max(0) as f32;
                                    let beta: f32 = std::env::var("OXI_S1026_BETA").ok()
                                        .and_then(|v| v.parse().ok()).unwrap_or(0.50);
                                    0.55 * comp * comp
                                        > slack * slack + beta * long_excess * long_excess
                                } else {
                                    // ★S1026 Part B (Origin A, §7): a paragraph's FINAL
                                    // token of ≥6 monospace chars uses r_terminal=0.05
                                    // (not the normal 0.55) — Word is lenient on a long
                                    // terminal token, FITTING it where r=0.55 would wrap
                                    // (years./state./partner./11.055.). 4/4 Word-correct;
                                    // NARROW (final-token + ≥6 chars + cap-fit): the broad
                                    // terminal-r blanket is falsified (would undo the
                                    // S1022b «for:» Origin B fix). r_terminal ∈ (0.0315,
                                    // 0.0816); 0.05 is mid-interval. Otherwise the normal
                                    // joint-plateau default 0.55 (S1022b).
                                    let use_terminal = s1026_on && s1026_final_token
                                        && word_width_tw >= 6 * c14_space_tw && cap_fits;
                                    // ★S1026 Stage R4 (Origin A, §7): a 4-monospace-char
                                    // candidate in the positive-slack branch uses r=0.40
                                    // (not 0.55) — Word FITS a 4-char token where r=0.55
                                    // wraps (dcc 320 «only» rcrit 0.4444 / 730 «has:», b198
                                    // 96 «land» rcrit 0.4874). A GLOBAL r=0.40 changes only
                                    // those 2 dcc paras (§7 diagnostic), so scoping to 4
                                    // chars is the narrow form; does NOT change the global r.
                                    let use_r4_4char = s1026_on
                                        && word_width_tw == 4 * c14_space_tw && cap_fits;
                                    let r: f32 = if use_terminal {
                                        std::env::var("OXI_S1026_RT").ok()
                                            .and_then(|v| v.parse().ok()).unwrap_or(0.05)
                                    } else if use_r4_4char {
                                        std::env::var("OXI_S1026_R4").ok()
                                            .and_then(|v| v.parse().ok()).unwrap_or(0.40)
                                    } else {
                                        std::env::var("OXI_S1022_R").ok()
                                            .and_then(|v| v.parse().ok()).unwrap_or(0.55)
                                    };
                                    r * comp * comp > slack * slack
                                }
                            }
                        };
                    // Compare expansion at the preceding break with compression
                    // needed to retain the next word, in addition to the capacity limit.
                    let justified_word_choice = s799_space_shrink
                        && self.compat_mode >= 15 && self.compat_mode_explicit
                        && std::env::var("OXI_JUSTIFIED_WORD_CHOICE").is_ok()
                        && right_tab_slack_tw == 0 && center_tab_stop_tw.is_none()
                        && !current_line.fragments.is_empty()
                        && {
                            let trailing_space = current_line.fragments.iter().rev()
                                .take_while(|f| !f.text.is_empty() && f.text.chars().all(|c| c == ' '))
                                .map(|f| pt_to_tw(f.width)).sum::<i32>();
                            let raw_compression = current_width_tw + word_fit_width_tw - available_tw;
                            // Hanging punctuation contributes only when ordinary
                            // inter-word compression cannot accommodate the token.
                            let preceding_space_credit = latin_space_credit_tw
                                - (trailing_space as f64 * 0.25).round() as i32;
                            let compression = raw_compression - if raw_compression > preceding_space_credit {
                                pt_to_tw(word_trail_hang_w)
                            } else { 0 };
                            let expansion = available_tw - current_width_tw + trailing_space;
                            let spaces = current_line.fragments.iter()
                                .skip_while(|f| f.text.chars().all(|c| c == ' '))
                                .map(|f| f.text.chars().filter(|c| *c == ' ').count()).sum::<usize>();
                            let trailing_count = current_line.fragments.iter().rev()
                                .take_while(|f| !f.text.is_empty() && f.text.chars().all(|c| c == ' '))
                                .map(|f| f.text.chars().count()).sum::<usize>();
                            let previous_spaces = spaces.saturating_sub(trailing_count);
                            // Wrapping drops the trailing separator, so compare
                            // adjustment per remaining inter-word space.
                            compression > 0 && expansion >= 0 && previous_spaces > 0
                                && compression as f64 * 2.0 * previous_spaces as f64
                                    > expansion as f64 * spaces as f64
                        };
                    let s1022_badness_wrap = s1022_badness_wrap || justified_word_choice;
                    // S1157 (2026-08-17, default ON, opt-out
                    // OXI_S1157_DISABLE): a token with NO internal break opportunity at
                    // all never reached the branch below, so Oxi placed it whole
                    // and let it run past the margin -- 640.5pt of token in a
                    // 510.2pt column on _pb_longtok's `plain` arms, where Word
                    // packs 97 characters and wraps the rest. Entering the branch
                    // costs nothing for such a token: `bounds` becomes the single
                    // whole-token segment below, and S1059's character packing
                    // inside the segment loop does the work. Gated on the token
                    // actually overflowing a full line, so anything that fits is
                    // untouched. Gate: _pb_longtok 24/24 against Word (the plain
                    // arms join the slashy ones at 97 characters), Phase 1 95/96
                    // with zero per-doc change, all 238 SSIM sentinel documents
                    // byte-identical. Not CJK-scoped -- an opportunity-free
                    // overlong token overflowed in Latin documents too.
                    let s1157_no_opp = latin_wordwrap
                        && word_breaks.is_empty()
                        && word_width_tw > available_tw
                        && word_char_ws.len() >= word.chars().count()
                        && std::env::var("OXI_S1157_DISABLE").is_err();
                    // Terminal punctuation is part of a lexical word, not an
                    // internal break opportunity. Let automatic hyphenation see
                    // these words while preserving long-token and mixed-run paths.
                    let automatic_lexical_word = self.auto_hyphenation
                        && !self.doc_body_has_real_cjk
                        && std::env::var("OXI_S1128_DISABLE").is_err()
                        && word_seg_styles.is_empty()
                        && word_width_tw <= available_tw
                        && hyphen::is_lexical_word(&word);
                    if latin_wordwrap && (!word_breaks.is_empty() || s1157_no_opp)
                        && !automatic_lexical_word {

                        // LATIN-WORDWRAP split: the token is a maximal Latin run with
                        // internal break opportunities. (1) wrap the WHOLE token to a
                        // fresh line if it doesn't fit on the current one; (2) split it
                        // across lines at the recorded opportunities only when it still
                        // overflows a full line.
                        // KINSOKU GATE: do NOT do the word-level wrap (1) when the token
                        // is immediately preceded by a line-end-PROHIBITED char (an opening
                        // bracket «（»「『 …). Word keeps that bracket WITH the token by
                        // breaking BEFORE the bracket — moving only the token would orphan
                        // the bracket at the line end (c7b923 «…ライセンス（https://…» = Word
                        // puts «（https://…» together on the next line; tokyoshugyo's URL is
                        // preceded by «：», not prohibited, so it DOES wrap). Skipping (1)
                        // here falls back to the greedy per-segment placement = the default
                        // (byte-identical), avoiding the bracket orphan.
                        let preceded_by_open = current_line.fragments.last()
                            .and_then(|f| f.text.chars().last())
                            .map_or(false, kinsoku::is_line_end_prohibited);
                        if std::env::var("OXI_DBGOPEN").is_ok() && (preceded_by_open || word.starts_with('3')) {
                            eprintln!("[OPENWRAP] word={:?} nfrag={} last_frag={:?} cw={} ww={} avail={}",
                                word, current_line.fragments.len(),
                                current_line.fragments.last().map(|f| f.text.clone()),
                                current_width_tw, word_width_tw, available_tw);
                        }
                        // S745: wordWrap=0 — skip the whole-token wrap (1); the
                        // per-char opportunities recorded above make the segment
                        // loop below pack the line to the last fitting char.
                        let s745_char_wrap = !para_style.word_wrap
                            && std::env::var("OXI_S745_DISABLE").is_err();
                        // ★Origin B: OR the SHARED S1022 badness into the whole-token
                        // wrap — a c14 token with an internal Latin break (e.g. «for:»)
                        // must not bypass the badness. false for non-c14 → byte-identical.
                        if s1026_replay_on && c14_active {
                            let cn = word.chars().filter(|c| !c.is_whitespace()).count();
                            let cr = if c14_active && c14_space_tw > 0 { latin_space_credit_tw } else { latin_space_credit_tw + wpj_credit_at(lines.len()) };
                            let hg = (if c14_active && c14_space_tw > 0 { if !s1026_final_token && std::env::var("OXI_S1028_HG_DISABLE").is_err() { (pt_to_tw(word_trail_hang_w) - 1).max(0) } else { 0 } } else { pt_to_tw(word_trail_hang_w) });
                            let ts = right_tab_slack_tw + s958_center_slack(center_tab_stop_tw, current_width_tw);
                            let cap = current_width_tw + word_fit_width_tw > available_tw + cr + ts + hg;
                            let wr = !preceded_by_open && !s745_char_wrap && (cap || s1022_badness_wrap) && !current_line.fragments.is_empty() && !para_all_whitespace;
                            eprintln!("[S1026-REPLAY] para={} pass={} start={} end={} cand={:?} cw_tw={} curw_tw={} avail_tw={} credit_tw={} tabslack_tw={} line={} branch=latin_wordwrap badness={} cap={} dec={}",
                                s1026_replay_para.unwrap(), s1026_replay_pass, s1026_nonws_consumed.saturating_sub(cn), s1026_nonws_consumed, word, word_width_tw, current_width_tw, available_tw, cr, ts, lines.len(), s1022_badness_wrap, cap, if wr {"WRAP"} else {"KEEP"});
                        }
                        // S1172 (2026-08-18, opt-IN OXI_OPENWRAP=1, default OFF =
                        // byte-identical): the kinsoku gate above declines to wrap a
                        // token that follows an opening bracket, so the greedy segment
                        // placement below keeps the bracket AND the head of the token on
                        // this line. Word does the opposite: it moves the bracket down
                        // WITH the token. c7b923e5 p3 measured 2026-08-18 --
                        //   Word  L12 ends «…表示4.0 国際ライセンス» (40 ch) and L13 opens
                        //         «（https://creativecommons.org/…», the short L12 then
                        //         justified out by 13.6%
                        //   Oxi   L12 ends «…国際ライセンス（https» (49 ch)
                        // So the comment above describes Word correctly and the code
                        // reaches the opposite arrangement. Wrapping WITH the bracket
                        // means popping the fragments already placed for it.
                        // The census finds five documents where an opening bracket
                        // precedes a 20+ character Latin token, and two of them
                        // (uklocalspending 16 sites, usnyserda 2) are large Latin
                        // documents on the most calibrated path in the breaker -- both
                        // come through untouched, because their brackets never land on
                        // a wrap boundary. GATE: c7b923e5 lines 87/94 -> 92/94 and SSIM
                        // +0.0222 (the only one of 238 bases whose bytes move),
                        // tokyoshugyo 2087 -> 2090, uklocalspending / usnyserda / d77a
                        // unchanged, Phase 1 95/96 with every per-document score
                        // identical. Opt-out OXI_OPENWRAP_DISABLE.
                        if std::env::var("OXI_OPENWRAP_DISABLE").is_err()
                            && preceded_by_open && !s745_char_wrap
                            && (current_width_tw + word_fit_width_tw > available_tw + s1346_credit_tw + (if c14_active && c14_space_tw > 0 { latin_space_credit_tw } else { s1475_last_word_cap(latin_space_credit_tw + wpj_credit_at(lines.len()), word_fit_width_tw, s1475_space_tw, s1475_on) }) + right_tab_slack_tw + s958_center_slack(center_tab_stop_tw, current_width_tw) + (if c14_active && c14_space_tw > 0 { if !s1026_final_token && std::env::var("OXI_S1028_HG_DISABLE").is_err() { (pt_to_tw(word_trail_hang_w) - 1).max(0) } else { 0 } } else { pt_to_tw(word_trail_hang_w) }) || s1022_badness_wrap)
                            && current_line.fragments.len() > 1 && !para_all_whitespace {
                            let mut carried: Vec<LineFragment> = Vec::new();
                            while current_line.fragments.len() > 1 {
                                let prohibited = current_line
                                    .fragments
                                    .last()
                                    .and_then(|f| f.text.chars().last())
                                    .map_or(false, kinsoku::is_line_end_prohibited);
                                if !prohibited {
                                    break;
                                }
                                let f = current_line.fragments.pop().unwrap();
                                current_width -= f.width;
                                current_width_tw -= pt_to_tw(f.width);
                                current_capw_tw -= pt_to_tw(f.width);
                                carried.push(f);
                            }
                            wrap_and_seed!(ws);
                            for f in carried.into_iter().rev() {
                                current_width += f.width;
                                current_width_tw += pt_to_tw(f.width);
                                current_capw_tw += pt_to_tw(f.width);
                                current_line.fragments.push(f);
                            }
                        }
                        if !preceded_by_open && !s745_char_wrap
                            && (current_width_tw + word_fit_width_tw > available_tw + s1346_credit_tw + (if c14_active && c14_space_tw > 0 { latin_space_credit_tw } else { s1475_last_word_cap(latin_space_credit_tw + wpj_credit_at(lines.len()), word_fit_width_tw, s1475_space_tw, s1475_on) }) + right_tab_slack_tw + s958_center_slack(center_tab_stop_tw, current_width_tw) + (if c14_active && c14_space_tw > 0 { if !s1026_final_token && std::env::var("OXI_S1028_HG_DISABLE").is_err() { (pt_to_tw(word_trail_hang_w) - 1).max(0) } else { 0 } } else { pt_to_tw(word_trail_hang_w) }) || s1022_badness_wrap)
                            && !current_line.fragments.is_empty() && !para_all_whitespace {
                            wrap_and_seed!(ws);
                        }
                        // Place the token as SEGMENTS split at the recorded break
                        // opportunities — IDENTICAL fragmentation to the default path (which
                        // flushes a fragment at each is_break_after char) so fitting tokens
                        // render byte-identically. The ONLY behavioural change: the WHOLE
                        // token first-fit above (Western word-wrap) — a too-long token is
                        // wrapped fresh rather than packed onto the current line, and its
                        // segments break a line only on genuine overflow.
                        let wchars: Vec<char> = word.chars().collect();
                        let total_chars = wchars.len();
                        let mut bounds = std::mem::take(&mut word_breaks);
                        if bounds.last().map_or(true, |&(cc, _)| cc < total_chars) {
                            bounds.push((total_chars, word_width));
                        }
                        let seg_meta = std::mem::take(&mut word_seg_meta);
                        let mut seg_start = 0usize;
                        let mut seg_start_w = 0.0f32;
                        // S1059: a token LONGER THAN A FULL LINE packs character by
                        // character — wrapping a segment to a fresh line cannot make it
                        // fit, and Word does not do it (probe `SL000`: line 2 runs to
                        // 521.43 inside a c-run; policies__000f7115: Word's line 1 ends
                        // at «…/1856/travel_», 522.1 against a 523.3 margin, where Oxi
                        // moved the whole «travel_…send/» segment down and left 36pt
                        // empty). Shorter tokens keep the per-segment wrap.
                        // S1156 (2026-08-17, default ON, opt-out
                        // OXI_S1156_DISABLE): the same rule holds in a CJK
                        // document -- the Latin-only gate was scope, not spec.
                        // `_pb_longtok_gen.py` (24 arms, Word PDF, the faithful
                        // w:hAnsi=ＭＳ 明朝 of tokyoshugyo's e-gov URL) has Word
                        // packing a 120-char token to 97 chars per line — and to
                        // the SAME 97 whether or not the token carries '/' every
                        // ten characters, i.e. Word fills to the margin and does
                        // not backtrack to the slash. Oxi breaks the slashy one
                        // at 90 (the last '/'), which is where tokyoshugyo p4
                        // loses its line and starts the document's whole zig-zag.
                        // Gate: probe 90 -> 97 on all 12 slashy arms, Phase 1
                        // 95/96 with zero per-doc change, all 238 SSIM sentinel
                        // documents byte-identical (tokyoshugyo is not among
                        // them), and tokyoshugyo itself -- scored against its own
                        // Word PDF -- 0.8560 -> 0.8575 with p4 +0.0782, p5
                        // +0.0527, p11 +0.0043 and nothing worse than -0.0001.
                        // `_kojin_rowgeom.py scan` loses the +19 on p4-5 outright
                        // and p27-28 come in from -35 to -17.
                        // Still open: a token with NO break opportunity at all
                        // (the probe's `plain` arms) never reaches this branch --
                        // word_breaks is empty -- so Oxi still runs it 130pt past
                        // the margin where Word packs 97 chars.
                        let s1156_cjk = std::env::var("OXI_S1156_DISABLE").is_err();
                        let s1059_overlong = (!self.doc_body_has_real_cjk || s1156_cjk)
                            && word_char_ws.len() >= total_chars
                            && word_width_tw > available_tw
                            && std::env::var("OXI_S1059_DISABLE").is_err();
                        for &(cc, cw) in bounds.iter() {
                            if cc <= seg_start || cc > total_chars { continue; }
                            let segment_style = word_seg_styles.iter().rev()
                                .find(|(start, _)| *start <= seg_start)
                                .map(|(_, style)| style).unwrap_or(&ws);
                            let seg_w = cw - seg_start_w;
                            let seg_w_tw = pt_to_tw(seg_w);
                            // The final segment must receive the same punctuation
                            // allowance as the whole-token fit decision.
                            let segment_hang_tw = if cc == total_chars {
                                if c14_active && c14_space_tw > 0 {
                                    if !s1026_final_token && std::env::var("OXI_S1028_HG_DISABLE").is_err() {
                                        (pt_to_tw(word_trail_hang_w) - 1).max(0)
                                    } else { 0 }
                                } else { pt_to_tw(word_trail_hang_w) }
                            } else { 0 };
                            // ★S1028_HG: the whole-token decision above includes the
                            // trailing-punct hang, but this SEGMENT placement re-tests
                            // WITHOUT it — a token whose only internal opportunity is its
                            // final ','/'.' (one segment) then wraps here despite the KEEP
                            // decision («revenues,»: decision KEEP, segment wrapped).
                            // Apply the same exclusive-boundary hang to the LAST segment.
                            if current_width_tw + seg_w_tw > available_tw + s1346_credit_tw + (if c14_active && c14_space_tw > 0 { latin_space_credit_tw } else { s1475_last_word_cap(latin_space_credit_tw + wpj_credit_at(lines.len()), seg_w_tw, s1475_space_tw, s1475_on) }) + right_tab_slack_tw + s958_center_slack(center_tab_stop_tw, current_width_tw)
                                + segment_hang_tw
                                && !current_line.fragments.is_empty() && !para_all_whitespace
                                && !s1059_overlong {
                                wrap_and_seed!(ws);
                            }
                            // Exact run metadata for this segment (matches the default
                            // per-«/» fragmentation), so DWrite shapes it identically.
                            let (ridx, choff) = seg_meta.iter()
                                .find(|m| m.0 == seg_start)
                                .map(|m| (m.1, m.2))
                                .unwrap_or((word_run_index, word_char_offset + seg_start));
                            // S1059 (2026-08-02, default ON, opt-out OXI_S1059_DISABLE):
                            // a segment that still does not fit after the wrap above is
                            // split at the CHARACTER level. Word truth (probe
                            // `scratchpad/urlwrap/`, 26 arms): an over-long token is
                            // packed to the last fitting character — 13/13 split arms
                            // break mid-segment, none backtracks to the preceding '/'
                            // (`SL000` line 2 ends inside a c-run, `UN080` inside «end»),
                            // and '_' is not a break class. Oxi placed such a segment
                            // whole and let it run past the margin, costing an extra
                            // line (policies__000f7115 p5: Word 2 lines, Oxi 3).
                            // Segments that fit are untouched → byte-identical.
                            let s1059 = (!self.doc_body_has_real_cjk || s1156_cjk)
                                && word_char_ws.len() >= total_chars
                                && std::env::var("OXI_S1059_DISABLE").is_err();
                            let mut piece_start = seg_start;
                            let mut piece_start_w = seg_start_w;
                            loop {
                                let piece_w = cw - piece_start_w;
                                let piece_w_tw = pt_to_tw(piece_w);
                                let limit = available_tw
                                    + if std::env::var("OXI_SEGMENT_CREDIT_DISABLE").is_err() { s1346_credit_tw } else { 0 }
                                    + (if c14_active && c14_space_tw > 0 { latin_space_credit_tw } else { s1475_last_word_cap(latin_space_credit_tw + wpj_credit_at(lines.len()), piece_w_tw, s1475_space_tw, s1475_on) })
                                    + right_tab_slack_tw
                                    + s958_center_slack(center_tab_stop_tw, current_width_tw)
                                    + segment_hang_tw;
                                let overflows = current_width_tw + piece_w_tw > limit;
                                if !s1059 || !overflows || piece_start >= cc {
                                    let seg: String = wchars[piece_start..cc].iter().collect();
                                    current_line.fragments.push(LineFragment {
                                        auto_space_shrink: 0.0,
                                        text: seg, width: piece_w, natural_width: piece_w, style: segment_style.clone(),
                                        tab_alignment: None, tab_position: None, field_type: wft,
                                        run_index: ridx, char_offset: choff + (piece_start - seg_start),
                                    });
                                    current_width += piece_w;
                                    current_width_tw += pt_to_tw(piece_w);
                                    current_capw_tw += pt_to_tw(piece_w);
                                    break;
                                }
                                // Longest prefix of [piece_start, cc) that fits.
                                let mut k = piece_start;
                                while k < cc {
                                    let w = word_char_ws[k] - piece_start_w;
                                    // Only the complete final segment can spend
                                    // its trailing punctuation allowance.
                                    let prefix_limit = if k + 1 == cc { limit } else { limit - segment_hang_tw };
                                    if current_width_tw + pt_to_tw(w) > prefix_limit { break; }
                                    k += 1;
                                }
                                if k <= piece_start {
                                    if !current_line.fragments.is_empty() && !para_all_whitespace {
                                        wrap_and_seed!(ws);
                                        continue; // retry on the fresh line
                                    }
                                    k = piece_start + 1; // an empty line takes one char
                                }
                                if k < cc {
                                    current_line.emergency_word_break = true;
                                }
                                let pw = word_char_ws[k - 1] - piece_start_w;
                                let seg: String = wchars[piece_start..k].iter().collect();
                                current_line.fragments.push(LineFragment {
                                    auto_space_shrink: 0.0,
                                    text: seg, width: pw, natural_width: pw, style: segment_style.clone(),
                                    tab_alignment: None, tab_position: None, field_type: wft,
                                    run_index: ridx, char_offset: choff + (piece_start - seg_start),
                                });
                                current_width += pw;
                                current_width_tw += pt_to_tw(pw);
                                current_capw_tw += pt_to_tw(pw);
                                piece_start_w = word_char_ws[k - 1];
                                piece_start = k;
                                if piece_start >= cc { break; }
                                wrap_and_seed!(ws);
                            }
                            let _ = seg_w_tw;
                            seg_start = cc; seg_start_w = cw;
                        }
                        word.clear();
                        word_char_ws.clear(); // S1059
                        word_width = 0.0;
                        word_natural_width = 0.0;
                    } else {
                    // S1022 badness (the whole-word branch) uses the SHARED
                    // s1022_badness_wrap computed before the latin_wordwrap branch above.
                    // Day 33 part 19: skip wrap break for all-whitespace paragraphs.
                    if s1026_replay_on && c14_active {
                        let cn = word.chars().filter(|c| !c.is_whitespace()).count();
                        let cr = if c14_active && c14_space_tw > 0 { latin_space_credit_tw } else { latin_space_credit_tw + wpj_credit_at(lines.len()) };
                        let hg = (if c14_active && c14_space_tw > 0 { if !s1026_final_token && std::env::var("OXI_S1028_HG_DISABLE").is_err() { (pt_to_tw(word_trail_hang_w) - 1).max(0) } else { 0 } } else { pt_to_tw(word_trail_hang_w) });
                        let ts = right_tab_slack_tw + s958_center_slack(center_tab_stop_tw, current_width_tw);
                        let cap = current_width_tw + word_fit_width_tw > available_tw + cr + ts + hg;
                        let wr = (cap || s1022_badness_wrap) && !current_line.fragments.is_empty() && !para_all_whitespace;
                        eprintln!("[S1026-REPLAY] para={} pass={} start={} end={} cand={:?} cw_tw={} curw_tw={} avail_tw={} credit_tw={} tabslack_tw={} line={} branch=whole_word badness={} cap={} dec={}",
                            s1026_replay_para.unwrap(), s1026_replay_pass, s1026_nonws_consumed.saturating_sub(cn), s1026_nonws_consumed, word, word_width_tw, current_width_tw, available_tw, cr, ts, lines.len(), s1022_badness_wrap, cap, if wr {"WRAP"} else {"KEEP"});
                    }
                    let mut hyphenated = false;
                    if (current_width_tw + word_fit_width_tw > available_tw + s1346_credit_tw + (if c14_active && c14_space_tw > 0 { latin_space_credit_tw } else { s1475_last_word_cap(latin_space_credit_tw + wpj_credit_at(lines.len()), word_fit_width_tw, s1475_space_tw, s1475_on) }) + right_tab_slack_tw + s958_center_slack(center_tab_stop_tw, current_width_tw) + (if c14_active && c14_space_tw > 0 { if !s1026_final_token && std::env::var("OXI_S1028_HG_DISABLE").is_err() { (pt_to_tw(word_trail_hang_w) - 1).max(0) } else { 0 } } else { pt_to_tw(word_trail_hang_w) }) || s1022_badness_wrap) && !current_line.fragments.is_empty()
                        && !para_all_whitespace {
                        // S1128 (2026-08-15, SHIPPED default-ON, opt-out
                        // OXI_S1128_DISABLE): `<w:autoHyphenation/>`.
                        // DERIVED in layout::hyphen from a 120-arm Word probe:
                        //   (a) if the gap left by the last WHOLE word is within the
                        //       hyphenation zone (w:hyphenationZone, default 18pt),
                        //       Word does not hyphenate at all;
                        //   (b) otherwise it takes the LONGEST legal prefix whose
                        //       width plus the hyphen still fits, and wraps the whole
                        //       word when none does.
                        // Latin-only (the JP corpus sets no autoHyphenation and
                        // hyphen::break_offsets rejects non-alphabetic tokens anyway).
                        // (declared above the wrap block — the advance below needs it)
                        if self.auto_hyphenation
                            && !self.doc_body_has_real_cjk
                            && std::env::var("OXI_S1128_DISABLE").is_err()
                        {
                            let gap = (available_tw - current_width_tw) as f32 / 20.0;
                            if gap > self.hyphenation_zone {
                                let hyphen_w = self
                                    .registry
                                    .char_width_pt_with_fallback(
                                        '-',
                                        self.resolve_font_size(&ws, para_style),
                                        &self.metrics_for(&ws, para_style),
                                    );
                                let room = gap - hyphen_w;
                                // char widths of the pending word, cumulative
                                let fs = self.resolve_font_size(&ws, para_style);
                                let m = self.metrics_for_text(&word, &ws, para_style);
                                let mut cum = 0.0f32;
                                let mut upto: Vec<(usize, f32)> = Vec::new();
                                for (bi, ch) in word.char_indices() {
                                    cum += self.registry.char_width_pt_with_fallback(ch, fs, &m);
                                    upto.push((bi + ch.len_utf8(), cum));
                                }
                                let mut best: Option<(usize, f32)> = None;
                                for off in hyphen::break_offsets(&word) {
                                    if let Some(&(_, w)) = upto.iter().find(|(b, _)| *b == off) {
                                        if w <= room {
                                            best = Some((off, w));
                                        }
                                    }
                                }
                                if let Some((off, w)) = best {
                                    let head = format!("{}-", &word[..off]);
                                    let hw = w + hyphen_w;
                                    current_line.fragments.push(LineFragment {
                                        auto_space_shrink: 0.0,
                                        text: head,
                                        width: hw,
                                        natural_width: hw,
                                        style: ws.clone(),
                                        tab_alignment: None,
                                        tab_position: None,
                                        field_type: wft,
                                        run_index: word_run_index,
                                        char_offset: word_char_offset,
                                    });
                                    current_width += hw;
                                    current_width_tw += pt_to_tw(hw);
                                    current_capw_tw += pt_to_tw(hw);
                                    let tail = word[off..].to_string();
                                    let tail_w = word_width - w;
                                    word = tail;
                                    word_width = tail_w;
                                    word_natural_width = tail_w;
                                    word_char_offset += off;
                                    hyphenated = true;
                                }
                            }
                        }
                        wrap_and_seed!(ws);
                    }
                    // S1128: the hyphenation branch replaces `word` with its TAIL, so
                    // the twips advance has to be recomputed — but ONLY then. Doing it
                    // unconditionally overwrote S1061b's exact-sum advance with a plain
                    // rounded one for every word, which silently took back S1061b on
                    // creative__0158c02ae (0.9247 -> 0.8622, found by bisecting the
                    // score against d42eac81 — the opt-out flag did not restore it,
                    // which is what pointed at a change outside the flag).
                    let word_width_tw = if hyphenated {
                        pt_to_tw(word_width)
                    } else {
                        word_width_tw
                    };
                    if std::env::var("OXI_DBGWRAP").is_ok() && word.starts_with("33") {
                        eprintln!("[FRAG-WORD] word={:?} ww={:.2} cw_tw={} avail={} nfrag={} lines={}", word, word_width, current_width_tw, available_tw, current_line.fragments.len(), lines.len());
                    }
                    current_line.fragments.push(LineFragment {
                        auto_space_shrink: 0.0,
                        text: std::mem::take(&mut word),
                        width: word_width,
                        natural_width: word_natural_width,
                        style: ws,
                        tab_alignment: None,
                        tab_position: None,
                        field_type: wft,
                        run_index: word_run_index,
                        char_offset: word_char_offset,
                    });
                    current_width += word_width;
                    current_width_tw += word_width_tw;
                    current_capw_tw += word_width_tw; // S475: words have no punct capacity
                    word_width = 0.0;
                    word_natural_width = 0.0;
                    if latin_wordwrap { word_breaks.clear(); word_seg_meta.clear(); word_char_ws.clear(); }
                    }
                    word_seg_styles.clear();
                    if latin_wordwrap { seg_pending = false; }
                }
            };
        }

        let n_fragments = fragments.len();
        // S547 (2026-06-12): w:kern resolved per paragraph (any fragment rPr
        // kern>0, else the paragraph/docDefaults default-run kern). Gates the
        // yakumono pair halving (break-time rule AND the S532 Stage-2 revert
        // protection). See the gate comment at yakumono_pair_enabled below.
        let para_kern_on = fragments
            .iter()
            .any(|&(_, st, _, _, _)| st.kern.map_or(false, |k| k > 0.0))
            || para_style
                .default_run_style
                .as_ref()
                .and_then(|rs| rs.kern)
                .map_or(false, |k| k > 0.0);
        let s547_kern_gate = std::env::var("OXI_S547_DISABLE").is_err();
        // S466 (2026-05-31, SHIPPED default-ON, opt-out OXI_S466_DISABLE):
        // hoisted once — see h8_trigger comment below. (a) compute positive grid
        // expansion for fs>=default, AND (b) apply it to char_width so the WRAP
        // (chars/line) matches Word, not just positioning. Pairs with the parser
        // raw_pitch change (ooxml.rs). Gate (drift-free OFF-vs-ON same binary):
        // charGrid family +0.0019, bottom-N floor up (tokumei p4/p5), only
        // tokumei p7 ×4 regress (above the floor); Phase-1 54/55 preserved.
        let s466_grid_expand = std::env::var("OXI_S466_DISABLE").is_err();
        // ORPHAN-OIKOMI experiment config, hoisted once (default OFF = byte-identical,
        // zero hot-path cost — the per-fragment char-count precompute below only runs
        // when enabled). Validated mechanism (fixes nedo para 333 «…規定する子» 2→1
        // line = Word) but not yet shippable: fixing 333 EXPOSES a compensating
        // downstream under-count (-1×3 at 400/434/465, 0.9979→0.9938). See
        // [[char_budget_wall]]. OXI_ORPHAN_OPEN / OXI_ORPHAN_LINEMULT tune.
        let orphan_oikomi_on = std::env::var("OXI_ORPHAN_OIKOMI").ok().as_deref() == Some("1");
        let orphan_open_cap: f32 = std::env::var("OXI_ORPHAN_OPEN")
            .ok()
            .and_then(|v| v.parse().ok())
            .unwrap_or(3.34);
        let orphan_line_mult: f32 = std::env::var("OXI_ORPHAN_LINEMULT")
            .ok()
            .and_then(|v| v.parse().ok())
            .unwrap_or(1.3);
        // Styling boundaries do not split a prohibited-start punctuation unit.
        // Keep paragraph character coordinates for reservations, while keeping
        // the original fragments and their metrics for measurement and painting.
        let punctuation_chars: Vec<char> = fragments.iter().flat_map(|f| f.0.chars()).collect();
        let mut punctuation_offsets = Vec::with_capacity(fragments.len());
        let mut punctuation_offset = 0;
        for fragment in fragments {
            punctuation_offsets.push(punctuation_offset);
            punctuation_offset += fragment.0.chars().count();
        }
        let mut accepted_punctuation_unit: Option<(usize, usize)> = None;
        for (frag_outer_idx, &(text, style, frag_field_type, frag_run_index, frag_char_start)) in
            fragments.iter().enumerate()
        {
            // S655 (2026-06-24): flush the accumulated word at a w:position
            // boundary so the per-run vertical shift survives as its OWN
            // fragment. Adjacent Latin runs with no break char between them
            // otherwise concatenate into a single word/fragment carrying only the
            // FIRST run's style (word_style), DROPPING the next run's position
            // (and the line-height growth it needs) — the OXI_DBG655 root cause.
            // NARROW: fires only when adjacent runs differ in position; every
            // non-position doc has None==None → no flush → byte-identical.
            // Opt-out OXI_S655_DISABLE.
            if std::env::var("OXI_S655_DISABLE").is_err()
                && frag_outer_idx > 0
                && fragments[frag_outer_idx - 1].1.position != style.position
            {
                flush_word!(fragments[frag_outer_idx - 1].1);
            }
            // S677: flush at a small-caps size boundary (full ↔ 0.8× segments).
            if caps_size_split
                && frag_outer_idx > 0
                && fragments[frag_outer_idx - 1].1.font_size != style.font_size
            {
                flush_word!(fragments[frag_outer_idx - 1].1);
            }
            // S899 (2026-07-17): flush at a vertAlign boundary — a
            // superscript/subscript run merged into the accumulated word
            // keeps the FIRST run's style, so a footnote-ref digit glued to
            // its word («Clark3;») measured AND rendered at FULL size (12pt
            // vs Word's 2/3 auto-shrink 8pt). 81e80 L8 carried 4 such refs =
            // +8pt phantom width → «Suckling6;» wrapped → a 3-line phase
            // shift by p3 (the {+1:3} residual). The S655/S677 fragment-
            // flattening class; None==None for ref-free text → byte-identical.
            if std::env::var("OXI_S899_DISABLE").is_err()
                && frag_outer_idx > 0
                && fragments[frag_outer_idx - 1].1.vertical_align != style.vertical_align
            {
                flush_word!(fragments[frag_outer_idx - 1].1);
            }
            // A private-use glyph is meaningful only in its specified face.
            // Preserve font boundaries as paint segments inside the pending word;
            // flushing here would introduce a break inside BBBB + symbol + CCCC.
            if latin_wordwrap && frag_outer_idx > 0 && !word.is_empty() {
                let (previous_text, previous_style, ..) = fragments[frag_outer_idx - 1];
                let has_private_glyph = |value: &str| value.chars().any(|c| matches!(c as u32, 0xE000..=0xF8FF));
                let script_size_boundary = std::env::var("OXI_SCRIPT_SIZE_SEGMENTS_DISABLE").is_err()
                    && previous_style.vertical_align == style.vertical_align
                    && matches!(style.vertical_align, Some(VerticalAlign::Superscript | VerticalAlign::Subscript))
                    && previous_style.font_size != style.font_size;
                // Preserve each source run's metrics, including a note reference
                // whose inherited size differs inside an unbroken word.
                let metric_style_boundary = std::env::var_os("OXI_RUN_STYLE_SEGMENTS_DISABLE").is_none()
                    && (previous_style.font_size != style.font_size
                        || previous_style.font_family != style.font_family
                        || previous_style.font_family_east_asia != style.font_family_east_asia
                        || previous_style.font_family_cs != style.font_family_cs
                        || previous_style.bold != style.bold
                        || previous_style.italic != style.italic
                        || previous_style.run_border != style.run_border);
                if metric_style_boundary || script_size_boundary || ((has_private_glyph(previous_text) || has_private_glyph(text))
                    && (previous_style.font_family != style.font_family
                        || previous_style.font_family_east_asia != style.font_family_east_asia
                        || previous_style.font_family_cs != style.font_family_cs))
                {
                    let offset = word.chars().count();
                    if word_breaks.last().map_or(true, |&(end, _)| end != offset) {
                        word_breaks.push((offset, word_width));
                    }
                    word_seg_styles.push((offset, style.clone()));
                    word_seg_meta.push((offset, frag_run_index, frag_char_start));
                }
            }
            let font_size = self.resolve_font_size(style, para_style);
            // S899b (2026-07-17): the BREAK width of a superscript/subscript
            // fragment uses the 2/3-shrunk size even when font_size came from
            // the STYLE chain (resolve_font_size's is_none() gate only covers
            // the direct-size-absent case, so a style-inherited 12pt fn-ref
            // digit measured at FULL width 6.67 vs Word's rendered 8.04pt →
            // 4.47 advance; 81e80 L9 carried 4 refs = +13.2pt phantom →
            // «Job v. Potton18.» wrapped where Word ends the para on L9).
            let font_size = if std::env::var("OXI_S899_DISABLE").is_err()
                && style.font_size.is_some()
                && matches!(
                    style.vertical_align,
                    Some(VerticalAlign::Superscript) | Some(VerticalAlign::Subscript)
                ) {
                LayoutEngine::vertical_align_font_size(font_size)
            } else {
                font_size
            };
            // S700 (2026-06-30): a vert_in_horz run (eastAsianLayout w:vert, 縦中横
            // / tate-chu-yoko) is an ATOMIC vertical column — its n chars stack
            // downward in ONE 1-em-wide cell, so it advances exactly fs horizontally
            // (independent of char count) and never breaks internally. Push the WHOLE
            // run as one LineFragment of width fs; the emit loop renders it as an
            // is_vertical column and the line-height fns grow the line to fit it.
            // Gated on the parsed flag (false for the whole corpus → byte-identical;
            // verified 0/corpus uses eastAsianLayout w:vert). Opt-out OXI_S700_DISABLE.
            if style.vert_in_horz && !text.is_empty() && std::env::var("OXI_S700_DISABLE").is_err()
            {
                flush_word!(style);
                let vw_tw = pt_to_tw(font_size);
                if current_width_tw + vw_tw > available_tw
                    && !current_line.fragments.is_empty()
                    && !para_all_whitespace
                {
                    lines.push(std::mem::take(&mut current_line));
                    current_width = 0.0;
                    current_width_tw = 0;
                    current_capw_tw = 0;
                    latin_space_credit_tw = 0;
                    latin_space_credit_remainder = 0.0;
                    right_tab_slack_tw = 0;
                    center_tab_stop_tw = None;
                    compress_used = false;
                }
                current_line.fragments.push(LineFragment {
                    auto_space_shrink: 0.0,
                    text: text.to_string(),
                    width: font_size,
                    natural_width: font_size,
                    style: style.clone(),
                    tab_alignment: None,
                    tab_position: None,
                    field_type: frag_field_type,
                    run_index: frag_run_index,
                    char_offset: frag_char_start,
                });
                current_width += font_size;
                current_width_tw += vw_tw;
                current_capw_tw += vw_tw;
                continue;
            }
            // S703 (2026-06-30): a `combine` run (eastAsianLayout w:combine, 割注 /
            // kumimoji — two-lines-in-one) is an atomic COMPACT unit: its n chars are
            // set as 2 small (~half-size) rows within ONE line height, optionally in
            // brackets. Width ≈ ceil(n/2) half-size chars + (brackets). Push as ONE
            // fragment; the emit loop renders the warichu. Opt-out OXI_S703_DISABLE.
            // s476_body gate: only the BODY emit renders the warichu (the cell /
            // estimate breakers produce a tuple that has no `combine` flag, so a
            // compact width there would cram full-size text). Cells stay full-size
            // (byte-identical) — the cell warichu is a deferred follow-up.
            // S1314b (2026-09-05, default ON, opt-out OXI_S1314_DISABLE): a ruby
            // field never breaks inside its base -- Word moves the whole field
            // (spread base + annotation) to the next line. reference__0cf9c879
            // split 吉|田松陰 across lines and wrapped 18 lines where Word wraps
            // 16. The field's width is the spread base (character spacing
            // included), exempt from the char grid like a kumimoji unit.
            if style.ruby_field
                && !text.is_empty()
                && std::env::var("OXI_S1314_DISABLE").is_err()
            {
                flush_word!(style);
                let cs_field = style.character_spacing.unwrap_or(0.0);
                let m_field = self.metrics_for_text(text, style, para_style);
                let cw: f32 = text
                    .chars()
                    .map(|c| self.registry.char_width_pt_with_fallback(c, font_size, &m_field) + cs_field)
                    .sum();
                let cw_tw = pt_to_tw(cw);
                if (if vertical { current_capw_tw } else { current_width_tw }) + cw_tw > available_tw
                    && !current_line.fragments.is_empty()
                    && !para_all_whitespace
                {
                    lines.push(std::mem::take(&mut current_line));
                    current_width = 0.0;
                    current_width_tw = 0;
                    current_capw_tw = 0;
                    latin_space_credit_tw = 0;
                    latin_space_credit_remainder = 0.0;
                    right_tab_slack_tw = 0;
                    center_tab_stop_tw = None;
                    compress_used = false;
                }
                current_line.fragments.push(LineFragment {
                    auto_space_shrink: 0.0,
                    text: text.to_string(),
                    width: cw,
                    natural_width: cw,
                    style: style.clone(),
                    tab_alignment: None,
                    tab_position: None,
                    field_type: frag_field_type,
                    run_index: frag_run_index,
                    char_offset: frag_char_start,
                });
                current_width += cw;
                current_width_tw += cw_tw;
                current_capw_tw += cw_tw;
                continue;
            }
            if style.combine
                && !text.is_empty()
                && s476_body
                && std::env::var("OXI_S703_DISABLE").is_err()
            {
                flush_word!(style);
                let n = text.chars().count();
                let rows_chars = (n + 1) / 2;
                let has_br = style
                    .combine_brackets
                    .as_deref()
                    .map_or(false, |b| b != "none");
                let cw = rows_chars as f32 * (font_size * 0.5)
                    + if has_br { font_size * 0.8 } else { 0.0 };
                let cw_tw = pt_to_tw(cw);
                if current_width_tw + cw_tw > available_tw
                    && !current_line.fragments.is_empty()
                    && !para_all_whitespace
                {
                    lines.push(std::mem::take(&mut current_line));
                    current_width = 0.0;
                    current_width_tw = 0;
                    current_capw_tw = 0;
                    latin_space_credit_tw = 0;
                    latin_space_credit_remainder = 0.0;
                    right_tab_slack_tw = 0;
                    center_tab_stop_tw = None;
                    compress_used = false;
                }
                current_line.fragments.push(LineFragment {
                    auto_space_shrink: 0.0,
                    text: text.to_string(),
                    width: cw,
                    natural_width: cw,
                    style: style.clone(),
                    tab_alignment: None,
                    tab_position: None,
                    field_type: frag_field_type,
                    run_index: frag_run_index,
                    char_offset: frag_char_start,
                });
                current_width += cw;
                current_width_tw += cw_tw;
                current_capw_tw += cw_tw;
                continue;
            }
            // S839 (2026-07-14, opt-out OXI_S839_DISABLE): an INLINE visual
            // vector-group drawing (wpg without txbxContent — hmrc's checkbox
            // strips / heavy rules) is a WIDTH-BEARING atomic object in its
            // host line: Word reserves cx and positions it via the normal tab
            // machinery (the NI 9-box strip is CENTERED at its tab stop —
            // measured x330.5 = 413.85 − 166.7/2). Push ONE U+FFFC fragment of
            // the drawing width; the emit loop draws tb.vector_shapes at the
            // fragment x and never emits the FFFC as text. The marker style
            // field is set ONLY for S535-signature inline visual drawings
            // (wpg = hmrc/framework only by corpus scan; framework's groups
            // are text-bearing → never marked → byte-identical elsewhere).
            if let Some((ow, _oh)) = style.inline_object_extent {
                if std::env::var("OXI_S839_DISABLE").is_err() {
                    flush_word!(style);
                    // S852: a horizontal rule (o:hr) occupies its OWN line —
                    // flush the current line (e.g. the title text) so the rule
                    // starts fresh, and end its line afterward.
                    let is_hr = style.hr_rule.is_some();
                    if is_hr && !current_line.fragments.is_empty() {
                        lines.push(std::mem::take(&mut current_line));
                        current_width = 0.0;
                        current_width_tw = 0;
                        current_capw_tw = 0;
                        latin_space_credit_tw = 0;
                    latin_space_credit_remainder = 0.0;
                        right_tab_slack_tw = 0;
                        center_tab_stop_tw = None;
                        compress_used = false;
                    }
                    let ow_tw = pt_to_tw(ow);
                    // NO overflow wrap for the object fragment (v1): hmrc's
                    // strips follow CENTER tabs — the raw width test would
                    // compare stop + full width against the column (543 > 521)
                    // and wrap, but the center post-process (below, ~14360)
                    // shifts the segment to stop − w/2 and Word fits it
                    // ([330.5..497.2] < 538). The corpus-scope objects always
                    // fit their host line in Word.
                    // S1040 (2026-07-29, opt-out OXI_S1040_DISABLE): ...except an
                    // inline PICTURE routed by S1034/S854, which Word DOES wrap.
                    // JA blind policies__03a9dca2 has a 425.2pt screenshot after a
                    // heading on a 425.2pt column: Word puts the heading on line 1
                    // and the picture on line 2; without a wrap the picture shared
                    // the heading's line and (being taller than the text ascent)
                    // rendered off the top of the page, SSIM 0.928 -> 0.854. Scoped
                    // to inline_object_image (a real picture) and to lines with NO
                    // tab fragment, which is exactly the exclusion the note above
                    // documents (hmrc's strips ride CENTER tabs and are re-placed
                    // by the post-pass, so a raw-width test would wrap them wrongly).
                    // S1252: inline MATHS wraps the same way. Word truth
                    // (_pb_inlmath `nofit` arm): with the lead text filling the
                    // column, `f(x)=4cos(3x)` moves whole to line 2
                    // (`line right up to 𝑓(𝑥) = 4 cos(3𝑥)`, adv 23.30 = two
                    // lines) — it is an atom in the normal wrap, not an
                    // overflow.
                    if std::env::var("OXI_S1040_DISABLE").is_err()
                        && (style.inline_object_image.is_some()
                            || style.inline_math.is_some())
                        && !current_line.fragments.is_empty()
                        && current_line
                            .fragments
                            .iter()
                            .all(|f| f.tab_alignment.is_none())
                        && current_width_tw + ow_tw > available_tw
                    {
                        lines.push(std::mem::take(&mut current_line));
                        current_width = 0.0;
                        current_width_tw = 0;
                        current_capw_tw = 0;
                        latin_space_credit_tw = 0;
                    latin_space_credit_remainder = 0.0;
                        right_tab_slack_tw = 0;
                        center_tab_stop_tw = None;
                        compress_used = false;
                    }
                    current_line.fragments.push(LineFragment {
                        auto_space_shrink: 0.0,
                        text: "\u{FFFC}".to_string(),
                        width: ow,
                        natural_width: ow,
                        style: style.clone(),
                        tab_alignment: None,
                        tab_position: None,
                        field_type: frag_field_type,
                        run_index: frag_run_index,
                        char_offset: frag_char_start,
                    });
                    current_width += ow;
                    current_width_tw += ow_tw;
                    current_capw_tw += ow_tw;
                    if is_hr {
                        lines.push(std::mem::take(&mut current_line));
                        current_width = 0.0;
                        current_width_tw = 0;
                        current_capw_tw = 0;
                        latin_space_credit_tw = 0;
                    latin_space_credit_remainder = 0.0;
                        right_tab_slack_tw = 0;
                        center_tab_stop_tw = None;
                        compress_used = false;
                    }
                    continue;
                }
            }
            let mut char_pos_in_run = frag_char_start;
            // Char counts in PRECEDING / SUBSEQUENT fragments (paragraph-level total/
            // remaining lookahead for the short-para oikomi gate). Only when enabled.
            let (chars_before_frag, chars_after_frag) = if orphan_oikomi_on {
                (
                    fragments[..frag_outer_idx]
                        .iter()
                        .map(|f| f.0.chars().count())
                        .sum::<usize>(),
                    fragments[frag_outer_idx + 1..]
                        .iter()
                        .map(|f| f.0.chars().count())
                        .sum::<usize>(),
                )
            } else {
                (0, 0)
            };
            // S725 (2026-07-03): chars in SUBSEQUENT fragments, computed
            // unconditionally (chars_after_frag above is orphan-gated). 0 means
            // this fragment carries the paragraph's FINAL characters — used by
            // the tail-natural rule at the s475 overflow test.
            let s725_chars_after: usize = fragments[frag_outer_idx + 1..]
                .iter()
                .map(|f| f.0.chars().count())
                .sum();

            // fitText runs: skip GDI snap to preserve exact target width
            let cs = if style.fit_text.is_some() || style.ruby_spread {
                style.character_spacing.unwrap_or(0.0)
            } else {
                snap_character_spacing(style.character_spacing.unwrap_or(0.0))
            };

            // Pre-resolve font metrics and GDI width maps for this fragment.
            // Avoids repeated font family resolution and HashMap lookups per character.
            let latin_metric_owner = self.metrics_for(style, para_style);
            let latin_metrics = &*latin_metric_owner;
            let cjk_metric_owner = self.metrics_for_cjk(style, para_style);
            let cjk_metrics = cjk_metric_owner.as_deref();
            // Break widths follow the substituted face used to draw CJK glyphs.
            let substitute_metric_owner = if (std::env::var_os("OXI_CJK_SUBSTITUTE_METRICS_DISABLE").is_none()
                    || std::env::var_os("OXI_CJK_SUBSTITUTE_METRICS").is_some()) {
                self.metrics_for_cjk_script(style, para_style, true)
            } else {
                cjk_metric_owner.clone()
            };
            let substitute_metrics = substitute_metric_owner.as_deref();
            let substitute_gdi_map = substitute_metrics
                .and_then(|m| self.registry.get_gdi_char_widths(&m.family, font_size));
            let latin_gdi_map = self
                .registry
                .get_gdi_char_widths(&latin_metrics.family, font_size);
            let cjk_gdi_map = cjk_metrics
                .map(|m| self.registry.get_gdi_char_widths(&m.family, font_size))
                .flatten();

            // Yakumono compression flags (約物詰め): COM-confirmed (2026-04-08,
            // refined 2026-04-18 bisect).
            // Trigger requires BOTH:
            //   - w:characterSpacingControl = "compressPunctuation" or "compressPunctuationAndJapaneseKana"
            //   - w:compat/w:compatSetting compatibilityMode >= 15 (Word 2013+)
            // See RESEARCH_LOG.md 2026-04-18 and pipeline_data/d77a_yakumono_bisect.json:
            //   - cSC alone: NO compression (minimal repro confirmed)
            //   - cSC + compat15: yakumono pair compression applies
            // Most modern docs use "doNotCompress" → no compression regardless of compat.
            //
            // 2026-04-21 update: COM evidence on 4 distinct compat=14+cP docs
            // (04b88, 7f272a, fded68, 34140) showed Word fits +1 to +3 more chars
            // on line 1 of yakumono+indent paragraphs vs Oxi. Minimal repro
            // (idx46_real with compat=14) confirmed Word=43 / Oxi=42. The
            // `compat>=15` gate excluded compat=14 docs that Word DOES compress.
            // Drop the compat gate — `compress_punctuation` alone matches
            // Word's behavior for both compat 14 and compat 15.
            //
            // 2026-04-27 attempted-but-reverted: V_CP + V_COMPAT15 8x8 matrix
            // measurement (180 fixtures) showed Word applies next-trigger
            // compression unconditionally on isolated 4-char paragraphs across
            // compat∈{14,15} × cSC∈{doNotCompress,compressPunctuation} ×
            // useFELayout∈{on,off} × kern∈{on,off}. Patched to
            // `yakumono_enabled = true` and ran pipeline.verify on 177 docs:
            // **18 page regressions, net -2.0184, bottom-5 floor 3.2645→2.9337
            // (-0.3308 catastrophic)**. Two docs collapsed:
            //   - 0e7af1ae8f21 pages 2-7,8,10: -0.18 to -0.29 each
            //   - 683ffcab86e2 pages 1-3: -0.04 to -0.25
            // 6 e3c5 pages improved (+0.01 to +0.03) but vastly outweighed.
            // **Implication**: COM 4-char isolated fixtures do not generalize
            // to multi-line real-world paragraphs. Word's actual gate involves
            // additional context (line position, surrounding chars, paragraph
            // structure?) NOT captured in the 8x8 grid. Reverted on 2026-04-27.
            // See RESEARCH_LOG 2026-04-27 falsified entry for full data.
            //
            // Day 34 part 23 (2026-05-13): COM measurement of e3c545 idx=29
            // (Meiryo 10.5pt, csControl=doNotCompress) showed Word DOES
            // compress the 、 of 、「 pair to 5.25pt (half). This contradicts
            // the OOXML doNotCompress flag — Word applies pair compression
            // when the CJK font has hwid (halfwidth-glyph) support.
            // Two-tier gate:
            //   - PAIR rule (close+open, ×0.5): compress_punctuation OR hwid
            //   - FULL rules (expand pair, standalone, line-start): compress_punctuation only
            //     (preserves existing behavior; 2026-04-27 unconditional broke MS Mincho docs)
            let cjk_font_has_hwid = cjk_metrics
                .map(|m| crate::font::font_supports_hwid(&m.family))
                .unwrap_or(false);
            // S547 (2026-06-12): the pair-halving gate is w:kern — NOT
            // compressPunctuation and NOT compat. 2×2×2 COM matrix
            // (_s547b_gate_matrix.py): kern=2 halves 、（/（「 even under
            // doNotCompress at any compat; kern absent never halves even with
            // compressPunctuation (the full 26×26 sweep at kern=0 had ZERO
            // non-natural pairs). S532's "unconditional" was measured on
            // kern=2 docs. kern is a RUN property; resolved per fragment:
            // run rPr → paragraph-style chain (default_run_style, basedOn
            // merge carries kern) → docDefaults (merge in ooxml.rs). The
            // pair scan below is per-fragment (chars_vec), so fragment
            // granularity is exact. Opt-out OXI_S547_DISABLE restores the
            // pre-S547 compress_punctuation gate.
            let frag_kern_on = style
                .kern
                .or_else(|| para_style.default_run_style.as_ref().and_then(|rs| rs.kern))
                .map_or(false, |k| k > 0.0);
            let yakumono_pair_enabled = if s547_kern_gate {
                frag_kern_on || cjk_font_has_hwid
            } else {
                self.compress_punctuation || cjk_font_has_hwid
            };
            let yakumono_enabled = self.compress_punctuation;
            // S472 (2026-06-01) DEMAND-DRIVEN yakumono refactor (user chose the big
            // refactor path). Word does NOT pre-compress standalone 、 at wrap time:
            // it uses NATURAL fullwidth (12pt) for the break decision, then compresses
            // 、 trailing space ONLY as much as a line's justify-slack demands
            // (COM: b837 p1 、 = 8.0-12.0pt variable per line; d77a divergent line
            // 、 = 11.2 = LIGHT compression, not the flat ×0.6667=8.0 Oxi pre-applies).
            // Oxi's flat pre-compression over-packs (fits 1 extra char/、-heavy line).
            // When enabled: (1) standalone 、，use natural width at break, (2) the
            // overflow-absorb budget becomes (count of standalone 、 on line)×(fs/3)
            // [max 4pt each at 12pt], (3) on absorb the line's 、 are retroactively
            // compressed by the absorbed overflow so the line fits exactly = matches
            // Word's demand-driven per-line compression. Default OFF (byte-identical)
            // for the canary. See session470 finding.
            // S473 (2026-06-01): break-flip-derived budget. Word's break-time punct
            // compression is DEMAND-driven up to a CAP of ~3.25pt/compressible
            // (fs×0.27, NOT fs/3=4.0), measured via repros/breakflip + d77a p1/p9.
            // s473_locomp implies the s472 upstream (leave 、 natural) + the s472
            // render water-fill, and additionally swaps the break budget to the
            // cap-based, no-0.95-exclusion model. Default OFF (byte-identical).
            // S474 (2026-06-01): pure-natural break diagnostic. Disables ALL
            // break-time standalone-punct compression (the ×0.6667 AND the
            // line-start narrow-yakumono reduction) and the demand-absorb, keeping
            // only the always-on pair compression. Renders the natural-greedy line
            // counts (Ng) = the count if Word broke at natural widths. Used to test
            // "Word breaks at natural, punct compression is render-only" and to
            // derive the fullness-gate rule (Ng vs Word count). Default OFF.
            // S492 (2026-06-03) — jc=left disentanglement (the R35 multi-session
            // refactor, user option A). Word does NOT apply yakumono break-time
            // compression to NON-justified paragraphs: jc=left/right/center break
            // at NATURAL widths + kinsoku (punct compression is justify-specific).
            // MEASURED decisively (tools/metrics/measure_jc_disentangle*.py):
            //   - jc=left synthetic repro = punct 12.0 natural, zero compression,
            //     at every punct density (10-50%); jc=both packs +1 via burasagari
            //     (hangable punct hangs past margin; mid-line punct stays 12.0) or
            //     light opener compression (「→11.25), NOT distributed K-compression.
            //   - Real docs 683f/0e7af/d77a (docGrid type=lines) OVER-PACK +1 on
            //     jc=left wrapping lines because Oxi's S475 capacity break is ungated
            //     on alignment; b837 (linesAndChars, grid-determined) is unaffected.
            // The render water-fill is ALREADY jc-gated (it lives inside the
            // should_justify block). Only the BREAK side leaks. When set + the para
            // is NOT justified, run the validated s474_natural pure-natural-greedy
            // path (disables standalone compression + absorb) and disable the S475
            // capacity break. Default OFF (byte-identical). Phase-1-sensitive (fewer
            // chars/line on jc=left → more lines → pagination shift) → env-gated +
            // full canary before ship.
            // S539 (2026-06-11): SHIPPED default-ON, scoped to NON-linesAndChars
            // (the former OXI_S492_JCNATURAL + OXI_S492_LINESONLY config).
            // The S492-era blocker was 3a4f p2-p5 SSIM regressions; S539 traced
            // them to the style-basedOn jc-inheritance bug (parser/styles.rs):
            // paragraphs Word justifies (Normal jc=both via pStyle chain) were
            // resolved jc=left by Oxi, so the natural break wrongly rewrapped
            // them. With jc resolution fixed, the full-corpus gate is clean:
            // SSIM 1 up (d77a p2 +0.0090) / 409 unchanged / 0 regress,
            // Phase-1 54/55 with 3a4f histogram identical to baseline.
            // linesAndChars (b837 family) stays EXCLUDED: full-scope round-1
            // gate showed b837 pagination cascade 0.9997->0.5775 (30 paras +1
            // page; p5-p7 SSIM -0.53) even though p1-p4 improved +0.039 — the
            // b837 jc=left grid-line break needs its own investigation before
            // widening (OXI_S492_JCNATURAL still forces the full scope).
            // Opt-out: OXI_S492_DISABLE restores capacity-break for all paras.
            let s492_full = std::env::var("OXI_S492_JCNATURAL").is_ok();
            // S572 (2026-06-14): extend the S568 LEGACY jc=left 約物 OIKOMI to
            // NO-TYPE docGrid docs (S568 covered only linesAndChars). ikujidetail
            // (compat=11, no-type docGrid linePitch=286, jc=left, compressPunctuation)
            // compresses a tight line's 約物 to fit — COM (_s572_charadv): on the
            // over-wrapping para i=199 Word renders the mid-line 、 at 9.0pt and the
            // line-end 。 at 5.25pt (half-width) where a slack line (i=5) keeps both
            // at 11.25 = DEMAND-DRIVEN oikomi, not general 約物詰め. Oxi broke at
            // NATURAL (natural_break_jc) → the trailing char over-wrapped 1→2 lines
            // (i=199/231/419) → +1×27. Same discriminator as S568 (compat<15 +
            // compressPunctuation + non-justified), just no-type instead of
            // linesAndChars. SCOPE: the ONLY compat<15 compressPunctuation no-type
            // docGrid doc in the corpus is ikujidetail (single-doc-scoped, like
            // S568). Opt-out OXI_S572_DISABLE.
            let s572_legacy_notype_oikomi = std::env::var("OXI_S572_DISABLE").is_err()
                && doc_grid_no_type
                && s476_body
                && !is_justified
                && self.compress_punctuation
                && self.compat_mode < 15;
            // S592 (2026-06-17, default ON, opt-out OXI_S592_DISABLE): a
            // PROPORTIONAL CJK font (MS PGothic / MS PMincho / HGPGothicM) in a
            // linesAndChars docGrid is OFF-GRID — its chars do NOT align to the
            // grid char pitch, so Word breaks it at NATURAL width (proportional),
            // NOT via the on-grid s476 capacity break. The natural_break_jc gate
            // excludes ALL lines_and_chars docs (the S492 §8 "off-grid footnote
            // vs on-grid body" distinction Oxi couldn't make) — but the font
            // proportionality IS that distinction. kojin (HGPGothicM, linesAndChars,
            // jc=None, compat=15): the s476 capacity break credited the 2 mid-line
            // 、 ~5.2pt of demand compression (over_tw=−26 where the ACTUAL width
            // overflows +78tw, OXI_DBG_KOJIN) → fit a trailing こ Word WRAPS →
            // para i297 packed 3 lines (Word 4) → i303 crept onto p11 = the −1.
            // Word does NOT compress these jc=None 約物 (left-aligned, no justify
            // demand; the line is proportional, off-grid). DISCRIMINATOR = the
            // dominant CJK font is proportional (pgothic family) — monospace
            // linesAndChars bodies (1ec/tokumei = MS Mincho/Gothic, on-grid) are
            // UNAFFECTED (they keep the grid capacity break). Only kojin + parttime
            // (HGPGothicM) match in the corpus. Opt-out OXI_S592_DISABLE.
            let para_off_grid = std::env::var("OXI_S592_DISABLE").is_err()
                && lines_and_chars
                && fragments
                    .iter()
                    .find(|(t, _, _, _, _)| {
                        t.chars().any(|c| !c.is_whitespace() && c != '\u{3000}')
                    })
                    .map_or(false, |(_, rs, _, _, _)| {
                        self.metrics_for_cjk(rs, para_style).map_or(false, |m| {
                            matches!(
                                m.family.as_str(),
                                "MS PGothic" | "MS PMincho" | "HGPGothicM"
                            )
                        })
                    });
            // Japanese run language permits a half-cell pull-in without a
            // character grid, and with a grid in modern compatibility mode. Word probes retain this limit with one to three
            // separated periods; English overrides retain natural wrapping.
            let english_ea_natural = s476_body && !vertical
                && style.east_asia_lang.as_deref().map_or(false, |lang|
                    lang.eq_ignore_ascii_case("en") || lang.to_ascii_lowercase().starts_with("en-"));
            // Proportional punctuation does not provide a full-width compression budget.
            let narrow_punctuation_natural = !lines_and_chars && cjk_metrics.map_or(false, |m| {
                ['\u{3001}', '\u{3002}'].iter().all(|&ch| {
                    let width = m.char_width_em(ch);
                    width > 0.0 && width < 0.99
                })
            });
            let natural_punctuation_boundary = english_ea_natural || narrow_punctuation_natural;
            let japanese_language_oikomi = s476_body && !vertical
                && !narrow_punctuation_natural
                && (self.compat_mode < 15 || is_justified)
                && !para_off_grid
                && (!lines_and_chars || self.compat_mode >= 15)
                && self.compress_punctuation
                && style.east_asia_lang.as_deref().map_or(false, |lang|
                    lang.eq_ignore_ascii_case("ja") || lang.to_ascii_lowercase().starts_with("ja-"));
            let natural_break_jc = std::env::var("OXI_S492_DISABLE").is_err()
                && !is_justified
                && (!lines_and_chars || s492_full || para_off_grid)
                && !s572_legacy_notype_oikomi
                && !japanese_language_oikomi;
            let s474_natural = std::env::var("OXI_S474_NATURAL").is_ok() || natural_break_jc || english_ea_natural || narrow_punctuation_natural;
            // S589 (2026-06-16, opt-IN OXI_S589=1, default OFF = byte-identical):
            // LEGACY (compat<15) JUSTIFIED body paras break standalone 、。，． at
            // NATURAL width instead of the flat ×0.6667 pre-compress (mod.rs:6764)
            // that over-packs 、-heavy lines by ~1 char. ROOT of tokyoshugyo #2:
            // _tks_oidashi.py (char-stream-aligned Word-PDF vs Oxi) localized the
            // 賃金 chapter over-fit to 102 "Word-NAT / Oxi-½" mid-、 — Oxi compresses
            // 、 at break, Word breaks at natural (compressing only on justify-slack).
            // natural_break_jc/s557 cover only !is_justified / c15; legacy JUSTIFIED
            // (compat<15, type=lines) was uncovered → ×0.6667 fired. compat=11 docs:
            // tokyoshugyo. See [[char_budget_wall]], [[tokyoshugyo_wrap_not_cellheight]].
            let s589_legacy_just_natural = std::env::var("OXI_S589").ok().as_deref() == Some("1")
                && is_justified
                && self.compress_punctuation
                && self.compat_mode < 15
                && !lines_and_chars;
            // S557 (2026-06-13, part of the OXI_S556_JUST15 opt-in scaffold):
            // c15-explicit JUSTIFIED paragraphs keep standalone 、。，．at
            // NATURAL width at break — Word defers ALL their compression to
            // the per-line pack decision (d77a para9-L6 ground truth: Word's
            // 38-char line = naturals overflowing 2.5 → packed −0.75×3 onto
            // 、）、; Oxi's legacy ×0.6667 pre-compress baked −4/punct into
            // width AND natural_width, blinding the pack tier's need (7.0 vs
            // Word's 14.5 for the 39th char) and its all-natural guard).
            let s557_natural_just15 = std::env::var("OXI_S556_JUST15").is_ok()
                && is_justified
                && self.compress_punctuation
                && self.compat_mode >= 15
                && self.compat_mode_explicit
                && !lines_and_chars;
            // S475 capacity-budget break (env-gated, default OFF = byte-identical).
            // Greedy first-fit where each punct contributes break-compression CAPACITY
            // (pair-first 6.0 / solo 1.5, env-tunable; flat-K = equal). Bypasses the
            // ×0.6667 standalone pre-compress + the S472/S473 absorb, and routes render
            // through the s472_render water-fill so glyphs justify correctly.
            // SCOPED to NO-char-grid sections only (docGrid type=lines / none →
            // grid_char_pitch is None). docGrid type=linesAndChars (grid_char_pitch
            // Some, e.g. b837) is GRID-determined (fixed char count/line) — a
            // SEPARATE mechanism (charGrid charsLine); S475 yakumono capacity must
            // NOT apply there (it cascaded b837 7→9 pages). See session471 finding.
            // S475 SHIPPED default-ON (2026-06-01, opt-out OXI_S475_DISABLE). flatK
            // params (PAIR=SOLO=2.5) — the break-decision capacity; reproduces d77a
            // [39,38,40,41,…]-class packing on type=lines docs. Gate: Phase-1 54/55
            // (no PASS→FAIL, b837 7pg), SSIM net +0.0398, bottom-5 +0.0109 (d77a
            // +0.062, ed025c +0.097). c7b923 −0.036 = latin-mixed residual (lever B).
            // S476 (2026-06-02): extend the S475 yakumono capacity break to
            // linesAndChars docs' MAIN BODY (b837/b35/tokumei = corpus bottom-N).
            // Lever C count-cap was FALSIFIED (Oxi charsLine already = Word's; Word
            // does NOT cap; Oxi UNDER-packs linesAndChars). The real gap is the SAME
            // yakumono per-line packing as S475 — Word fits more per line by
            // compressing punct. linesAndChars needs a heavier cap (K≈3.0 vs lines'
            // 2.5; the char-grid context). COM-verified Phase-1-safe: the whole
            // linesAndChars family (b837/b35/1636/31420/6514/a1d6/87b29/29dc6e/1ec1)
            // keeps its baseline=Word page count at K=3.0. b837 +0.0455. aux/cell
            // calls (s476_body=false) stay excluded to avoid the 7→9 cell cascade.
            // S568 (2026-06-14): LEGACY (compat ≤14) linesAndChars compressPunctuation
            // docs apply jc=left 約物 OIKOMI (the s476 capacity break) that the
            // compat≥15 gate excludes. harassmanual (compat=11) orphans a trailing
            // char (く) that Word fits by compressing a mid-line 読点 、 to half-em
            // (COM _s568_p16_adv: 、 advance 6.0pt, every other char 12.0). The
            // discriminator is compat: modern (15) jc=left breaks at NATURAL widths
            // (S492/S539 measured + shipped), legacy demands oikomi (see the
            // compat_mode_explicit note at mod.rs:1512). The ONLY compat<15
            // linesAndChars compressPunctuation doc in the corpus is harassmanual
            // (compat=14 docs are type=lines, not linesAndChars), so this is a
            // single-doc-scoped change. Cap = full half-em (6.0). Opt-out OXI_S568_DISABLE.
            let s568_legacy_oikomi = std::env::var("OXI_S568_DISABLE").is_err()
                && lines_and_chars
                && s476_body
                && self.compress_punctuation
                && self.compat_mode < 15;
            // S1234 (2026-08-26, default ON, opt-out OXI_S1234_DISABLE): the S568
            // legacy-oikomi cap is SIZE-REGIME-dependent. A para whose runs sit ON
            // the doc-default size keeps the S568 half-em cap (harassmanual's 、→6.0
            // pack). A para OFF the default regime (parttime 就業規則本文: 8pt vs
            // docDefaults 12pt) gets only the standard light capacity (3.0): Word
            // bills its marks NATURAL at break (第25条 stays 3 lines) yet still
            // rescues small overflows (第24条: ・6.84/）5.66 on L1, line-end 。4.26
            // on L2 → 2 lines). |Δ| ≥ 1.5pt = the S1231 regime threshold.
            // S1237 (2026-08-27, default ON, opt-out OXI_S1237_DISABLE): the
            // AT-default regime gets NO mid-mark oikomi credit at all — probe C
            // (、×5, demand 3.55, natural 50) and slice v5 (sz21, demand 5.2,
            // M=2, natural) both REFUSE where the cap-6 branch packs. With the
            // opt-out set, at-default counts as ABOVE (the pre-S1237 cap-6).
            let s1236_regime_delta = fragments
                .iter()
                .find(|(t, _, _, _, _)| {
                    t.chars().any(|c| !c.is_whitespace() && c != '\u{3000}')
                })
                .map(|(_, rs, _, _, _)| {
                    self.resolve_font_size(rs, para_style) - self.doc_regime_fs
                });
            // S1269 (2026-09-01, default ON, opt-out OXI_S1269_DISABLE): S1237
            // only applies to the compat regime it was DERIVED in. Its evidence
            // — probe C and the harassmanual slice — is compat 11; the S568 gate
            // it rides on admits compat < 15, and S568's own note justified its
            // width with "the ONLY compat<15 linesAndChars compressPunctuation
            // doc in the corpus is harassmanual ... compat=14 docs are
            // type=lines, not linesAndChars". Both halves of that are now false:
            // `_s568_gate_census.py` finds **12** docs through the gate, **9 of
            // them compat 14**. jaBlindB50 joined dev on 2026-08-31 and brought
            // them in.
            //
            // On its derivation regime S1237 changes NO page count at all
            // (harassmanual 4=4, parttime 7=7, 9e4d04b4 6=6 in every arm — its
            // evidence was line-level). On compat 14 it costs pages:
            //
            //   doc         Word  S1237 ON  S1237 OFF
            //   0ea3ec86     43     45        43 =
            //   167853753    29     33        30
            //   0b6f3b32     25     25 =      24
            //   sum|pcd| over the 8 scored gate docs:  7  ->  3
            //
            // Turning S1237 off wholesale was NOT taken: its law is measured.
            // Narrowing it to compat < 14 keeps the measured behaviour where it
            // was measured and stops it acting on a population it was never
            // tested against — a scope correction, not another regime carve-out
            // (the Ra no-EXCEPTION-stacking rule: S1234 -> S1236 -> S1237 are
            // already three cuts on this one gate).
            //
            // ★ 0b6f3b32 loses its exact 25 (-> 24) and is left NAMED as the
            // next target rather than bought back with a fourth cut. It is a
            // 25-page doc moving by one page — the +-1 knife-edge the pcd-first
            // rule calls a sub-pt noise floor.
            let s1269_derived_regime_only = std::env::var("OXI_S1269_DISABLE").is_err();
            // S1318 (2026-09-05, default ON, opt-out OXI_S1318_DISABLE): the
            // at-default refusal holds at compat 14 as well. MEASURED with
            // `_pb_oikomi_default_gen.py` (03ca64d7's first body line rebuilt:
            // 41 chars on a 40-cell floor, ＭＳ 明朝 10.5 = docDefaults 10.5,
            // linesAndChars 298, compressPunctuation): Word keeps 40 and every
            // mark at 1.00em for M = 1, 2, 3, 4 marks (demand 10.5 .. 2.6 per
            // mark), with or without kern / balance, in a cell, and at compat
            // 11; only compat 15 packs (41, marks 0.74). The real doc agrees
            // (25 mid-line marks at 10.5pt, 0 compressed) and so do 0728f6dd
            // (0/4) and 0422f651 (0/5). S1269 had narrowed the refusal to
            // compat < 14 on page counts alone; the body regime it applies to
            // is the same at 14. With the opt-out set, compat 14 falls back to
            // the S1269 narrowing (cap-6 packing).
            // HELD 2026-09-05 (opt-in OXI_S1318=1): the JA blind A/B REGRESSED
            // 89 -> 88, sum|pcd| 6 -> 10 -- 0b6f3b32 (25 pages, 0.2566),
            // 167853 (pcd +1 -> +2), 0ea3ec86 (0.73 -> 0.12, pcd +2): the same
            // three docs S1269 measured. Their Word PDFs DO compress 7-12% of
            // the at-default 11pt marks (88/891, 36/531, 22/186), so the
            // regime the probe measured (plain body paragraph, one column) is
            // not theirs; the discriminator is still open (column / cell
            // width, numbering, hang). See the S1318 archive note.
            let s1318_c14_refuse = std::env::var("OXI_S1318").ok().as_deref() == Some("1");
            // S1318 v2 (2026-09-05, default ON, opt-out OXI_S1318_DISABLE): the
            // at-default refusal is decided PER CHARACTER, by the character
            // that would overflow. `_pb_oikomi_default_gen.py` (33 arms): a
            // normal next character is refused at every compat <= 14 (also
            // c11), but a LINE-FINAL 、。 (kinsoku forbids it at a line start)
            // is pulled in by compressing the line's mid marks evenly, each
            // by at most half (end_m1 0.51, end_m2 0.50/0.50, end_m3 0.67 x3;
            // end_m0 hangs). That is the S568 cap-half machinery -- kept for
            // that character class -- while S1237's paragraph-level refusal
            // (and the S1269 compat narrowing that patched its damage on the
            // 2-column docs, whose packed lines are all line-final marks)
            // becomes the per-character rule below (`s1318_refuse_here`).
            // HELD as opt-in (OXI_S1318=1) 2026-09-05: line-level correct on
            // every probe arm and on 0ea3ec86 / 0b6f3b32's paragraphs (their
            // clean 2-column lines hold 20 chars like Word; the old cap-6 path
            // packed 21-25), yet the JA blind page counts regress because that
            // over-packing was compensating a VERTICAL surplus elsewhere in the
            // same docs (0b6f3b32 25->26 pages from +16 lines; 0ea3ec86 43->45).
            // Ships when that partner is identified (Ra: no ship on a negative
            // gate without the compensating error named).
            // S1318 v3 (2026-09-06, default ON, opt-out OXI_S1318_DISABLE): the
            // refusal is not absolute -- a normal character IS pulled in when
            // the line overflows by at most HALF A CELL and has a mark to
            // absorb it (the marks share the demand; a lone 、 gives up to its
            // whole 0.5em aki, two marks 0.25 each). DERIVED with
            // `_oikomi_census.py` on Word's own PDFs (every body line of three
            // two-column compat-14 docs + the 06ee35d5 cell): 0ea3ec86 28/29
            // grants + 493/493 refusals, 167853 6/7 + 351/351, 0b6f3b32 5/5 +
            // 120/120 agree; the two "beyond" grants are lines that would
            // otherwise end in an OPENING bracket (kinsoku forces the pull-in,
            // every mark 0.5). A line-final 、。 or closing bracket keeps the v2
            // cap-half + hang path. The v2 "refuse every normal character" left
            // 0ea3ec86 at W43/O45 -- lines Word packs to 21 (e.g. p3 col1
            // 「介します。（各…等は各」, demand 0.49 cell over 。 and ・) broke a
            // character early, and every such line cascades.
            // HELD OPT-IN again 2026-09-06 (`OXI_S1318=1`): the JA blind A/B of v3
            // read 93/93, mean 0.9952 -> 0.9850 (0ea3ec86 W43/O45, 167853 W29/O30;
            // 0b6f3b32 1.0 either way). The normal-character half-cell rule holds
            // on every witness; what is still wrong is the LINE-FINAL unit (X + 、。):
            // Word refuses 「お、」 with （）。 on the line (2.0 cells of marks) and
            // 「事、」 with 、、, yet keeps 「た。」 by halving three brackets. The
            // faithful-slice probe `_pb_kinsokufinal_gen.py` sweeps it.
            // S1346 (2026-09-07): default ON (was opt-in OXI_S1318=1), opt-out
            // OXI_S1318_DISABLE. JA blind 100: PASS 93 = 93, mean 0.9952 -> 0.9953,
            // the only two documents whose bytes move both improve (0ea3ec86
            // 0.9957 -> 0.9983, 167853 0.9967 -> 0.9989); Phase 1 96/96; the five
            // faithful-slice probes (unitcap 660 lines, unitcap18 631, unitcap5
            // 436, trackedge 84, kinsokufinal 51) agree on every break.
            let s1318_v2 = std::env::var("OXI_S1318_DISABLE").is_err()
                && std::env::var("OXI_S1318").ok().as_deref() != Some("0");
            // S1490 (2026-09-19, default ON, opt-out OXI_S1490_DISABLE): a compat-15
            // compressPunctuation body without a linesAndChars grid (the 大野 /
            // NEDO / d77a family, docGrid type=lines, jc=both) breaks under the
            // regime's machinery with its own capacity law (`compcap.py`, 18
            // arms on the ohnoshugyo slice, kinds 、。）・（ identical):
            //   k marks on the line -> up to k/(k+1) cell of overflow (-0.045),
            //   each mark shrunk by 1/(k+1): 0.452 / 0.624 / 0.714 / 0.762.
            // S475's summed 3.25pt caps under-pack at k=1 (0.31) and over-pack
            // at k>=3 (0.93, 1.24) -- the census pattern on all nine witnesses.
            // A non-justified paragraph breaks at natural width (S492):
            // technical__978ec9c102290205 (jc=left x176) wraps at 0.06 cell
            // with five marks on the line.
            let s1490_regime = std::env::var_os("OXI_S1490_DISABLE").is_none()
                && !narrow_punctuation_natural
                && !natural_break_jc
                && !vertical
                && (!lines_and_chars || (is_justified
                    && cjk_metrics.is_some_and(|m|
                        ['\u{3001}', '\u{3002}'].iter().all(|&ch| m.char_widths.contains_key(&ch) && m.char_width_em(ch) >= 0.99))
                    && style.east_asia_lang.as_deref().is_some_and(|lang|
                        lang.eq_ignore_ascii_case("ja") || lang.to_ascii_lowercase().starts_with("ja-"))))
                && s476_body
                && self.compress_punctuation
                && self.compat_mode >= 15
                && self.compat_mode_explicit;
            let s1318_at_default_regime = (s568_legacy_oikomi
                    && s1236_regime_delta.map_or(false, |d| d.abs() < 1.5))
                || s1490_regime || japanese_language_oikomi;
            let s1318_at_default_regime = s1318_at_default_regime && !english_ea_natural && !narrow_punctuation_natural;
            // S1346: a Latin word overflowing the floor is rescued by the same
            // elective half-cell as a CJK character (`_pb_unitcap_gen.py`: 19 kana
            // + 、 + 「111」 keeps 111 with the 、 at 5.8; 「…セ）12」 「…）123」
            // 「…の80」 and the tracked 「及び（エオ）…ナニ1」 likewise), so
            // flush_word gets half the grid cell when the line holds a mark.
            s1346_regime_credit.set(if s1318_v2 && s1318_at_default_regime {
                match (grid_char_cw_ratio, grid_char_pitch) {
                    (Some(ratio), Some(pitch)) if ratio > 0.0 && pitch > 0.0 => pt_to_tw(0.5 * pitch),
                    _ => 0,
                }
            } else {
                0
            });
            let full_word_credit = s476_body && self.compress_punctuation
                && grid_char_pitch.is_none() && !english_ea_natural && !narrow_punctuation_natural
                && (self.compat_mode <= 14 || is_justified);
            word_full_punctuation_credit.set(full_word_credit);
            if full_word_credit {
                s1346_regime_credit.set(pt_to_tw(font_size * 0.5));
            }
            let s1237_at_default_refuse = !s1318_v2
                && std::env::var("OXI_S1237_DISABLE").is_err()
                && s568_legacy_oikomi
                && (!s1269_derived_regime_only
                    || self.compat_mode < 14
                    || (s1318_c14_refuse && self.compat_mode == 14))
                && s1236_regime_delta.map_or(false, |d| d.abs() < 1.5);
            let s568_legacy_oikomi = s568_legacy_oikomi && !s1237_at_default_refuse;
            let s1234_offdefault_light = std::env::var("OXI_S1234_DISABLE").is_err()
                && s568_legacy_oikomi
                && fragments
                    .iter()
                    .find(|(t, _, _, _, _)| {
                        t.chars().any(|c| !c.is_whitespace() && c != '\u{3000}')
                    })
                    .map(|(_, rs, _, _, _)| self.resolve_font_size(rs, para_style))
                    // S1236 (2026-08-27): DIRECTIONAL + regime-referenced. The
                    // harassmanual slice ablation (v4/v5) pinned the oikomi
                    // regime to run size vs the grid default (docDefaults else
                    // Normal — font-irrelevant: MS Mincho halves too at sz24):
                    // ABOVE the default (12 > 10.5) Word demand-compresses
                    // mid 、to fs/2 (= the s568 cap-6 branch, kept); AT the
                    // default it refuses (probe C 3.55/M5, slice v5 5.2/M2);
                    // BELOW it (parttime 8 < 12) the S1235 small caps apply.
                    .map_or(false, |fs| fs - self.doc_regime_fs <= -1.5);
            // S592 (2026-06-17): a PROPORTIONAL CJK font (pgothic family) in a
            // linesAndChars grid is OFF-GRID, so the s476 capacity break must NOT
            // fire — its 約物 are already at the font's narrow proportional advance
            // (HGPGothicM 、 = 6.75pt, no fullwidth aki to remove), so Word does NOT
            // demand-compress them at break time (it justifies via inter-char
            // expansion instead — S579 "Word adds ~0.2pt justify on top"). kojin
            // (justified via docDefaults jc=both) was crediting ~5.2pt of phantom
            // 約物 compression to fit a trailing こ Word WRAPS (OXI_DBG_KOJIN:
            // over_tw=−26 capacity vs +78tw actual). Excluding para_off_grid drops
            // s475_break → natural break → こ wraps → para i297 4 lines = Word.
            let s476_grid = (std::env::var("OXI_S476_DISABLE").is_err()
                && lines_and_chars
                && s476_body
                && self.compress_punctuation
                && self.compat_mode >= 15
                && !para_off_grid)
                || s568_legacy_oikomi
                || s572_legacy_notype_oikomi;
            // S590 (2026-06-16, opt-IN OXI_S590=1, default OFF = byte-identical):
            // LEGACY (compat<15) JUSTIFIED body paras use the s475 CAPACITY break
            // (greedy + compress-to-fit only when overflow ≤ Σ約物-caps, cap≈2.5)
            // instead of the flat ×0.6667 pre-compress (−3.5pt, over-compresses) OR
            // S589 pure-natural (0, under-fits the real oikomi lines). DERIVED:
            // _tks_oidashi.py --absorb on the Word PDF — Word's per-約物 oikomi
            // compression caps at ~2.9pt (median 1.93), and only 14/219 full lines
            // compress (176 expand at natural 約物). So the capacity model with
            // cap≈2.5 (≈ Word max) reproduces Word's compress-14/expand-176 split,
            // unlike ×0.6667 (over) / S589-natural (under) / S543-fs/2 (way over).
            // S689 (2026-06-29, SHIPPED default ON, opt-out OXI_S590_DISABLE): the body
            // half of the tokyoshugyo joint-solve. A LEGACY (compat<15) JUSTIFIED
            // type=lines compressPunctuation body para uses the s475 CAPACITY break
            // (greedy + demand 約物 compression, cap solo 1.5 / pair 6.0) instead of the
            // flat ×0.6667 pre-compress (which OVER-compresses 約物 → fits ~1 extra
            // char/line → −1 page drift). DERIVED S590 (2026-06-16): _tks_oidashi --absorb
            // shows Word compresses RARELY (14/219 full lines) at cap ~2.5 max; the
            // capacity model reproduces the compress-14/expand-176 split. ★Shipped now
            // (was opt-in OXI_S590=1): the memory kept it opt-in citing "page count
            // 90→91", but that PREDATED S591/S585b (cell-clamp) being default-ON — with
            // the cells clamped, the body capacity break keeps tokyoshugyo at 90 pages
            // (= Word) AND improves the reliable page-top metric 41→36 (gate 0.9817→
            // 0.9824, one −1 fixed). ★SINGLE-DOC-SCOPED by construction (compat<15 +
            // justified + type=lines + compressPunctuation = tokyoshugyo ALONE in the
            // corpus; verified byte-identical for gen/gen2/test + all others). The CELL
            // under-compression (the regulation boxes' jc=left wrapper does ZERO 約物
            // compression) is the SEPARATE joint-solve piece, still pending.
            let s590_legacy_just_cap = std::env::var("OXI_S590_DISABLE").is_err()
                && is_justified
                && self.compress_punctuation
                && self.compat_mode < 15
                && !lines_and_chars;
            let s475_break = ((std::env::var("OXI_S475_DISABLE").is_err()
                && self.compress_punctuation
                && self.compat_mode >= 15
                && !lines_and_chars)
                || s476_grid
                || s590_legacy_just_cap)
                && !natural_break_jc && !natural_punctuation_boundary; // S492: non-justified paras break at natural
            let s476_cap: f32 = std::env::var("OXI_S476_CAP")
                .ok()
                .and_then(|v| v.parse().ok())
                .unwrap_or(
                    if (s568_legacy_oikomi && !s1234_offdefault_light)
                        || s572_legacy_notype_oikomi
                    {
                        6.0
                    } else {
                        3.0
                    },
                );
            // S558 (2026-06-13): s475_pair default 2.5 → 6.0. A CLOSING bracket
            // before another bracket collapses a full half-em at break (matching
            // the render pair-halving); the old 2.5 under-counted it, so
            // bracket-cluster justified lines (d77a para9 L3 ）」（) broke a char
            // early — an SSIM cascade. Comma/period-first pairs still trim lightly
            // (s475_max_compress split — see kinsoku.rs). Env-tunable.
            // S590 refinement (2026-06-16): per-TYPE caps — solo (、。，．) = 1.5
            // (the derived break demand), but bracket-PAIR clusters keep 6.0 (S558,
            // the lever-3 heavy cluster compression). Measured (S591 cells clamped,
            // break divergence): solo1.5/pair1.5=593 → solo1.5/pair6.0=542 (best).
            let s475_pair: f32 = if s476_grid {
                s476_cap
            } else {
                std::env::var("OXI_S475_PAIR")
                    .ok()
                    .and_then(|v| v.parse().ok())
                    .unwrap_or(6.0)
            };
            // S575 (2026-06-15): BODY oikomi — raise the solo 約物 cap to 3.0 for the
            // MAIN body flow (s476_body) so jc=both type=lines compat=15 bodies fit
            // Word's demand compression (ikujikaigo i=41/i=57: mid 、 renders 9.0 = −3.0;
            // +1×4 → PASS). GATED to paras WITHOUT a lastRenderedPageBreak: the higher cap
            // REDISTRIBUTES chars within a para's lines (line COUNT unchanged), which SHIFTS
            // which line a run's char_offset==0 lands on → the S391 per-line-LRPB respect
            // then fires on a DIFFERENT line → a spurious mid-para page break (d77a's
            // "イは、編集…" para, 1 LRPB: +18pt continuation cascade → cell over page 6/7 =
            // the 16-session d77a blocker, isolated via OXI_S391_PER_LINE_LRPB=0 + env
            // bisection). ikujikaigo i=41/i=57 have 0 LRPBs → safe to redistribute. The
            // break itself is solo-STABLE (count unchanged); only the LRPB attribution is
            // sensitive, so skip the oikomi when there's an LRPB to preserve. Opt-out
            // OXI_S575_DISABLE.
            // S590: derived break-time 約物 demand cap ≈ 1.5pt (sweep minimum,
            // _tks_oidashi: divergence 666@0 → 618@1.5 → 925@2.5). Word's BREAK
            // cap (~1.5) < its RENDER cap (~2.9) — break is conservative, render
            // (justify) compresses more. Env OXI_S475_SOLO overrides.
            // OXI_S575_CAP: targeted sweep of the NON-LRPB body 約物 cap ONLY
            // (keeps the LRPB branch at 2.5 so the S391 redistribution gate stays
            // conservative — unlike OXI_S475_SOLO which overrides every branch).
            // S604 (2026-06-18): default raised 3.0 → 3.1. The body 約物 demand cap
            // 3.1 (+line-end ぶら下げ S601) matches Word's oikomi line counts for the
            // ohno-family regulation docs (matsuiikuji/ohnochingin FAIL→PASS) and the
            // 3a4f/model paras 69/173/294 (Word 5/1/6 lines = Oxi cap-3.1; cap-3.0
            // under-fit by 1 line each). The cap-3.1 −1 it once caused on 3a4f/model
            // (a typed-grid page-bottom compensating error at para278) is now fixed by
            // S603 (page-bottom full-cell before a table). Phase-1 73→75, 0 PASS→FAIL.
            // S607 (2026-06-18, ATTEMPTED 3.1→3.4, REVERTED to 3.1): the body solo 約物
            // break cap 3.1 (= 0.258em, scaled by fs/12 in s475_max_compress) under-packs
            // 約物 lines vs Word on the MEASURED paras — nedocontract (all-12pt regulation
            // doc) para41 fits 34/Word 35 (Word compresses each mid-line 、 by 3.36pt to
            // fit the trailing char; cap 3.1 gives 3.1pt → "お" orphaned → +1×15) AND
            // model/3a4f para301 fit 38/Word 39. cap 3.4 = the minimal value reproducing
            // Word's char counts there (nedocontract 0.9688→0.9938, +1×15→−1×3). ★BUT the
            // raise NET-REGRESSES SSIM: the correct ssim_ab.py A/B (cap 3.1 vs 3.4, DWrite
            // 235 bases) = net −0.0261, kyodokenkyuyoushiki05 −0.0252 (the 約物 redistribute
            // OIKOMI where Word OIDASHI'd — the same oikomi/oidashi wall as nedocontract's
            // own residual −1×3). nedocontract has NO word_png so its pagination gain is
            // SSIM-untracked, and it does NOT pass either way → the trade is a tuned-doc
            // SSIM loss for a non-passing pagination gain. ★The original S607 commit
            // (e71add4c) claimed "SSIM net +0.0000" but that was measured with a BROKEN
            // ssim_ab tool (calculate_ssim called with the wrong signature → every call
            // errored → false 0.0000); the corrected tool exposed the real −0.0261. So
            // the default reverts to 3.1; cap 3.4 stays reachable via OXI_S575_CAP=3.4 for
            // the nedocontract/model case. Matching Word's per-line oikomi/oidashi (not a
            // single cap) is the char-budget wall. The S604 default 3.0→3.1 stands.
            // S639 (2026-06-21, OPT-IN OXI_S639=1, default OFF = byte-identical):
            // body solo 約物 cap 3.1→3.4 + OPENING-bracket cap 3.0 (< solo). The cap
            // 3.1 UNDER-compressed marks (、。) → nedocontract +1×15; cap 3.4 fixes it,
            // BUT a bare 3.4 over-credited OPENING brackets on 約物-dense lines → −3
            // over-fit + the S607 kyodoken05 −0.0261. The s475_open=3.0 cap removes the
            // OPENING over-credit: nedocontract pagination 0.9688→0.9979 (0 PASS→FAIL,
            // n_pass 81→81) AND kyodoken05 SSIM 0.9705→0.9845 (+0.0140). ★HELD OPT-IN:
            // the verified SSIM A/B (ssim_ab.py OXI_S639) shows 3a4f −0.0057 (a passing
            // canary's RENDER regresses — pagination unchanged) because cap 3.4 also
            // OVER-compresses 3a4f's MARKS, where Word's per-line demand is <3.4 (the
            // demand-proportional wall; the opening fix addresses opening over-credit,
            // NOT mark over-compression on lower-demand docs). Net SSIM +0.0080 but a
            // canary regresses → default OFF until the demand-proportional break (compress
            // only as NEEDED) replaces the flat cap. OXI_S575_CAP / OXI_S475_OPEN override.
            // S639b (2026-06-22, default ON, opt-out OXI_S639_DISABLE): the body
            // oikomi break DECOUPLES the 、。/closing cap (3.4) from the opening-bracket
            // cap (3.1). The body solo 、。/closing cap 3.1 (S604) UNDER-compressed marks
            // → nedocontract +1×15 (Word fits a trailing char by compressing each mid-
            // line 、 ~3.4pt @12pt = its render demand p10; cap 3.1 under-credited →
            // orphaned char → +1). Raising it to 3.4 fixes nedo (0.9688→0.9979) AND
            // CLEARS the S607 kyodoken05 cap-3.4 regression (+0.0140 — the −0.0252 S607
            // saw was the OPENING over-credit, not the mark cap). The opening-bracket
            // cap stays at the pre-S639 body value 3.1 (Word trims an opening bracket's
            // LEFT aki only, ~3.1pt — LESS than 、。/closing which have full right aki).
            // ★The original S639 (opt-in) set opening 3.0, which was too LOW: 3a4f p4
            // para28 L0 «…昭和２２年厚生省令» (2 opening （, overflow ~5.3pt @10.5)
            // needs 2.65pt/（ to fit 令 like Word; open 3.0 scaled (2.625@10.5) dropped
            // 令 → a 41→40 over-WRAP that regressed 3a4f p4 SSIM −0.0057. open=3.1 fits
            // 41 (=Word) AND keeps nedo's −3 fixed (nedo's −3 over-fit only reappears at
            // open≥3.4=solo; the clean window is open∈[3.04,3.3]). So 3a4f recovers to
            // +0.0000 while nedo stays 0.9979 + kyodoken05 +0.0140. Scoped to the
            // body-oikomi context (solo=3.4): non-body/LRPB/grid/legacy paras keep
            // opening = s475_solo (unchanged). OXI_S575_CAP / OXI_S475_OPEN override.
            let s639 = std::env::var("OXI_S639_DISABLE").is_err();
            let s575_body_cap: f32 = std::env::var("OXI_S575_CAP")
                .ok()
                .and_then(|v| v.parse().ok())
                .unwrap_or(if s639 { 3.4 } else { 3.1 });
            let s475_solo_default = if s590_legacy_just_cap {
                1.5
            } else if s476_body && !para_has_lrpb && std::env::var("OXI_S575_DISABLE").is_err() {
                s575_body_cap
            } else {
                2.5
            };
            let s475_solo: f32 = if s476_grid {
                s476_cap
            } else {
                std::env::var("OXI_S475_SOLO")
                    .ok()
                    .and_then(|v| v.parse().ok())
                    .unwrap_or(s475_solo_default)
            };
            // S721 orphan re-break pass: escalate the 約物 caps (see the caller's
            // two-pass comment — accepted only when it saves a line).
            let s721_retry = S721_ORPHAN_RETRY.with(|f| f.get());
            let s475_solo = if s721_retry {
                // 4.4 code-scale (s475 caps scale by fs/12) = 3.85 effective @10.5pt,
                // covering the ③-measured 3.78 demand; the (注) L1 (Word declines,
                // demand 7.75) still wraps → its re-break saves no line → first pass
                // kept. OXI_S721_CAP to sweep.
                s475_solo.max(
                    std::env::var("OXI_S721_CAP")
                        .ok()
                        .and_then(|v| v.parse().ok())
                        .unwrap_or(4.4),
                )
            } else {
                s475_solo
            };
            // S639b: opening-bracket break cap (< solo) in the body-oikomi context only
            // (where solo=3.4). Elsewhere = s475_solo (byte-identical to pre-S639b).
            let s639_body_oikomi = s639
                && s476_body
                && !para_has_lrpb
                && !s476_grid
                && !s590_legacy_just_cap
                && std::env::var("OXI_S575_DISABLE").is_err();
            let s475_open: f32 = std::env::var("OXI_S475_OPEN")
                .ok()
                .and_then(|v| v.parse().ok())
                .unwrap_or(if s639_body_oikomi { 3.1 } else { s475_solo });
            // S721 orphan re-break: openings escalate too (nedo 子 needs （ at 3.31).
            let s475_open = if s721_retry {
                s475_open.max(
                    std::env::var("OXI_S721_OPEN")
                        .ok()
                        .and_then(|v| v.parse().ok())
                        .unwrap_or(3.4),
                )
            } else {
                s475_open
            };
            // S645 (2026-06-23) FALSIFIED + reverted: a demand-aware closing-bracket
            // cap (」/） before a non-bracket → light ~0.84, Word-measured) did NOT
            // fix nedo's cumulative over-pack (still −1×3 at close 1.0-2.0; the
            // over-pack is not a per-約物-type cap issue) AND regressed b837 1.0→0.9859
            // (the closing cap is canary-load-bearing for b837's footnote/linesAndChars).
            // Combined with n_period (openings), still no clean nedo fix → the caps are
            // inter-doc-coupled; nedo needs Word's reflow, not a cap model. See [[char_budget_wall]].
            let s473_locomp = std::env::var("OXI_S473_LOCOMP").is_ok();
            let s473_cap: f32 = std::env::var("OXI_S473_CAP")
                .ok()
                .and_then(|v| v.parse().ok())
                .unwrap_or(3.25);
            // S473b (2026-06-01): per-type break caps. Render-advance COM showed
            // Word compresses brackets/。 ~4× more than 、 (）=remove 6.0 vs 、=remove
            // 1.5). A UNIFORM cap could not reconcile d77a (bracket-heavy lines, want
            // heavy) with b837 pi30 (、-only line, wants light). Per-type caps (env-
            // tunable for the sweep): comma/opening-bracket = light, period/closing-
            // bracket = heavy. Defaults from render data (1.5 / 6.0). Used only when
            // OXI_S473_ASYM is set (else the uniform s473_cap path runs).
            let s473_asym = std::env::var("OXI_S473_ASYM").is_ok();
            let s473_cc: f32 = std::env::var("OXI_S473_CC")
                .ok()
                .and_then(|v| v.parse().ok())
                .unwrap_or(1.5); // 、，
            let s473_cp: f32 = std::env::var("OXI_S473_CP")
                .ok()
                .and_then(|v| v.parse().ok())
                .unwrap_or(6.0); // 。．
            let s473_ccl: f32 = std::env::var("OXI_S473_CCL")
                .ok()
                .and_then(|v| v.parse().ok())
                .unwrap_or(6.0); // closing brackets
            let s473_cop: f32 = std::env::var("OXI_S473_COP")
                .ok()
                .and_then(|v| v.parse().ok())
                .unwrap_or(1.5); // opening brackets
            let s472_demand = std::env::var("OXI_S472_DEMAND").is_ok() || s473_locomp;
            let chars_vec: Vec<char> = text.chars().collect();
            // Complex-script (Devanagari) width: the per-codepoint sum this loop
            // uses is wrong for virama conjuncts / reordered matras (MEASURED
            // 2× too wide on क्ष; see font::shape). Shape the whole fragment once
            // and use the cluster advances. The shaping font is the one Word AND
            // DirectWrite actually draw with: the run's own family if it is
            // installed and covers Devanagari, else Nirmala UI (Word's Devanagari
            // fallback, confirmed by reading the fonts out of Word's own PDF for
            // igrsup_md_v1/v9/v11). Entered ONLY when the fragment carries a
            // complex-script char, so the frozen Latin/CJK corpus never reaches
            // it (byte-identical by construction).
            let deva_adv: Option<(Vec<f32>, Vec<bool>)> = if std::env::var("OXI_DEVA_DISABLE")
                .is_err()
                && chars_vec.iter().any(|&c| crate::font::is_complex_script(c))
            {
                let emit_fam = self
                    .resolve_font_family_for_text(text, style, para_style)
                    .map(|s| s.to_string());
                let shape_fam = match emit_fam {
                    Some(f) if crate::font::shape::family_covers(&f, '\u{0915}') => f,
                    _ => "Nirmala UI".to_string(),
                };
                crate::font::shape::cluster_advances(
                    &shape_fam, style.bold, style.italic, text, font_size,
                )
            } else {
                None
            };
            // Yakumono pair compression for line break width calculation.
            // Rule 1 (close+open ×0.5) is gated by yakumono_pair_enabled
            // (compress_punctuation OR hwid font); Rules 2-4 below use
            // yakumono_enabled (compress_punctuation only).
            let yakumono_compressed: Vec<bool> = if yakumono_pair_enabled {
                let n = chars_vec.len();
                let mut v = vec![false; n];
                // S1494 (2026-09-20, default ON, opt-out OXI_S1494_DISABLE): a pair
                // is a pair across a RUN boundary too. technical__c5bb0090235dfedb
                // keeps （ / 「キャッシュ… / 、 / 「令和… in separate runs; Word halves
                // the 、 (5.25) and one bracket of （「 (COM Information(5)), Oxi
                // saw every mark alone (10.16) and lost a character per line --
                // two lines became three. Minimal repro (pair_1run/pair_2run):
                // the same text in one run halves 」、。, split into runs only 」
                // (its partner stayed in-run) does.
                let s1494 = std::env::var_os("OXI_S1494_DISABLE").is_none();
                let next_frag_first: Option<char> = if s1494 {
                    fragments[frag_outer_idx + 1..].iter().flat_map(|f| f.0.chars()).next()
                } else { None };
                let prev_frag_last: Option<char> = if s1494 {
                    fragments[..frag_outer_idx].iter().rev().flat_map(|f| f.0.chars().rev()).next()
                } else { None };
                let next_of = |i: usize| -> Option<char> {
                    if i + 1 < n { Some(chars_vec[i + 1]) } else { next_frag_first }
                };
                let prev_of = |i: usize| -> Option<char> {
                    if i > 0 { Some(chars_vec[i - 1]) } else { prev_frag_last }
                };
                for i in 0..n {
                    let c = chars_vec[i];
                    if kinsoku::is_yakumono_closing(c) {
                        if next_of(i).map_or(false, kinsoku::is_yakumono_trigger) {
                            v[i] = true;
                        }
                    } else if kinsoku::is_yakumono_opening(c) {
                        // S1217 (2026-08-25, opt-out `OXI_S1217_DISABLE`): an opening
                        // bracket followed by ANOTHER opening bracket compresses --
                        // and it is the FIRST of the pair that gives up its half em,
                        // exactly like the closing side. MEASURED with
                        // `tools/metrics/_pb_yakuwidth_gen.py` (10 arms, jc=left short
                        // lines, glyph origins out of Word's own PDF): 「（（」 puts the
                        // brackets at 106.22 and 111.39 -- a 5.164pt advance on the
                        // first, half of the 10.56 every solitary mark gets. The same
                        // reading at 9pt under charSpace=-2714 gives 4.077 of 8.337.
                        if std::env::var("OXI_S1217_DISABLE").is_err()
                            && next_of(i).map_or(false, kinsoku::is_yakumono_opening)
                        {
                            v[i] = true;
                        } else if prev_of(i).map_or(false, kinsoku::is_yakumono_trigger)
                            && !(i > 0 && v[i - 1])
                        {
                            v[i] = true;
                        }
                    }
                }
                v
            } else {
                vec![false; chars_vec.len()]
            };

            // S1251: "nothing follows this tab" needs two scans; both are
            // per-FRAGMENT, so resolve them once here rather than per tab
            // character (a tab-heavy paragraph would otherwise be quadratic).
            let s1251_rest_blank = fragments[frag_outer_idx + 1..]
                .iter()
                .all(|f| f.0.chars().all(|c| matches!(c, ' ' | '\t')));
            let s1251_last_content =
                chars_vec.iter().rposition(|c| !matches!(c, ' ' | '\t'));
            // S1488 v2: set when a hung mark just closed the line; the ASCII
            // spaces that follow belong to that line, not to a new one.
            let mut s1488_after_hang = false;
            for (char_index, ch) in chars_vec.iter().copied().enumerate() {
                // Voicing marks are part of the preceding glyph cluster. They
                // neither consume a grid cell nor create a new break/gap. Keep
                // source text and character offsets, including across runs.
                if crate::font::is_nonspacing_kana_mark(ch)
                    && std::env::var_os("OXI_KANA_COMBINING_DISABLE").is_none()
                {
                    flush_word!(style);
                    if let Some(last) = current_line.fragments.last_mut().filter(|last|
                        last.run_index == frag_run_index && last.field_type == frag_field_type)
                    {
                        last.text.push(ch);
                    } else {
                        current_line.fragments.push(LineFragment {
                            auto_space_shrink: 0.0,
                            text: char_to_string(ch), width: 0.0, natural_width: 0.0,
                            style: style.clone(), tab_alignment: None, tab_position: None,
                            field_type: frag_field_type, run_index: frag_run_index,
                            char_offset: char_pos_in_run,
                        });
                    }
                    char_pos_in_run += 1;
                    continue;
                }

                if s1488_after_hang && std::env::var_os("OXI_S1488_DISABLE").is_none() {
                    if ch == ' ' && current_line.fragments.is_empty() {
                        if let Some(prev) = lines.last_mut() {
                            if let Some(last) = prev.fragments.last().cloned() {
                                let mut f = last;
                                f.char_offset += f.text.chars().count();
                                f.text = " ".to_string();
                                f.width = 0.0;
                                f.natural_width = 0.0;
                                f.auto_space_shrink = 0.0;
                                prev.fragments.push(f);
                            }
                        }
                        continue;
                    }
                    s1488_after_hang = false;
                }
                // S1443 (2026-09-17, default ON, opt-out OXI_S1443_DISABLE): a manual
                // page break inside a table cell is inert in Word — no line, no page
                // (tests/fixtures/cellbr: 24/24 arms over compat 14/15 x tblHeader x
                // cantSplit x break position keep the cell text on one line; a body
                // paragraph's break still starts a page). technical__00c13e6a's
                // Table 2-3 title cell opens with three breaks: Oxi drew two empty
                // lines above the title and the page's last row fell over.
                if ch == '\x0C'
                    && IN_TABLE_LAYOUT.with(|c| c.get()) > 0
                    && std::env::var_os("OXI_S1443_DISABLE").is_none()
                {
                    continue;
                }
                // LATINQUOTE (2026-07-07, default ON, opt-out
                // OXI_LATINQUOTE_DISABLE): a curly
                // quote (U+2018/2019/201C/201D) ADJACENT to an ASCII
                // alphanumeric is LATIN punctuation, not a CJK fullwidth char.
                // is_cjk()'s blanket General-Punctuation range classified it
                // CJK, which (a) split the Latin token at the quote (db9ca:
                // Oxi broke the token after the closing quote where Word keeps
                // the whole space-delimited token together and wraps it) and
                // (b) priced it at the 0.5em unknown-char fallback (5.25pt vs
                // TNR's true 909/2048em = 4.66; tables now carry the measured
                // advances). Word segments script runs by adjacency; a quote
                // glued to Latin letters joins the Latin run.
                // S801: the en/em dash (U+2013/U+2014) joins the ambiguous class —
                // in a LATIN document a dash is Latin punctuation regardless of
                // adjacency (dashes sit between SPACES, unlike glued quotes;
                // ukframework «auditor – shall»: the eastAsia-class dash split
                // the token and took the eastAsia width/line-height). Doc-level
                // gate = !doc_body_has_real_cjk → JP byte-identical.
                let s801_latin_dash = matches!(ch, '\u{2013}' | '\u{2014}')
                    && !self.doc_body_has_real_cjk
                    && std::env::var("OXI_S801_DISABLE").is_err();
                // S888: U+2011 NON-BREAKING HYPHEN (S747's noBreakHyphen) +
                // U+2010 HYPHEN join the ambiguous class — is_cjk routed them
                // to the eastAsia metrics, where S634's Latin-only
                // substitution (MS Mincho, upm-256) priced them at fs/2 =
                // 6.0 @TNR12 vs Word's hyphen 4.0 (legal__0001482d's
                // noBreakHyphen ISBN/Gazette lines wrapped one line early:
                // +14..17pt ×4 gap anomalies). Doc-level Latin gate like
                // S801 → JP byte-identical by construction.
                let s888_latin_hyphen = matches!(ch, '\u{2011}' | '\u{2010}')
                    && !self.doc_body_has_real_cjk
                    && std::env::var("OXI_S888_DISABLE").is_err();
                // S951 (2026-07-20): Arrows + Mathematical Operators
                // (U+2190..U+22FF, East-Asian-Width Ambiguous) join the class —
                // reference__0029c1c's «AHI ≥ 30» lines: is_cjk routed ≥ to the
                // eastAsia chain (Book Antiqua eastAsia → S634 MS Mincho) and
                // the LINE grew 22.37→25.50 where Word keeps the Latin height
                // (measured uniform 22.32-22.44 on every ≥ line). Doc-level
                // Latin gate like S801/S888 → JP byte-identical by construction
                // (golden census: 0 Latin docs carry these chars).
                let s951_latin_mathop = matches!(ch, '\u{2190}'..='\u{22FF}')
                    && !self.doc_body_has_real_cjk
                    && std::env::var("OXI_S951_DISABLE").is_err();
                let s966_latin_bullet = ch == '\u{2022}'
                    && !self.doc_body_has_real_cjk
                    && std::env::var("OXI_S966_DISABLE").is_err();
                // S1103 (2026-08-08): U+2015 HORIZONTAL BAR joins the ambiguous
                // class — see resolve_font_family_for_text_g.  Deliberately NOT
                // added to `s1100_dash_break`: this only fixes the LINE HEIGHT,
                // and legal__000ad039's catchword block already wraps into the
                // same 12 lines as Word (its drift is a flat +3.40/line), so
                // introducing a new break opportunity could only move them.
                let s1103_latin_hbar = ch == '\u{2015}'
                    && !self.doc_body_has_real_cjk
                    && std::env::var("OXI_S1103_DISABLE").is_err();
                // S1178 (2026-08-20): U+00D7 × / U+00F7 ÷ join the ambiguous
                // class — Latin-1's two EAW-Ambiguous math signs. is_cjk claims
                // them unconditionally, splitting the fragment and pricing the
                // sign down the eastAsia chain; Word draws creative__0158c02a's
                // ÷ lines whole in ArialMT at the plain pitch (44 slips) and
                // the _pb_symline probe measures both = the Latin control in
                // every font/variant arm. Doc-level Latin gate like the
                // siblings → JP byte-identical by construction.
                let s1178_latin_muldiv = matches!(ch, '\u{00D7}' | '\u{00F7}')
                    && !self.doc_body_has_real_cjk
                    && std::env::var("OXI_S1178_DISABLE").is_err();
                // S1442 (2026-09-16, default ON, opt-out OXI_S1442_DISABLE): the
                // ellipsis U+2026 / U+2025 in a Latin body is Latin text — a run of
                // them wraps as an unbreakable word that fills the line (Word: 36
                // per 432pt line, legal__0022399405 p6), not as a kinsoku mark that
                // may never open a line (Oxi: one per line until the tail fits).
                let s1442_latin_ellipsis = matches!(ch, '\u{2026}' | '\u{2025}')
                    && !self.doc_body_has_real_cjk
                    && std::env::var_os("OXI_S1442_DISABLE").is_none();
                let latin_ctx_quote = s801_latin_dash
                    || s888_latin_hyphen
                    || s951_latin_mathop
                    || s1442_latin_ellipsis
                    || s966_latin_bullet
                    || s1103_latin_hbar
                    || s1178_latin_muldiv
                    || matches!(ch, '\u{2018}' | '\u{2019}' | '\u{201C}' | '\u{201D}')
                        && std::env::var("OXI_LATINQUOTE_DISABLE").is_err()
                        && {
                            // neighbors ACROSS fragment boundaries (db9ca's quotes
                            // sit at run boundaries: «…as “» | «This…» | «”). …»);
                            // a glued non-space ASCII neighbor makes it Latin
                            // («Use”)» — the ')' is ASCII punctuation, still Latin).
                            let prev = char_index
                                .checked_sub(1)
                                .and_then(|i| chars_vec.get(i))
                                .copied()
                                .or_else(|| {
                                    fragments[..frag_outer_idx]
                                        .iter()
                                        .rev()
                                        .find_map(|f| f.0.chars().last())
                                });
                            let next = chars_vec.get(char_index + 1).copied().or_else(|| {
                                fragments[frag_outer_idx + 1..]
                                    .iter()
                                    .find_map(|f| f.0.chars().next())
                            });
                            let latin_side = |c: Option<char>| {
                                c.map_or(false, |c| c.is_ascii() && !c.is_ascii_whitespace())
                            };
                            latin_side(prev) || latin_side(next)
                        };
                // Vertical Word keeps curly quotes in the East Asian face even
                // beside rotated Latin letters (mixed-font primary controls).
                let latin_ctx_quote = latin_ctx_quote && !(vertical
                    && crate::font::vertical_font_advance_on()
                    && matches!(ch, '\u{2018}' | '\u{2019}' | '\u{201C}' | '\u{201D}'));
                let use_east_asia = self.ambiguous_symbol_east_asia(ch, style, para_style)
                    .unwrap_or_else(|| kinsoku::is_cjk(ch) && !latin_ctx_quote);
                let (char_metrics, gdi_map) = if use_east_asia {
                    let (selected_metrics, selected_gdi_map) = if LayoutEngine::s1370_is_cjk_script(ch) {
                        (substitute_metrics, substitute_gdi_map)
                    } else {
                        (cjk_metrics, cjk_gdi_map)
                    };
                    if let Some(cjk_m) = selected_metrics {
                        (cjk_m, selected_gdi_map)
                    } else {
                        (latin_metrics, latin_gdi_map)
                    }
                } else {
                    (latin_metrics, latin_gdi_map)
                };
                let mut char_width =
                    self.registry
                        .char_width_pt_with_gdi_map(ch, font_size, &char_metrics, gdi_map);
                // KERNBREAK (2026-07-07, ★default ON, opt-out
                // OXI_KERNBREAK_DISABLE): a
                // KERN-ACTIVE Latin char (w:kern set, fs >= threshold —
                // Word applies font kerning to both RENDER and BREAK) breaks
                // at the UNROUNDED em advance PLUS the font's kern-pair
                // adjustment. db9ca derivation (TNR 10.5, docDefaults
                // kern=2): Word PDF non-space advance sum over a line =
                // em+kern within -0.20pt/71 advances; the com_tw model runs
                // ~4.5pt/line NARROW (over-packs ~1 word), plain em runs
                // +0.5 letters + kern-less WIDE (under-packs 1-2 words,
                // the falsified LATIN_TRUE_BREAK experiment). Kern pairs =
                // fontTools-extracted legacy kern tables (ASCII + curly
                // quotes). Scope: kern-active runs only — no-kern docs
                // (gen/gen2/test authored without w:kern) keep the
                // S672-validated com_tw break.
                let kern_active = std::env::var("OXI_KERNBREAK_DISABLE").is_err()
                    && style
                        .kern
                        .or_else(|| para_style.default_run_style.as_ref().and_then(|rs| rs.kern))
                        .map_or(false, |k| k > 0.0 && font_size >= k);
                if std::env::var("OXI_DBG_KERN").is_ok() && char_index == 0 {
                    eprintln!("[DBG_KERN] fs={} ch={:?} style.kern={:?} drs.kern={:?} active={} widths_has={} fam={}",
                        font_size, ch, style.kern,
                        para_style.default_run_style.as_ref().and_then(|rs| rs.kern),
                        kern_active, char_metrics.char_widths.contains_key(&ch), char_metrics.family);
                }
                // NBSP aliases the space glyph when the metrics omit it.
                // Apply that alias to the unrounded Latin advance as well:
                // the legacy rounded path otherwise gives TNR 11pt 2.5pt
                // for NBSP while ordinary space advances 2.75pt. Retain the
                // original character for kerning and break opportunities.
                let space_metric_char = if ch == '\u{00a0}'
                    && crate::font::s892_nbsp_as_space()
                    && std::env::var_os("OXI_NBSP_EM_DISABLE").is_none()
                    && !char_metrics.char_widths.contains_key(&ch)
                { ' ' } else { ch };
                if kern_active
                    && (!kinsoku::is_cjk(ch) || latin_ctx_quote)
                    && char_metrics.char_widths.contains_key(&space_metric_char)
                {
                    char_width = char_metrics.char_width_em(space_metric_char) * font_size;
                    if let Some(&nxt) = chars_vec.get(char_index + 1) {
                        if !kinsoku::is_cjk(nxt) {
                            char_width += self.registry.latin_kern_em(
                                &char_metrics.family,
                                char_metrics.units_per_em,
                                ch,
                                nxt,
                            ) * font_size;
                        }
                    }
                }
                // LATINEM (2026-07-09, ★default ON, opt-out OXI_LATINEM_DISABLE): a
                // NO-KERN pure-Latin char breaks at the UN-ROUNDED em advance. Word
                // breaks no-kern Latin (Times New Roman etc.) at the true em; Oxi's
                // com_tw (10tw-per-char round) sums ~4.5pt/line NARROW → over-packs ~1
                // word/line. This is the break-boundary fix S672 deferred (S672 fixed
                // the RENDER x to true-em but kept the com_tw BREAK). The old blanket
                // "LATIN_TRUE_BREAK" was falsified because it also hit KERN-ACTIVE runs
                // (which need em+kern, now KERNBREAK); scoped to !kern_active, KERNBREAK
                // handles kern-active and LATINEM handles no-kern — no exceptions.
                // Scope: non-CJK docs (!doc_body_has_real_cjk → JP byte-identical,
                // Phase-1 safe) + metric fonts. GATE: tracked corpus ssim_ab +0.0624
                // (test_line_heights +0.0428, test_lists, gen_headings, gen_long),
                // 5 improved / 0 regressed / 0 page-count shifts; real_en mean +0.0002;
                // fixes nyserda pagination (57→56=Word, Exhibit boundaries aligned).
                let no_break_hyphen_em = ch == '\u{2011}' && s888_latin_hyphen
                    && std::env::var("OXI_NOBREAK_HYPHEN_EM_DISABLE").is_err();
                let latin_metric_char = if no_break_hyphen_em
                    && !char_metrics.char_widths.contains_key(&ch) { '-' } else { space_metric_char };
                if !kern_active
                    && std::env::var("OXI_LATINEM_DISABLE").is_err()
                    && latinem_in_scope(self.doc_body_has_real_cjk)
                    && (!kinsoku::is_cjk(ch) || no_break_hyphen_em)
                    && char_metrics.char_widths.contains_key(&latin_metric_char)
                {
                    char_width = char_metrics.char_width_em(latin_metric_char) * font_size;
                }
                // Complex-script cluster width: replace the per-codepoint value
                // with the shaped cluster advance (0 on non-first cluster chars),
                // BEFORE text_scale so w:w still applies. `deva_cont` marks a char
                // that continues a cluster, so letter-spacing is added once per
                // cluster (as Word does), not on every mark.
                let deva_cont = if let Some((ref adv, ref cstart)) = deva_adv {
                    char_width = adv[char_index];
                    !cstart[char_index]
                } else {
                    false
                };
                let vertical_font_advance = if vertical
                    && crate::font::vertical_font_advance_on()
                {
                    self.registry.vertical_advance_pt(&char_metrics.family, ch, font_size)
                } else {
                    None
                };
                // Only replace the boundary model where actual vertical metrics
                // are available. Missing tables retain their structural spacing.
                let s1446_proportional = vertical_font_advance
                    .is_some_and(|a| a < font_size * 0.98);
                let vertical_natural_boundary = vertical_natural_boundary_enabled
                    && ((s1446_natural_boundary_explicit && vertical_font_advance.is_some())
                        || s1446_proportional
                        || (crate::font::vertical_char_grid_on()
                            && grid_char_pitch.is_some() && grid_char_cw_ratio.is_some()));
                if let Some(advance) = vertical_font_advance {
                    char_width = advance;
                } else if vertical && crate::font::vertical_font_advance_on()
                    && !kinsoku::is_cjk(ch) && deva_adv.is_none() && !kern_active
                {
                    // Rotated Latin uses its horizontal glyph advance, not an em
                    // cell or the half-point-rounded legacy break width.
                    char_width = char_metrics.char_width_em(ch) * font_size;
                }
                if let Some(scale) = style.text_scale {
                    if (scale - 100.0).abs() > 0.01 {
                        char_width *= scale / 100.0;
                    }
                }
                // S812 (2026-07-13) ATTEMPTED + FALSIFIED + REVERTED: "a justified
                // paragraph's SPACE-run w:spacing is excluded from the break" —
                // ukframework wp15 render-truth seemed to show it (Word fits 'Board'
                // where the cs-inclusive width overflows by ~5.3pt, rendered spaces
                // shrunk toward natural). But the CONTROLLED sweep (_pb_cs_gen.py:
                // per-space w:spacing V in {0,20,40} x jc x margin sweep, + a
                // substituted-font variant) proves Word COUNTS space-cs FULLY in the
                // break (flip boundary = cs-less width + n*V EXACT, both installed
                // Arial and substituted Humnst777 BT; ZERO shrink granted in the
                // synthetic), and dropping it over-fit 16 framework lines (-1x16
                // cascade from wp18). The framework wp15 line is a JUSTIFY-SHRINK
                // case (~0.66pt/space granted) whose enabling condition the synthetic
                // lacks — the underived S799-shrink model (dedicated session; vary
                // para line count / following content / numPr / cs magnitude).
                char_width += if deva_cont { 0.0 } else { cs };
                // §17.15.1.7 balanceSingleByteDoubleByteWidth (Session 56 Finding 3,
                // COM-confirmed via V19/V25/V26/V27 minimal repros 2026-05-06):
                // when this compat flag is set, character_spacing is applied TWICE
                // for CJK fullwidth chars (effective_cs = 2 * cs). Apply the extra
                // cs here so per-char fragment advance reflects the doubled spacing.
                // Day 37 (2026-05-14): EXCLUDE fitText runs — resolve_fit_text_runs
                // already produces the FINAL effective cs (post-balance-doubling) so
                // adding here would over-pump by another factor.
                if self.balance_single_byte_double_byte_width
                    && if vertical && crate::font::vertical_font_advance_on() {
                        LayoutEngine::vertical_balance_spacing_char(ch)
                    } else {
                        crate::font::is_fullwidth(ch) && !yakumono_compressed[char_index]
                    }
                    && style.fit_text.is_none() && !style.ruby_spread
                {
                    char_width += cs;
                }
                // 2-pass wrap: remember pre-yakumono width to compute yakumono savings.
                let mut pre_yakumono_width = char_width;
                // Physical yakumono compression (COM-confirmed b837 2026-04-16):
                //   Pair (both chars): 6pt (×0.5) — e.g., 。）→ 6+6pt
                //   Standalone 、。 between non-trigger CJK: 7pt (×0.583)
                //   Other brackets: use native font width (bracket shapes vary widely by
                //     context in Word — 6, 10.5, 11, 11.5, 12pt — no simple compression rule)
                // 2026-04-20: Opening brackets have visible glyph at right side of
                // cell (ABC A-offset = 7.5pt for 「, 11pt for （). Compressing advance
                // to 6pt would place next char at 6pt offset, overwriting the bracket
                // glyph at 7.5-11.25pt. Skip compression for these — keep fullwidth
                // advance so glyph fits within its cell. Closing brackets (A=0) are
                // unaffected and still compress fine.
                let is_opening_bracket = matches!(
                    ch,
                    '（' | '「' | '『' | '〔' | '【' | '《' | '〈' | '｛' | '［'
                );
                // S1217: the 2026-04-20 carve-out above is about an opening bracket
                // whose INK sits in the right half of its box being overrun by the
                // next character. That cannot happen when the next character is
                // itself an opening bracket: Word's own origins for 「（（」 (106.22 /
                // 111.39) put the first bracket's ink in the second bracket's left
                // half, and the second's ink follows it. So the carve-out is lifted
                // for exactly that pair.
                let s1217_next_open = std::env::var("OXI_S1217_DISABLE").is_err()
                    && chars_vec
                        .get(char_index + 1)
                        .copied()
                        // S1494: the pair may continue in the next run
                        .or_else(|| if std::env::var_os("OXI_S1494_DISABLE").is_none() {
                            fragments[frag_outer_idx + 1..].iter().flat_map(|f| f.0.chars()).next()
                        } else { None })
                        .map_or(false, kinsoku::is_yakumono_opening);
                // vmtx after vertical substitution already contains the vertical
                // punctuation advance. Horizontal pair compression would halve it
                // a second time (Word vertical controls retain the raw advance).
                if vertical_font_advance.is_none()
                    && yakumono_compressed[char_index] && (!is_opening_bracket || s1217_next_open)
                {
                    char_width *= 0.5;
                } else if vertical_font_advance.is_none() && yakumono_enabled {
                    // S532 (2026-06-10): the former "expand pair" rule (a yakumono
                    // ADJACENT to a pair-compressed one also compresses ×0.5) is
                    // REMOVED — Word compresses ONLY the FIRST char of an adjacent
                    // pair; the second keeps its natural advance. Measured
                    // (_s532_pair_repro.py, MS Gothic 12pt, PDF per-char origins):
                    // 。」=6.00/12.00, ）」=6.00/12.00, 。「=6.00/12.00 — identical
                    // in centered, loose-justified and wrapping-justified lines.
                    // (The 2026-04-16 b837 "。）→6+6" COM note conflated the pair
                    // rule with justify-demand compression of the second char.)
                    let is_yakumono_any = matches!(
                        ch,
                        '（' | '）'
                            | '「'
                            | '」'
                            | '『'
                            | '』'
                            | '〔'
                            | '〕'
                            | '【'
                            | '】'
                            | '《'
                            | '》'
                            | '〈'
                            | '〉'
                            | '｛'
                            | '｝'
                            | '［'
                            | '］'
                            | '、'
                            | '。'
                            | '，'
                            | '．'
                    );
                    if is_yakumono_any {
                        if matches!(ch, '、' | '。' | '，' | '．') {
                            // Standalone 、 。 between non-triggers: spec §4.7b round 5
                            // floor = fontSize × 2/3. Trying 0.667 instead of 0.583.
                            let prev_non_tr = char_index == 0
                                || !kinsoku::is_yakumono_trigger(chars_vec[char_index - 1]);
                            let next_non_tr = char_index + 1 >= chars_vec.len()
                                || !kinsoku::is_yakumono_trigger(chars_vec[char_index + 1]);
                            if prev_non_tr && next_non_tr {
                                // S472: ALL standalone 、，。．use NATURAL width at break
                                // (Word defers compression to justify-demand; COM: 、/。
                                // standalone = near-full, only compressed on line-slack).
                                // The demand-absorb below compresses any of them as a
                                // line's overflow requires.
                                // S1318b (2026-09-05): the at-default legacy regime
                                // (S1237/S1318) breaks with standalone marks at their
                                // NATURAL width too -- the probe's 3- and 4-mark arms
                                // still packed a 41st character because this flat
                                // x0.6667 pre-compress banked 70tw per mark at break
                                // time (Word: every mark 1.00em, line holds 40).
                                if (s472_demand
                                    || s474_natural
                                    || s475_break
                                    || s557_natural_just15
                                    || s589_legacy_just_natural
                                    || para_off_grid
                                    || japanese_language_oikomi
                                    || s1237_at_default_refuse)
                                    && matches!(ch, '、' | '，' | '。' | '．')
                                {
                                    // no compression at break; demand-absorb handles fit
                                    // (s474_natural: leave natural, no absorb either =
                                    // pure natural-greedy diagnostic; s557: c15 justified
                                    // keeps naturals for the pack tier)
                                    // S592: para_off_grid (proportional CJK font in a
                                    // linesAndChars grid) keeps its 約物 at the font's
                                    // natural proportional advance — HGPGothicM 、 = 6.75pt
                                    // is already narrow (no fullwidth aki), Word does NOT
                                    // pre-compress it ×0.6667 (→4.5pt over-packs the line).
                                } else {
                                    char_width *= 0.6667;
                                }
                            }
                        }
                    }
                }
                // Line-start yakumono demand-driven compression (COM-verified 2026-04-21
                // on d77a pi=24-27 + 3a4f pi=300 + 1ec1/e3c5 no-overflow):
                // Word compresses ・/、/。 at line start by ~2.5pt at 12pt when the
                // line would otherwise overflow. Apply the compression speculatively;
                // Stage 2 revert (below) undoes it on short lines where
                // natural_total_width ≤ available_width (loose-line rule).
                // Compression: font_size × 5/24 = 2.5pt at 12pt, 2.1875pt at 10.5pt.
                // Gated on compress_punctuation + compat_mode>=15 to match Word 2016+.
                if vertical_font_advance.is_none() && yakumono_enabled
                    && self.compat_mode >= 15
                    && !s474_natural
                    && matches!(ch, '・' | '、' | '。' | '，' | '．')
                    && current_line.fragments.is_empty()
                    && word.is_empty()
                {
                    let reduction = font_size * 5.0 / 24.0;
                    let floor = char_width * 0.5;
                    char_width = (char_width - reduction).max(floor);
                }
                // 2-pass wrap: compute yakumono savings (difference between pre-yakumono
                // and post-yakumono width). Natural = final_char_width + yakumono_saved.
                let yakumono_saved = (pre_yakumono_width - char_width).max(0.0);
                // Only the document's single/double-byte balancing setting
                // widens ASCII spaces adjacent to CJK. Font declarations and
                // run boundaries do not change this setting. Neighbours may
                // belong to the preceding or following text fragment.
                let s1333_balance = self.balance_single_byte_double_byte_width
                    && std::env::var("OXI_S1333_DISABLE").is_err();
                // NBSP keeps its non-breaking semantics and Latin font slot,
                // but shares the balanced space advance next to CJK. Saved
                // Word controls measure 6pt at 12pt for both space characters;
                // disabling balance restores their natural space advance.
                if matches!(ch, ' ' | '\u{00a0}') && s1333_balance {
                    let prev_is_cjk = chars_vec
                        .get(char_index.wrapping_sub(1))
                        .copied()
                        .or_else(|| {
                            if char_index == 0 && s1333_balance {
                                fragments
                                    .get(frag_outer_idx.wrapping_sub(1))
                                    .and_then(|f| f.0.chars().last())
                            } else {
                                None
                            }
                        })
                        .map_or(false, kinsoku::is_cjk_ideograph_or_kana);
                    let next_is_cjk = chars_vec
                        .get(char_index + 1)
                        .copied()
                        .or_else(|| {
                            if char_index + 1 == chars_vec.len() && s1333_balance {
                                fragments.get(frag_outer_idx + 1).and_then(|f| f.0.chars().next())
                            } else {
                                None
                            }
                        })
                        .map_or(false, kinsoku::is_cjk_ideograph_or_kana);
                    if prev_is_cjk || next_is_cjk {
                        // S1337S (2026-09-06, HELD opt-in OXI_S1337S=1): on a character
                        // grid the balanced space is half the CELL, like every other
                        // single-byte character -- 0ea3ec86 p5 「能訓練 31か所 生活訓練
                        // 74か所 402・403」: both ASCII spaces 5.76 at cell 11.52 (Word),
                        // and the 24th character ㌻ then overflows by 0.54 cell and
                        // wraps; 167853 (97) and 0b6f3b32 (17) show only half-cell
                        // spaces, none at half an em. Held because the 12-base SSIM A/B
                        // reads net -0.0002 (tokumei_08 -0.0006 / -0.0005, whose Word
                        // PDF cannot tell 5.43 from 5.25 at its 0.12pt quantisation):
                        // find the compensating error there before defaulting it.
                        char_width = match (grid_char_cw_ratio, grid_char_pitch) {
                            (Some(ratio), Some(pitch))
                                if ratio > 0.0
                                    && pitch > 0.0
                                    && (std::env::var("OXI_S1337S").ok().as_deref() == Some("1")
                                        || std::env::var("OXI_GRID_SPACE_DISABLE").is_err()) =>
                            {
                                let default_fs = pitch / ratio;
                                let char_space_pt = pitch - default_fs;
                                let base = if char_space_pt >= 0.0 {
                                    font_size * pitch / default_fs
                                } else {
                                    font_size + char_space_pt
                                };
                                0.5 * base + if std::env::var("OXI_GRID_SPACE_DISABLE").is_err() {
                                    2.0 * style.character_spacing.unwrap_or(0.0)
                                } else { 0.0 }
                            }
                            _ => font_size / 2.0 + char_metrics.synthetic_bold_advance * font_size,
                        };
                    }
                }
                let _ = char_index;
                // charGrid: ONLY full-width chars are padded to 1 grid cell.
                // §11.2.1 (Round 14, COM-confirmed): half-width Latin chars
                // (ASCII 0-9 / A-Z / etc.), CJK punctuation under yakumono
                // compression, and other halfwidth glyphs use their NATURAL
                // advance width — they are NOT snapped to the grid pitch.
                // Reference: b837808d0555 P13 L1 measurement showed
                //   '2'=6pt, '」'=6pt (yakumono), 成=15pt (12+autoSpaceDE),
                //   '7'=9pt (6+autoSpaceDE), ' '=6pt (TNR space natural).
                // Previous (buggy) behavior padded ALL chars, halving the
                // chars/line and causing 177-doc max-error of 0.5366 SSIM.
                // 2026-04-19 (revised): cw = fs + charSpace_pt (absolute, not scaled).
                // COM-measured b35 fs=9→8.3pt, fs=10.5→9.8pt: both = fs − 0.7pt.
                // Previous fs*ratio formula over-compressed at small fs in docs
                // where default_fs ≠ fs.
                // fit_text EXPAND mode (natural ≤ target, character_spacing>0): skip
                // charGrid padding so Word's fitText cs applies verbatim. Without this,
                // the negative char_grid_extra swallows the cs for CJK chars and breaks
                // uniform spread (b837 p1 meta block).
                // fit_text SCALE mode (natural > target, text_scale<100) keeps charGrid
                // padding — otherwise scaled CJK chars in table cells become narrower
                // than the grid pitch, shifting downstream content (3a4f regression).
                let fit_text_expand = style.fit_text.is_some()
                    && style.character_spacing.map_or(false, |cs| cs > 0.01);
                // S1347: the run's w:w scale factor for the grid cell (see below).
                let s1347_scale = if std::env::var("OXI_S1347_DISABLE").is_err()
                    && style.fit_text.is_none()
                {
                    style.text_scale.map_or(1.0, |sc| if (sc - 100.0).abs() > 0.01 && sc > 0.0 { sc / 100.0 } else { 1.0 })
                } else {
                    1.0
                };
                let char_grid_extra = if vertical && crate::font::vertical_char_grid_on()
                    && crate::font::is_fullwidth(ch) && !fit_text_expand
                    && grid_char_pitch.is_some() && grid_char_cw_ratio.is_some()
                {
                    let natural = LayoutEngine::vertical_grid_fullwidth_advance(font_size,
                        grid_char_pitch.unwrap(), grid_char_cw_ratio.unwrap(), quantized_char_grid);
                    let tracking = if self.balance_single_byte_double_byte_width
                        && LayoutEngine::vertical_balance_spacing_char(ch) && style.fit_text.is_none()
                        && !style.ruby_spread { 2.0 * cs } else { cs };
                    natural * s1347_scale + tracking - char_width
                } else if fit_text_expand {
                    0.0
                } else if let (Some(ratio), Some(pitch)) = (grid_char_cw_ratio, grid_char_pitch) {
                    if ratio > 0.0
                        && pitch > 0.0
                        && char_width > 0.0
                        && ch != ' '
                        && ch != '\t'
                        && ch != '\n'
                        && (crate::font::is_fullwidth(ch)
                            || (use_east_asia
                                && char_width >= 0.98 * font_size
                                && std::env::var_os("OXI_S1449_DISABLE").is_none()))
                        && !yakumono_compressed[char_index]
                    {
                        let default_fs = pitch / ratio;
                        let char_space_pt = pitch - default_fs;
                        // R7.59 (Day 36 part 3, 2026-05-13): hybrid grid-extra formula.
                        // charSpace>=0 (expansion): proportional. COM-verified d4d126
                        //   w_i=245 fs=10 default=10.5 cs=+0.575: Word renders ~10.547pt
                        //   advance (proportional, NOT the 10.5pt COM Information(WD_HPOS)
                        //   reports — that's the snapped logical width). Old linear
                        //   cw = 10+0.575 = 10.575pt over-expanded → 1-line→2-line wrap
                        //   regression on 35-char paragraphs.
                        // charSpace<0 (compression): linear. COM-verified b35 fs=9
                        //   cs=-0.66: Word=8.3pt (= 9-0.7).
                        // 10tw-snap variant tested 2026-05-13: slightly worse SSIM
                        //   (+2.1481 net vs +2.1859 net) because Word's INTERNAL
                        //   rendering uses raw proportional advance, not snapped.
                        // S141 H6 (2026-05-20): OXI_H6_GRID_GATE=1 gates expansion to
                        //   only fire when font_size >= default_fs. Word doesn't expand
                        //   small-font (sz < default) cell text to grid pitch even
                        //   though Oxi did via this formula. COM-confirmed: 法人等 sz=10
                        //   cell in a1d6/d4d126/de6e/6514f all render 1 line in Word
                        //   (33 chars × 10pt = 330pt natural fits 345pt cell) but Oxi
                        //   wraps to 2 (33 × 10.555 expansion = 348.3pt overflows).
                        let h6_gate_enabled = std::env::var("OXI_H6_GRID_GATE").is_ok();
                        let h7_gate_enabled = std::env::var("OXI_H7_GRID_GATE_LE").is_ok();
                        // S145 H8 (2026-05-21): per LibreOffice ww8par.cxx ImportDop,
                        // MS_WORD_COMP_GRID_METRICS is SET unconditionally for all Word
                        // imported docs. LibreOffice itrform2.cxx then SKIPS grid kern
                        // portions when MS_WORD_COMP_GRID_METRICS && !vertical. So MS
                        // Word actually NEVER applies grid char-pitch for horizontal text.
                        // OXI_H8_NO_GRID_KERN=1 skips entirely (no font_size check).
                        // S148 (2026-05-21) H8 refinement: only skip POSITIVE expansion
                        // (kern portions). Word DOES apply negative compression (e.g.
                        // b35 charSpace=-2714 → chars narrower than natural).
                        // S239 (2026-05-23): removed OXI_LEGACY_GRID_KERN
                        // legacy env-var fallback during hardening pass.
                        // S466 (2026-05-31, env-gated test): the H8 "Word never
                        // applies grid char-pitch for horizontal text" came from
                        // reading LibreOffice source (ww8par/itrform2). Direct Word
                        // COM on a charSpace=1453 (tokumei grid) BODY repro
                        // CONTRADICTS it: MS Mincho 10.5pt(=default) fits 44 chars/
                        // line in Word but Oxi (H8 skip => natural advance) fits 46.
                        // i.e. Word DOES expand fullwidth chars to the grid pitch
                        // when fs >= default_fs (S141 already COM-confirmed Word does
                        // NOT expand when fs < default, the cell case). So the correct
                        // skip is fs < default, not unconditional. Gated so the corpus
                        // (charGrid family tokumei/b35/b837) can be A/B re-gated;
                        // Phase-1-sensitive (re-wrap moves pagination).
                        // MEASURED (2026-08-24, `tools/metrics/_s1210_pitch.py`,
                        // a1d6e4ef + 6514f214, unjustified lines only): Word's own
                        // PDF export lays the page out at 600dpi -- every span size
                        // is a whole 0.12pt device pixel (9.5pt -> 9.48, 10pt ->
                        // 9.96, 10.5pt -> 10.56, 16pt -> 15.96), and so is every
                        // glyph origin. Read that way the advances are ONE law,
                        // ADDITIVE and independent of the run's size:
                        //     pitch = fs + charSpace/4096
                        // fs 9    -> 9.3547 = 77.96px: 95% of advances land on the
                        //                     78th pixel (9.36), 5% on the 77th
                        // fs 10   -> 10.3547 = 86.29px: 68% / 32%
                        // fs 10.5 -> 10.8547 = 90.46px: 55% / 45%
                        // fs 12   -> 12.3547 = 102.96px: 93% on 103px
                        // The proportional form (fs * pitch / default_fs, below)
                        // predicts 9.304 for fs 9 = 77.5px, i.e. an EVEN split
                        // between the 77th and 78th pixel; the measured 95/5
                        // falsifies it. The two forms differ by 0.027pt at fs 10 --
                        // which is why the 2026-05-13 single-size COM check could
                        // not separate them, and picked the wrong one.
                        // S1210 (2026-08-24, default ON 2026-08-26 (opt-out `OXI_S1210_DISABLE`), shipped with the derived-cell bundle): with the additive pitch the S141 carve-out
                        // ("Word does not expand a font SMALLER than the grid
                        // default") is unnecessary. The sz=10 cell it was derived
                        // from holds 33 x 10.3547 = 341.7pt in its 345pt cell -- one
                        // line, exactly as Word renders it -- so that observation was
                        // never evidence against expansion, only against the pitch
                        // Oxi used then. The carve-out costs a1d6e4ef its note
                        // column: 8 chars at the natural 9.00 fit where Word breaks
                        // after 7, and the 9th line that follows pushes the ※2 note
                        // onto the next page in Word but not in Oxi.
                        let s1210 = std::env::var("OXI_S1210_DISABLE").is_err();
                        let h8_trigger = char_space_pt > 0.0
                            && (!s466_grid_expand || (font_size < default_fs && !s1210));
                        let h7_trigger =
                            h7_gate_enabled && char_space_pt > 0.0 && font_size <= default_fs;
                        let h6_trigger =
                            h6_gate_enabled && char_space_pt > 0.0 && font_size < default_fs;
                        // S466: when the docGrid has NO charSpace (char_space_pt≈0), the
                        // grid is line-pitch-only and Word does NOT horizontally expand
                        // chars. Under raw_pitch (S466) such a doc yields char_space_pt=0
                        // (b837: charSpace absent, default 12pt), which would otherwise
                        // fall through to expected_w=fs and widen natural<fs chars,
                        // over-wrapping (7->9). Skip expansion for the no-charSpace case.
                        // S1315 (2026-09-05, default ON, opt-out OXI_S1315_DISABLE): a
                        // NEGATIVE charSpace shrinks the pitch below the font size.
                        // DERIVED (`_pb_charspaceneg_gen.py`, 21 linesAndChars arms
                        // x 10.5/11/12pt, first-line character counts): the pitch
                        // is fs + charSpace/4096 for every sign and size (11pt:
                        // -1440 -> 42 chars/453pt, -2880 -> 44, -4320 -> 45; +1440
                        // -> 39, +2880 -> 38); a `lines` grid ignores charSpace.
                        // S466 lumped "negative" with "no charSpace" and kept the
                        // natural advance (41), the 7% that put reference__0cf9c879
                        // on two pages.
                        let s466_no_grid = s466_grid_expand
                            && if std::env::var("OXI_S1315_DISABLE").is_err() {
                                char_space_pt.abs() < 0.01
                            } else {
                                char_space_pt < 0.01
                            };
                        // S344 (2026-05-27): when S344 fed grid values through despite
                        // snap_to_grid=false, gate compression to fs < default_fs only.
                        // (Effective only when paired with S342/S344 pass-through at
                        // mod.rs:4073/4246.)
                        let s344_fs_gate = std::env::var("OXI_S344_FS_LT_DEFAULT")
                            .map(|v| v != "0" && v != "false")
                            .unwrap_or(false);
                        let s344_skip =
                            s344_fs_gate && !para_style.snap_to_grid && font_size >= default_fs;
                        if h6_trigger || h7_trigger || h8_trigger || s344_skip || s466_no_grid {
                            0.0
                        } else {
                            // S1315 keeps the ADDITIVE form for a negative charSpace: the
                            // b35123 truth PDF advances 12/11/10/9pt text at fs - 0.663
                            // (11.34 / 10.35 / 9.35 / 8.34), neither proportional
                            // (11.20) nor twips-floored (11.30).
                            let expected_w = if char_space_pt >= 0.0 && !s1210 {
                                font_size * pitch / default_fs
                            } else if std::env::var_os("OXI_S1510_DISABLE").is_none()
                                && std::env::var_os("OXI_S1592_DISABLE").is_none()
                                && face_has_proportional_kana(&char_metrics, &self.registry, font_size)
                            {
                                // S1592: on a proportional face the increment is by CLASS
                                // (see grid_half_increment_class): half for kana / U+30FC /
                                // U+3001-3002, the whole charSpace for everything else, each
                                // on top of the glyph's own advance.
                                char_width + if grid_half_increment_class(ch) {
                                    0.5 * char_space_pt
                                } else {
                                    char_space_pt
                                }
                            } else if std::env::var_os("OXI_S1510_DISABLE").is_none()
                                // S1539 (2026-09-25): judge "proportional" on the UNSCALED
                                // width. A w:w-scaled full-width glyph (66%: 7.26 of 11)
                                // passed this test, took the half-charSpace expected width
                                // (7.63) and then S1347's cell scaling on top of it ->
                                // 4.46 per glyph, where Word draws 7.91/8.03 (probe
                                // `_pb_wscale_gen.py`, g_kana_lc_w66 / g_mixed_lc_w66:
                                // Word PDF glyph origins on a linePitch 411 / charSpace
                                // 3042 grid). Unscaled, the glyph is full-width, the
                                // expected width is the pitch and S1347 gives 7.87.
                                && char_width / s1347_scale < 0.98 * font_size
                            {
                                // S1510 (2026-09-20, default ON, opt-out OXI_S1510_DISABLE):
                                // the additive pitch is per GLYPH -- a proportional
                                // glyph narrower than the em keeps its own advance plus
                                // the charSpace, it is not stretched to the full pitch.
                                // MEASURED (reports__393aa9, linesAndChars charSpace
                                // 4884 = +1.19pt, MS PMincho 10pt, Info5 steps): kanji
                                // 11.25 = pitch; kana オ 9.75 / ス 9.0 / ト 6.75 / リ
                                // 6.75 (natural 9.5 / 8.0 / 6.0 / 6.0 + 1.19); ・ 6.0
                                // (5.0 + 1.19); 、 6.75; digits 5.25..6.0.
                                // v3 (kana_grid_probe.py, 40 P-Mincho kana / 30 kanji
                                // on one line, charSpace 0 / 4884 / 9768 / 19536 /
                                // -2000): kanji +1.19 / +2.35 / +4.76 / -0.52 per glyph
                                // (= charSpace/4096), kana +0.58 / +1.17 / +2.37 /
                                // -0.27 -- exactly HALF. A proportional glyph takes
                                // half the charSpace, the way a single-byte glyph
                                // takes half the cell increment under balance.
                                char_width + 0.5 * char_space_pt
                            } else {
                                // S1210: additive for BOTH signs (see above).
                                font_size + char_space_pt
                            };
                            // S1347 (2026-09-07, default ON, opt-out OXI_S1347_DISABLE): a
                            // run's character scale (w:w) scales the grid CELL, not just
                            // the glyph -- Word draws 0ea3ec86's 90% runs at 10.37 (pitch
                            // 11.51) and 10.6 (11.76) so 「担当課　中部総合精神保健福祉セン
                            // ター事務室」 (21 glyphs) fits its 20-cell column. The scaled
                            // advance, from the faithful slice's 20-value sweep
                            // (`_pb_unitcap_gen.py` u_w*_nat, glyph origins, 10.5pt body):
                            //   a(s) = s * (pitch + fs) / 2 + (pitch - fs) / 2
                            // exact (<= 0.01) at s = 1/2, 3/4, 1, 5/4, 3/2, 2 (6.00, 8.74,
                            // 11.51, 14.24, 17.00, 22.50); the other values sit 0.03-0.085
                            // BELOW it (80%: 9.24, 90%: 10.36, 110%: 12.52) -- a
                            // quantisation still unexplained. pitch * s would miss 50% by
                            // 0.25 and 200% by 0.5 per character. The floor stays the
                            // section's 20 cells (the 23rd at 90% = 237.8 wraps; the 25th
                            // at 80% = 231.0 wraps), digits take half the scaled advance,
                            // a 、 at 90% gives 3.5 of elective compression. Open: a 12pt
                            // run on the 10.5 grid reads 12.50 = fs + (pitch - fs)/2, not
                            // Oxi's fs * pitch / default_fs = 13.15 (pre-existing, not
                            // touched here). A fit_text scale keeps the whole cell (3a4f).
                            let expected_w = if (s1347_scale - 1.0).abs() > 1e-6 {
                                s1347_scale * (expected_w + font_size) / 2.0 + (expected_w - font_size) / 2.0
                            } else {
                                expected_w
                            };
                            // S1340 (2026-09-06, default ON, opt-out OXI_S1340_DISABLE): a
                            // run's tracking survives the grid cell -- the padding to the
                            // cell used to erase it (char_width already carried cs, the
                            // balance doubling too, and expected_w did not). Word's PDF of
                            // reference__0ea3ec86: the `w:spacing=-4` run 「…研修を行う。」
                            // advances 11.04/11.15 (= cell 11.5 - 2 x 0.2) on its LAST
                            // line, so 21 characters + a hanging 。 fit where Oxi, at
                            // 11.5, wrapped 「う。」 and pushed a line onto every page
                            // after. (-12 reads 10.56 there, not the 10.32 of a strict
                            // doubling -- one witness, kept as the doubling.)
                            let s1340_track = if std::env::var("OXI_S1340_DISABLE").is_err()
                                && style.fit_text.is_none()
                                && !style.ruby_spread
                            {
                                cs * if self.balance_single_byte_double_byte_width { 2.0 } else { 1.0 }
                            } else {
                                0.0
                            };
                            expected_w + s1340_track - char_width
                        }
                    } else if ratio > 0.0
                        && pitch > 0.0
                        && char_width > 0.0
                        && ch.is_ascii_graphic()
                        && (char_width - 0.5 * font_size).abs() < 0.06 * font_size
                        && std::env::var("OXI_S1337_DISABLE").is_err()
                        && (self.balance_single_byte_double_byte_width
                            || std::env::var("OXI_S1337B").ok().as_deref() == Some("1"))
                    {
                        // S1337 (2026-09-06, default ON, opt-out OXI_S1337_DISABLE): on a
                        // character grid a HALF-WIDTH character (digit, ASCII letter or
                        // punctuation drawn at half an em) is not left at its natural
                        // advance. With balanceSingleByteDoubleByteWidth it advances
                        // HALF THE CELL: Word's PDF of reference__0ea3ec86 (ＭＳ 明朝 11pt)
                        // puts every digit, ASCII paren and even the letters of FAX at
                        // 5.76 in its charSpace-2048 sections (cell 11.52) and at 5.88 in
                        // the 3194 sections (cell 11.76); reports__167853 (126 digits at
                        // 5.76) and reference__0b6f3b32 (66) read the same. Without the
                        // balance flag the character takes the WHOLE charSpace like a
                        // full-width one (reference__0cf9c879, charSpace -2880: digits
                        // 4.80 = 5.5 - 0.70; one witness -- opt-in OXI_S1337B=1).
                        // The natural 5.5 packed 22 characters into p4's
                        // 「25年４月から「障害者自立支援法」が「障害」 line (Word: 21).
                        let default_fs = pitch / ratio;
                        let char_space_pt = pitch - default_fs;
                        if self.balance_single_byte_double_byte_width {
                            let cell = if char_space_pt >= 0.0 {
                                font_size * pitch / default_fs
                            } else {
                                font_size + char_space_pt
                            };
                            if std::env::var("OXI_DBG1337").is_ok() {
                                eprintln!("[S1337] ch={:?} fs={:.2} pitch={:.3} ratio={:.4} default_fs={:.2} cs={:.3} cell={:.2} natural={:.2}",
                                    ch, font_size, pitch, ratio, default_fs, char_space_pt, cell, char_width);
                            }
                            // S1347: half the scaled cell (see the full-width branch)
                            let cell = if (s1347_scale - 1.0).abs() > 1e-6 {
                                s1347_scale * (cell + font_size) / 2.0 + (cell - font_size) / 2.0
                            } else {
                                cell
                            };
                            // S1350 (2026-09-07): ...and half the TRACKED cell -- 0ea3ec86
                            // p9's -8 run draws 「(20」 at 5.40 / 5.28 / 5.40 (half of
                            // 10.69), the faithful slice the same; 5.75 put the line's
                            // 「る。」 on a second line and a page onto the rest.
                            let s1350_track = if std::env::var("OXI_S1340_DISABLE").is_err()
                                && std::env::var("OXI_S1350_DISABLE").is_err()
                                && style.fit_text.is_none()
                                && !style.ruby_spread
                            {
                                style.character_spacing.unwrap_or(0.0)
                                    * if self.balance_single_byte_double_byte_width { 2.0 } else { 1.0 }
                            } else {
                                0.0
                            };
                            0.5 * (cell + s1350_track) - char_width
                        } else {
                            char_space_pt
                        }
                    } else if ratio > 0.0
                        && pitch > 0.0
                        && char_width > 0.0
                        && !kinsoku::is_cjk(ch)
                        && !crate::font::is_fullwidth(ch)
                        && !use_east_asia
                        && char_width < 0.98 * font_size
                        && !ch.is_whitespace()
                        && style.fit_text.is_none()
                        && !style.ruby_spread
                        && std::env::var_os("OXI_S1449_DISABLE").is_none()
                    {
                        // S1449 (2026-09-17, default ON, opt-out OXI_S1449_DISABLE): on a
                        // `linesAndChars` grid a PROPORTIONAL half-width character carries
                        // the grid's charSpace too — half of it when the document sets
                        // `balanceSingleByteDoubleByteWidth`, all of it when it does not.
                        // S1337 above only reached an ASCII glyph that is already half an
                        // em wide (an MS-Mincho-style Latin), so a proportional face kept
                        // its natural width. COM (tools/metrics/_pb_halfcharspace_gen.py,
                        // tests/fixtures/halfcharspace): Century 10.5 over charSpace
                        // -0.862, 'm(μ)Gy' spans 35.25 without the setting and 37.50 with
                        // it (= natural 30.14 - 5 x 0.862 or - 5 x 0.431, plus the
                        // full-width mu 9.638), matching technical__9e4d04b4's cell
                        // character for character; kerning, justification, a first-line
                        // indent and being inside a cell change nothing.
                        let default_fs = pitch / ratio;
                        let char_space_pt = pitch - default_fs;
                        char_space_pt
                            * if self.balance_single_byte_double_byte_width {
                                0.5
                            } else {
                                1.0
                            }
                    } else {
                        0.0
                    }
                } else {
                    0.0
                };
                // S239 (2026-05-23): removed OXI_LEGACY_GRID_KERN legacy
                // env-var-only branch (was `else if let Some(pitch) = grid_char_pitch`
                // gated entirely by env::var().is_ok()).
                // For negative extras, fold into char_width directly so fragment
                // widths (positioning) reflect the shrink. For positive extras,
                // keep the existing separate-accumulator model (padding for positioning).
                if char_grid_extra < 0.0 {
                    char_width += char_grid_extra;
                    // S1492 (2026-09-20, default ON, opt-out OXI_S1492_DISABLE): the
                    // S475 break CAPACITY is accumulated from `pre_yakumono_width`,
                    // which is taken before the grid fold -- so a docGrid that
                    // COMPRESSES its characters (negative w:charSpace) was rendered at
                    // the compressed advance but broken at the natural em.
                    // legal__08a3b60be53504c7 (linesAndChars, charSpace=-2714, 12pt on a
                    // 9.837 pitch): Word fits 40 characters on the 453.5pt line, Oxi 38
                    // (capw += 240 per character where the fragment advances 227), so
                    // every full paragraph took one line too many and the form ran to 4
                    // pages against Word's 3.
                    if std::env::var_os("OXI_S1492_DISABLE").is_none() {
                        pre_yakumono_width += char_grid_extra;
                    }
                } else if s466_grid_expand && char_grid_extra > 0.0 {
                    // S466: fold POSITIVE grid expansion into char_width so the wrap
                    // (chars/line) reflects Word's grid-pitch advance. Default-OFF
                    // keeps the legacy separate-accumulator positioning behavior.
                    char_width += char_grid_extra;
                    // S1584 (2026-09-27, default ON, opt-out OXI_S1584_DISABLE): the
                    // S1492 fold for an EXPANDING grid too. The S475 capacity was
                    // accumulated at the bare em (210 tw at 10.5pt) while the
                    // character advances the 10.84pt cell (217 tw): `_pb_hang_bracket_gen.py`
                    // (linesAndChars 350/1382, jc=left) -- Word holds 41 あ on the
                    // 453.5pt line, Oxi 42.
                    // Scoped to explicit compat 15 (the probe's mode): reference__13e1b7fc
                    // (compat 14, charSpace 3194) packs its bracket-heavy lines at the bare
                    // em in Word, and this fold broke them a character early.
                    if std::env::var_os("OXI_S1584_DISABLE").is_none()
                        && self.compat_mode >= 15
                        && self.compat_mode_explicit
                    {
                        pre_yakumono_width += char_grid_extra;
                    }
                }
                // S1317 (2026-09-05, default ON, opt-out OXI_S1317_DISABLE): a
                // grid-pitched character advances the line by the TRUE pitch,
                // accumulated -- Word's device origins (600 dpi in its PDF) sit at
                // round(i x pitch), so the advances of ＭＳ 明朝 11pt under
                // charSpace -2880 alternate 10.20 / 10.32 around 10.297 and 44 of
                // them fit the S1211C floor of 44 cells EXACTLY. Rounding each
                // char to whole twips first (10.297 -> 206tw = 10.30) drifts
                // +0.0625tw per char: 44 x 206 = 9064 > 9061, the 44th char
                // wraps and reference__0cf9c879 spills to a 2nd page (Word: 1).
                // The twips accumulator therefore takes round(cum + pitch) -
                // round(cum) for such a character (integer comparisons kept).
                // A synthesized bold advance is a fractional font unit;
                // retain that precision across the line as well.
                let s1317_grid_char = ((char_grid_extra < 0.0
                    || (s466_grid_expand && char_grid_extra > 0.0))
                    && std::env::var("OXI_S1317_DISABLE").is_err())
                    || char_metrics.synthetic_bold_advance > 0.0;
                let s1317_inc = |cum: f32, cw: f32| -> i32 { pt_to_tw(cum + cw) - pt_to_tw(cum) };

                if ch == ' ' || ch == '\t' || ch == '\n' || ch == '\x0C' || ch == '\x0B' {
                    // Whitespace: flush word, then handle the whitespace
                    flush_word!(style);

                    if ch == '\n' || ch == '\x0C' || ch == '\x0B' {
                        // Set break type on the current line before pushing
                        let break_type = match ch {
                            '\x0C' => LineBreakType::PageBreak,
                            '\x0B' => LineBreakType::ColumnBreak,
                            '\n' => LineBreakType::SoftBreak,
                            _ => LineBreakType::Normal,
                        };
                        if ch == '\n' && current_line.fragments.is_empty() {
                            current_line.empty_break_style = Some((*style).clone());
                        }
                        current_line.break_type = break_type;
                        current_line.break_source = Some((frag_run_index, char_pos_in_run, (*style).clone()));
                        lines.push(std::mem::take(&mut current_line));
                        current_width = 0.0;
                        current_width_tw = 0;
                        current_capw_tw = 0;
                        latin_space_credit_tw = 0;
                    latin_space_credit_remainder = 0.0;
                        right_tab_slack_tw = 0;
                        center_tab_stop_tw = None;
                        compress_used = false;
                    } else {
                        // Space or tab
                        if ch == '\t' {
                            // S885 (2026-07-16, default ON, opt-out OXI_S885_DISABLE):
                            // content under a RIGHT/CENTER tab ends AT its stop
                            // (right-aligned) or centered ON it — the provisional
                            // left-advance (stop + content width) overstates the line
                            // position, which (a) skips the S881 implied hanging stop
                            // whenever stop + w crosses ind_left, and (b) resolves the
                            // next tab from a phantom position. Correct current_width
                            // to the aligned end before resolving this tab (the S841
                            // post-pass already fixes the RENDER positions; this fixes
                            // the BREAK-time resolution). legal__0001482d Defpara2
                            // 「⇥(ii)⇥it is classified…」: right@102.05 + (ii) 14.66 →
                            // provisional 116.71 crossed the hanging stop 116.3 → fell
                            // to the defaultTabStop grid (Word tabs to 116.3).
                            if let Some((stop_rel, cw_before, cw_after, align, njump)) =
                                rc_prev_tab.take()
                            {
                                if njump == lines.len()
                                    && !self.doc_body_has_real_cjk
                                    && std::env::var("OXI_S885_DISABLE").is_err()
                                {
                                    let w_content = current_width - cw_after;
                                    if w_content >= 0.0 {
                                        let true_cw = match align {
                                            TabStopAlignment::Center => (stop_rel
                                                + w_content / 2.0)
                                                .max(cw_before + w_content),
                                            _ => stop_rel.max(cw_before + w_content),
                                        };
                                        if true_cw < current_width - 0.01 {
                                            current_width = true_cw;
                                            current_width_tw = pt_to_tw(current_width);
                                            current_capw_tw = current_width_tw;
                                        }
                                    }
                                }
                            }
                            // COM-confirmed: tab positions are absolute from left margin.
                            // current_width is relative to the indent start, so we add
                            // indent_left to get the absolute position from margin.
                            // S1349: a chars-given indent (leftChars / hangingChars) is
                            // an indent here too -- db9ca183 p3 「a. National
                            // government (If …」 tabs from 0 without it and holds 88
                            // characters where Word holds 78.
                            let indent_left = para_style
                                .indent_left
                                .or_else(|| self.s1349_left_pt_style(para_style, fragments.first().map(|f| f.1), grid_char_pitch, grid_char_cw_ratio))
                                .unwrap_or(0.0);
                            // S881 (re-derived 2026-07-16, default ON, opt-out
                            // OXI_S881_DISABLE): a HANGING indent creates an IMPLIED
                            // tab stop at indent_left (Word merges it into the sorted
                            // stop list). ★The prior opt-in v1/v2 failed by re-basing
                            // line_start_abs at ind_left+fli while `current_width` is
                            // ALREADY SEEDED with first_line_indent (12729) — abs_pos
                            // was correct all along; re-basing double-counted |fli| and
                            // wrapped every clause early. The ONLY missing piece is the
                            // implied stop in the CANDIDATE LIST, including when
                            // explicit stops exist but are exhausted (all <= abs_pos):
                            // legal__0001482d Defpara 「⇥(b)⇥that is not ammunition…」
                            // (right@66.6 + ind left=80.8 hanging=80.8) — after the
                            // right stop, Word tabs to the hanging stop 80.8 (text x
                            // 201.0); Oxi fell to the defaultTabStop grid 100.1 →
                            // line 1 was 19.3pt narrower → 90 paragraphs doc-wide
                            // wrapped one line early (each ~+13pt, net +1178pt ≈ the
                            // doc's +2 pages). The effective (seeded) first_line_indent
                            // param is the gate — list_consumes_hanging paras zero it,
                            // so numbering suffix-tab machinery is untouched. Scope:
                            // Latin (!doc_body_has_real_cjk); the hanging+literal-tab
                            // pattern is 0/1439 golden-test + 0 docx_corpus/ja → JP
                            // byte-identical by construction.
                            // S1560 (2026-09-26, default ON, opt-out OXI_S1560_DISABLE): the
                            // implied stop is not Latin-only. legal__09904427 p2 (JA, TOC 1
                            // style: ind left=310 hanging=310, one right dot-leader stop at
                            // 9072): 「1.⇥臨床研究の名称…⇥2」 -- Word tabs the title to the
                            // hanging position 15.5pt (PDF: continuation lines at 77.9 =
                            // 62.4 + 15.5, one line per entry); Oxi jumped the first tab
                            // to the right stop, drew the leader after 「1.」 and wrapped
                            // the title to a second line.
                            let s881 = std::env::var("OXI_S881_DISABLE").is_err()
                                && (!self.doc_body_has_real_cjk
                                    || std::env::var_os("OXI_S1560_DISABLE").is_none())
                                && first_line_indent < -0.01;
                            // S1636 (2026-10-02, default ON, opt-out OXI_S1636_DISABLE): a
                            // line set in a lane beside a float starts `lane_shift` right
                            // of the paragraph's left edge; tab stops keep the margin
                            // origin, so a tab inside the lane jumps to the next stop
                            // measured from there (09422f63's heading «⇥くしゃみは…»: Word
                            // tabs to the default stop 504.55 or, when that leaves no
                            // room, moves the whole line below the box; Oxi added the
                            // stop to the lane start and set the text at x 618).
                            let line_start_abs = indent_left
                                + if std::env::var_os("OXI_S1636_DISABLE").is_none() {
                                    self.s1636_lane_shift.get()
                                } else {
                                    0.0
                                };
                            let abs_pos = current_width + line_start_abs;
                            if dbg_frags.is_some() {
                                eprintln!("[TAB] indent_left={:.2} cur={:.2} abs_pos={:.2} s881={} first_indent={:.2} line={}", indent_left, current_width, abs_pos, s881, first_line_indent, lines.len());
                            }
                            // Continuation lines start at rel=0 → abs_pos == ind_left,
                            // so the `>` guard auto-scopes the implied stop to line 1.
                            let implied_stop = if s881 && indent_left > abs_pos + 0.01 {
                                Some(indent_left)
                            } else {
                                None
                            };
                            let explicit_stop = para_style
                                .tab_stops
                                .iter()
                                .find(|ts| ts.position > abs_pos + 0.01);
                            let (next_pos, tab_align) = match (explicit_stop, implied_stop) {
                                (Some(ts), Some(imp)) if imp < ts.position - 0.01 => {
                                    (imp, TabStopAlignment::Left)
                                }
                                (Some(ts), _) => (ts.position, ts.alignment),
                                (None, Some(imp)) => (imp, TabStopAlignment::Left),
                                (None, None) => {
                                    let tab_stop = self.default_tab_stop;
                                    (
                                        ((abs_pos / tab_stop).floor() + 1.0) * tab_stop,
                                        TabStopAlignment::Left,
                                    )
                                }
                            };
                            // Convert absolute tab position back to relative width
                            let mut next_relative = next_pos - line_start_abs;
                            // S883 (2026-07-16, default ON, opt-out OXI_S883_DISABLE):
                            // a RIGHT-tab stop BEYOND the line's wrap boundary is a
                            // stale stop — Word CLAMPS it to the boundary and the tab's
                            // content does NOT force a wrap. legal__0001482d's TOC8:
                            // stop 340.2 vs boundary 297.65 (content − ind_right); Oxi
                            // jumped to the stop, so the page number '1' flush wrapped
                            // whenever titleW > avail − w('1') — the measured Oxi flip
                            // window (221.16, 223.32] vs Word's (226.35, 230.34] ∋
                            // avail 226.75 differs by EXACTLY w('1') 5.5pt (209 ToC
                            // records; #97=220.26pt / #112=223.32pt are a perfect 3pt
                            // A/B pair, both 1 line in Word). Clamp + exempt the tab
                            // content from the overflow check (the right-aligned
                            // segment pulls LEFT from the stop; it never extends the
                            // line). Latin scope; the ToC over-wrap is the first-2-page
                            // ×4 (+12.6pt each) driver of legal's +1 band.
                            // Legacy compatibility preserves aligned tab stops outside
                            // the paragraph boundary; modern layout clamps them.
                            // S1401 (2026-09-15, default ON, opt-out OXI_S1401_DISABLE; the
                            // checkpoint's opt-in OXI_CJK_ALIGNED_TAB_BOUNDARY promoted):
                            // MEASURED (`_pb_tabbeyond_gen.py`, Word COM Information(5) of
                            // the page number, A4 text width 425.25, right tab 9360 = 468pt):
                            //   compat 15, no indent      number ends at the right margin
                            //   compat 15, ind left 36/72 ends at margin - 36 / - 72
                            //   compat 15, ind right 18   ends at margin - 18
                            //   compat 15, stop AT 8505   one line (the at-boundary stop
                            //                             clamps like a beyond one)
                            //   compat 14 / no settings   ends at the 468pt stop, past the
                            //                             margin, as S883's Latin docs do
                            // one line in every arm. The clamp is the paragraph's usable
                            // width (content - ind_left - ind_right) laid off from the
                            // LEFT MARGIN, i.e. line-relative avail - ind_left; the
                            // checkpoint's avail alone missed the indented TOC levels.
                            // policies__07543a6b9776a1cf: 35 TOC entries (Word's default
                            // US-Letter 9360 stop on an A4 sheet) each wrapped their page
                            // number to a 2nd line, doubling the TOC, one page over.
                            let s1401_target = (available_tw as f32 / 20.0 - indent_left).max(0.0);
                            let cjk_aligned_tab_boundary = self.doc_body_has_real_cjk
                                && self.compat_mode_explicit
                                && self.compat_mode >= 15
                                && std::env::var("OXI_S1401_DISABLE").is_err()
                                && matches!(tab_align, TabStopAlignment::Right | TabStopAlignment::Center)
                                && next_relative > s1401_target - 0.01;
                            let tab_align = if cjk_aligned_tab_boundary {
                                TabStopAlignment::Right
                            } else {
                                tab_align
                            };
                            let mut s883_nowrap = false;
                            if tab_align == TabStopAlignment::Right
                                && std::env::var("OXI_S883_DISABLE").is_err()
                                && (!self.doc_body_has_real_cjk || cjk_aligned_tab_boundary)
                            {
                                let avail_rel = if cjk_aligned_tab_boundary {
                                    s1401_target
                                } else {
                                    available_tw as f32 / 20.0
                                };
                                if cjk_aligned_tab_boundary {
                                    next_relative = next_relative.min(avail_rel);
                                    s883_nowrap = true;
                                } else if next_relative > avail_rel + 0.01 {
                                    next_relative = avail_rel;
                                    s883_nowrap = true;
                                }
                            }
                            let next_pos = if s883_nowrap {
                                next_relative + line_start_abs
                            } else {
                                next_pos
                            };
                            // S1044 (2026-07-30, default ON, opt-out OXI_S1044_DISABLE):
                            // a LEFT tab whose
                            // resolved stop lies BEYOND the line's right boundary does
                            // not fit — Word breaks the line AT that tab and re-resolves
                            // it from the next line's start. Oxi placed every left tab
                            // unconditionally (S883 clamps only RIGHT tabs), so trailing
                            // tabs never wrapped and a whole line went missing.
                            // ★DERIVED (tab_overflow_probe, Word PDF, 10 arms; tabs paint
                            // a 3pt space glyph at their START so the landing of each tab
                            // is readable): 70-underscore arms end at 527.41 with the
                            // boundary at 540 and a 36pt default grid — 1 tab → the tab
                            // itself wraps (its stop 540 is placeable but the following
                            // text no longer fits, so the greedy break lands on the tab);
                            // 2 tabs → tab1 stays at 540 and tab2 (stop 576 > 540) wraps,
                            // landing at 108 on line 2; 4 tabs → tabs 2-4 land at
                            // 108/144/180. Identical for jc=left/both and with the target's
                            // pBdr. forms__002a64445e58ed78 para 26 (4 trailing tabs,
                            // text ending 522.7) reproduces exactly: tab1 → 540, tab2 →
                            // 576 overflows → line 2 holds only tab whitespace at x0=72,
                            // which is the 12pt line Word renders at bl=449.590 and Oxi
                            // omitted (the paragraph's −13.816pt junction error).
                            // ★A tab always ADVANCES TO AT LEAST THE BOUNDARY; it wraps
                            // only when it cannot advance at all, i.e. the line is
                            // already AT the boundary. Measured discriminator (both
                            // sides Word truth):
                            //   target para 26        cw 468.00 == avail 468.00, stop 504 → WRAP
                            //   forms__00042714       cw 529.20 <  avail 558.00, stop 562.50
                            //                         → Word CLAMPS (its line ends at
                            //                           571.61 ≈ the 571.5 boundary), no wrap
                            // and every probe arm agrees: L2/L4/E4 wrap because tab1 already
                            // landed exactly ON the boundary, while S1/S4 (room left) and L1
                            // (room left) do not wrap AT THE TAB. Without the at-boundary
                            // guard the rule fires on a tab that still has room and costs
                            // forms__00042714 its PASS (+1 page).
                            // Latin scope, like the sibling tab rules S881/S883/S885/TABTW.
                            let mut next_pos = next_pos;
                            let mut next_relative = next_relative;
                            let mut tab_align = tab_align;
                            // S1251 (default ON, opt-out OXI_S1251_DISABLE): S1044's at-boundary guard
                            // misses a TRAILING tab. Measured on the parts of
                            // legal__001a2c7f07cd358f (same styles/numbering/settings,
                            // only the paragraph tail swept):
                            //   text + tab          2 lines   (the tab wraps)
                            //   text + tab + "X"    1 line    (the tab clamps, X follows)
                            //   text + tab + tab    2 lines
                            //   text                1 line
                            // i.e. a tab whose stop is past the boundary wraps when
                            // NOTHING follows it, and clamps when content does -- the
                            // line need not already be full. The document's own line is
                            // then no longer the last, so Word justifies it out to the
                            // boundary; Oxi kept one line, left it unjustified 20pt
                            // short, and the tab "fitted" -- a self-consistent but wrong
                            // fixpoint. Only a space and a tab count as "nothing follows":
                            // U+00A0 renders, so `char::is_whitespace` is too broad.
                            let s1251_trailing_tab =
                                std::env::var("OXI_S1251_DISABLE").is_err()
                                    && s1251_rest_blank
                                    && s1251_last_content
                                        .map_or(true, |i| i <= char_index);
                            if std::env::var("OXI_S1044_DISABLE").is_err()
                                && !self.doc_body_has_real_cjk
                                && !s883_nowrap
                                && tab_align == TabStopAlignment::Left
                                && !current_line.fragments.is_empty()
                                && next_relative > available_tw as f32 / 20.0 + 0.01
                                && (current_width >= available_tw as f32 / 20.0 - 0.01
                                    || s1251_trailing_tab)
                            {
                                if std::env::var("OXI_DBG1044").is_ok() {
                                    let head: String = current_line
                                        .fragments
                                        .iter()
                                        .map(|f| f.text.as_str())
                                        .collect::<String>()
                                        .chars()
                                        .take(40)
                                        .collect();
                                    eprintln!(
                                        "[S1044] cw={:.2} next_rel={:.2} avail={:.2} \
indent_l={:.2} fli={:.2} stops={} | {:?}",
                                        current_width,
                                        next_relative,
                                        available_tw as f32 / 20.0,
                                        indent_left,
                                        first_line_indent,
                                        para_style.tab_stops.len(),
                                        head
                                    );
                                }
                                lines.push(std::mem::take(&mut current_line));
                                current_width = 0.0;
                                current_width_tw = 0;
                                current_capw_tw = 0;
                                latin_space_credit_tw = 0;
                    latin_space_credit_remainder = 0.0;
                                right_tab_slack_tw = 0;
                                center_tab_stop_tw = None;
                                compress_used = false;
                                rc_prev_tab = None;
                                // Re-resolve from the fresh line start (current_width 0).
                                // The S881 implied hanging stop is line-1 only and its
                                // `indent_left > abs_pos` guard excludes it here.
                                let abs2 = line_start_abs;
                                let (np, ta) = match para_style
                                    .tab_stops
                                    .iter()
                                    .find(|ts| ts.position > abs2 + 0.01)
                                {
                                    Some(ts) => (ts.position, ts.alignment),
                                    None => {
                                        let t = self.default_tab_stop;
                                        (((abs2 / t).floor() + 1.0) * t, TabStopAlignment::Left)
                                    }
                                };
                                next_pos = np;
                                tab_align = ta;
                                next_relative = np - line_start_abs;
                            }
                            // A left tab advances to its stop even when the gap is
                            // smaller than the fallback glyph width. Inflating a short
                            // gap overshoots the stop and can wrap a later trailing tab.
                            let min_tab_width = if tab_align == TabStopAlignment::Left {
                                0.0
                            } else {
                                char_width
                            };
                            let w = (next_relative - current_width).max(min_tab_width);
                            current_line.fragments.push(LineFragment {
                                auto_space_shrink: 0.0,
                                text: TAB_STRING.to_owned(),
                                width: w,
                                natural_width: w,
                                style: style.clone(),
                                tab_alignment: Some(tab_align),
                                tab_position: Some(next_pos),
                                field_type: None,
                                run_index: frag_run_index,
                                char_offset: char_pos_in_run,
                            });
                            // S885: remember a right/center tab so the NEXT tab (and
                            // the implied-stop guard) resolves from the aligned end.
                            let s885_cw_before_jump = current_width;
                            current_width += w;
                            rc_prev_tab = match tab_align {
                                TabStopAlignment::Right | TabStopAlignment::Center => Some((
                                    next_relative,
                                    s885_cw_before_jump,
                                    current_width,
                                    tab_align,
                                    lines.len(),
                                )),
                                _ => None,
                            };
                            // TABTW (2026-07-10, default ON, opt-out OXI_TABTW_DISABLE):
                            // a tab re-anchors the position at its ABSOLUTE tab stop, but
                            // only the float track advanced — current_width_tw (the track
                            // EVERY overflow check uses) stayed at the pre-tab text width,
                            // so a line with tabs followed by long text wrapped LATE by
                            // the whole tab advance (nyserda «[CONTRACTOR] ⇥⇥⇥⇥⇥ NEW YORK
                            // STATE…» ran ~50pt past the right margin, 2 lines vs Word 3).
                            // Re-anchor the tw track at the exact post-tab position; the
                            // capacity track too (compression of pre-tab punctuation
                            // cannot move post-tab text — the tab gap absorbs it).
                            // ★SCOPE !doc_body_has_real_cjk (the LATINEM discriminator):
                            // on JP justified-CJK tab lines the missing tab-tw was
                            // COMPENSATING the under-credited 約物 capacity (ohnochingin
                            // 第１５条 line: Word fits to the margin 518.8 by compressing
                            // ~18pt Oxi's capacity model doesn't credit — counting the
                            // tab without that credit over-wraps → PASS→FAIL). The JP
                            // side needs tab-tw + the full per-line 約物 credit TOGETHER
                            // (the char-budget wall); until then JP keeps the
                            // compensating pair, byte-identical by construction.
                            if (!self.doc_body_has_real_cjk || cjk_aligned_tab_boundary)
                                && std::env::var("OXI_TABTW_DISABLE").is_err()
                            {
                                current_width_tw = pt_to_tw(current_width);
                                current_capw_tw = current_width_tw;
                                // S774: a RIGHT tab's jump is slack the following
                                // segment consumes leftward (see the declaration).
                                // A new tab of any kind resets the prior slack.
                                center_tab_stop_tw = if tab_align == TabStopAlignment::Center
                                    && std::env::var("OXI_S958_DISABLE").is_err()
                                {
                                    // A centered segment cannot consume the separating space
                                    // after preceding text when it is clamped leftward.
                                    let separator = if s885_cw_before_jump > 0.0 {
                                        pt_to_tw(latin_metrics.char_width_pt(' ', font_size))
                                    } else { 0 };
                                    Some((current_width_tw, (pt_to_tw(w) - separator).max(0)))
                                } else {
                                    None
                                };
                                right_tab_slack_tw = if s883_nowrap {
                                    // S883: a clamped out-of-bounds right tab's content
                                    // never wraps the line (right-aligned content pulls
                                    // left from the boundary).
                                    i32::MAX / 4
                                } else if tab_align == TabStopAlignment::Right
                                    && std::env::var("OXI_S774_DISABLE").is_err()
                                {
                                    pt_to_tw(w)
                                } else {
                                    0
                                };
                            } else if tab_align == TabStopAlignment::Right
                                && std::env::var_os("OXI_S1567_DISABLE").is_none()
                                && std::env::var("OXI_S774_DISABLE").is_err()
                            {
                                // S1567 (2026-09-26, default ON, opt-out OXI_S1567_DISABLE):
                                // the CJK side keeps TABTW's compensating pair (the tw
                                // track does not re-anchor), so the tab jump rides in the
                                // NEXT word's fit width -- and without S774's slack a
                                // right-aligned page number after a stop just inside the
                                // boundary wrapped. policies__1a7a3fec p3 (toc 1, right
                                // dot stop 415.15 on a 415.65 line): every chapter entry's
                                // '3' went to a second line (Word: one line per entry,
                                // COM Info(6) 247.5 / 274.5 / 469.5). The slack is the
                                // jump itself: the segment pulls left from the stop.
                                right_tab_slack_tw = pt_to_tw(w);
                                center_tab_stop_tw = None;
                            }
                        } else {
                            // Regular space
                            current_line.fragments.push(LineFragment {
                                auto_space_shrink: 0.0,
                                text: SPACE_STRING.to_owned(),
                                width: char_width,
                                natural_width: char_width,
                                style: style.clone(),
                                tab_alignment: None,
                                tab_position: None,
                                field_type: None,
                                run_index: frag_run_index,
                                char_offset: char_pos_in_run,
                            });
                            // Preserve fractional synthetic-bold advances across spaces.
                            let space_tw = if char_metrics.synthetic_bold_advance > 0.0 {
                                pt_to_tw(current_width + char_width) - pt_to_tw(current_width)
                            } else {
                                pt_to_tw(char_width)
                            };
                            current_width += char_width;
                            current_width_tw += space_tw;
                            current_capw_tw += space_tw; // S475: space, no punct capacity
                            s1475_space_tw = space_tw;
                            if dbg_flush {
                                eprintln!(
                                    "[DBGSPACE] w={:.3} fam={} fs={} has32={} em32={:.4}",
                                    char_width,
                                    char_metrics.family,
                                    font_size,
                                    char_metrics.char_widths.contains_key(&' '),
                                    char_metrics.char_widths.get(&' ').copied().unwrap_or(-1.0)
                                );
                            }
                            if kernbreak_para && ch == ' ' {
                                latin_space_credit_tw += pt_to_tw(char_width * kernbreak_cap);
                            } else if s799_space_shrink && ch == ' ' {
                                // S825 (2026-07-13, opt-out OXI_S825_DISABLE): the
                                // COMPAT-15 justified space-shrink capacity, DERIVED
                                // (_pb_cs3_gen m15 sweeps, 4 cs profiles, fits ±0.03
                                // ..0.3pt over 10 spaces):
                                //   capacity/space = 0.365×em + 0.24×cs
                                // — the em space compresses to ~63.5% and baked
                                // w:spacing (cs) to ~76% of itself. WITHOUT a
                                // compat-15 settings part the allowance collapses to
                                // a small per-LINE constant (fs/4, cs-independent —
                                // the same sweeps without settings.xml) — the legacy
                                // 0.10 approximation stays for that regime. The
                                // enabler was pinned by adding ONLY the compat-15
                                // settings to the synthetic (allow 2.74 → ≥7.7);
                                // falsified en route: substitution, pStyle, theme,
                                // numbering, docDefaults line=23, line index.
                                // S825b (2026-07-13): the em coefficient RE-DERIVED at
                                // 0.25 (quarter-space) — the original 0.365 was fit
                                // against METRICS-computed naturals that over-state
                                // Word's rendered width (Calibri-11 11-word line:
                                // Oxi metrics 455.99 vs Word rendered 453.19, +2.8pt
                                // → the derived allow inflated by the same). Fresh
                                // render-measured naturals (left-aligned line1
                                // extents) across TNR-12/TNR-11/Calibri-12/Calibri-11
                                // give allow/space = 0.243/0.266/0.240/0.253 ≈ 0.25 ×
                                // the em space — unifying with KERNBREAK's 0.25 cap
                                // (Word's justify shrink limit = space/4). The 0.24×cs
                                // term survives (cs deltas were RELATIVE in the pc3
                                // sweep, natural-inflation cancels). nyserda p24
                                // «…agrees to submit | to» (13 spaces, needed 12.9):
                                // 0.365-model granted 14.2 → over-fit (Word wraps);
                                // 0.25-model grants 9.75 → wraps = Word.
                                if c14_active
                                    && (char_metrics.char_width_em('i')
                                        - char_metrics.char_width_em('M'))
                                    .abs()
                                        < 0.001
                                {
                                    // C14JUST (monospace-scoped): accumulate the
                                    // half-em compression capacity per space (letters
                                    // unchanged). The fit consumers use this as the
                                    // budget DIRECTLY (S996: no alt_slack gate, no
                                    // hang) = the capacity rule. The half-em floor was
                                    // DERIVED on Courier (monospace); it OVER-fits
                                    // proportional fonts (reference__0029c1c Book
                                    // Antiqua / technical__00549a8f TNR went PASS→FAIL
                                    // under an ungated version), so a proportional
                                    // compat14 doc (i-width ≠ M-width) keeps its
                                    // original S933 fs/4 allowance below.
                                    latin_space_credit_tw += pt_to_tw(char_width * 0.5);
                                    c14_space_tw = pt_to_tw(char_width);
                                } else if self.compat_mode >= 15
                                    && self.compat_mode_explicit
                                    && std::env::var("OXI_S825_DISABLE").is_err()
                                {
                                    let em_part = char_width - cs;
                                    let em_credit = if std::env::var("OXI_SPACE_CREDIT_FIXED_DISABLE").is_err() {
                                        (em_part * 0.25 * 4096.0).round() / 4096.0
                                    } else { em_part * 0.25 };
                                    let credit = em_credit + cs.max(0.0) * 0.24 + latin_space_credit_remainder;
                                    let credit_tw = pt_to_tw(credit);
                                    latin_space_credit_tw += credit_tw;
                                    latin_space_credit_remainder = credit - credit_tw as f32 / 20.0;
                                } else if self.compat_mode_explicit
                                    && self.compat_mode <= 14
                                    && std::env::var("OXI_S1046_DISABLE").is_err()
                                {
                                    // S1046 (2026-07-30, default ON, opt-out
                                    // OXI_S1046_DISABLE): a doc that EXPLICITLY declares
                                    // a LEGACY compatibilityMode (<=14) and uses a
                                    // PROPORTIONAL font gets ~ZERO justify-shrink
                                    // allowance — the same as the settings-without-compat
                                    // class below. Such a doc previously fell through to
                                    // the S933 fs/4 arm, whose flat fs/4 was measured on
                                    // NO-SETTINGS hosts (the S825 booster-hunt probes had
                                    // no settings.xml at all); compat<=14-explicit was
                                    // never derived and only borrowed that value.
                                    // ★DERIVED on the real-doc population
                                    // (administrative__0021fbead5c6467d — a compat14
                                    // A4 justified TNR-12 report, 163 body lines):
                                    // comparing Oxi's line BREAKS to Word's PDF lines,
                                    //   fs/4 credit  ->  15 / 163 lines identical
                                    //   zero credit  -> 158 / 163 lines identical
                                    // and the line counts become [47,47,48,21] = Word
                                    // EXACTLY on all four pages. Traced boundary
                                    // (OXI_DBGFLUSH): «...Brussels on 28 | &» needs
                                    // 8169+187 = 8356 tw against avail 8300, and the
                                    // fs/4 credit of 60tw (3.0pt) wrongly absorbed the
                                    // 56tw overflow that Word wraps.
                                    // MONOSPACE compat14 keeps the C14JUST capacity rule
                                    // (that arm is earlier in this chain); compat>=15
                                    // keeps the S825 quarter-space.
                                    // credit: none.
                                } else if self.settings_part_exists
                                    && !self.compat_mode_explicit
                                    && std::env::var("OXI_S933_DISABLE").is_err()
                                {
                                    // S933b (2026-07-18): a doc whose settings.xml
                                    // EXISTS but declares NO compatibilityMode gets
                                    // ~ZERO justify-shrink allowance. In-doc derived
                                    // on legal__000ad039 (Supreme Court judgment,
                                    // settings w/o compatSetting, justified TNR-12):
                                    // Word wraps 'Plan' (needs <2pt) and 'The'
                                    // (needs 4.24pt) at natural width — the
                                    // OXI_S799_CAP=0 sweep lands the whole doc at
                                    // 0.9481/pcd 0 vs 0.5849 with fs/4, while
                                    // uklocalspending (compat14-EXPLICIT) stays
                                    // PASS 1.0 at either value. The fs/4 flat
                                    // allowance below is the NO-SETTINGS-part class
                                    // (the S825 booster-hunt cs3 hosts had no
                                    // settings.xml at all).
                                    // credit: none.
                                } else if std::env::var("OXI_S933_DISABLE").is_err() {
                                    // S933 (2026-07-18, default ON, opt-out
                                    // OXI_S933_DISABLE): the NON-explicit-compat15
                                    // class's per-LINE justify-shrink allowance is
                                    // the FLAT fs/4 the S825 booster hunt measured
                                    // (no-settings Calibri-11 probes: allow
                                    // 2.74-2.75 = fs/4, cs-independent) — the 0.10
                                    // per-space cap is a pre-derivation
                                    // approximation that OVER-credits space-rich
                                    // lines (grows ~0.3pt/space without bound).
                                    // legal__000ad039 (compat ABSENT, justified
                                    // TNR-12 judgment): its 14-space line granted
                                    // 84tw and fit 'The' with 0tw to spare where
                                    // Word wraps (needs 4.24 > fs/4 = 3.0) — one
                                    // wrap per ~page = the doc-wide −1 cascade.
                                    // Keep the per-space growth for space-poor
                                    // lines (unprobed; they rarely reach the
                                    // boundary) but clamp the line total at fs/4.
                                    let cap_tw = pt_to_tw(font_size * 0.25);
                                    latin_space_credit_tw = (latin_space_credit_tw
                                        + pt_to_tw(char_width * s799_cap))
                                    .min(cap_tw);
                                } else {
                                    latin_space_credit_tw += pt_to_tw(char_width * s799_cap);
                                }
                            }
                        }
                    }
                } else if is_break_after(ch) || (s1100_dash_break && s801_latin_dash) {
                    // Characters like '-', '/' that allow a line break AFTER them.
                    // S1100 adds the Latin-doc EM/EN DASH (s801_latin_dash, which
                    // already carries the `!doc_body_has_real_cjk` gate).
                    // Include them in the current word, flush, and allow a break.
                    if word_style.is_none() {
                        word_first_width_tw = pt_to_tw(char_width);
                        word_style = Some(style.clone());
                        word_field_type = frag_field_type;
                        word_run_index = frag_run_index;
                        word_char_offset = char_pos_in_run;
                    }
                    if latin_wordwrap && (word.is_empty() || seg_pending) {
                        word_seg_meta.push((word.chars().count(), frag_run_index, char_pos_in_run));
                        seg_pending = false;
                    }
                    word.push(ch);
                    if !ch.is_whitespace() {
                        s1026_nonws_consumed += 1;
                    } // S1026 final-token
                    word_width += char_width;
                    if latin_wordwrap {
                        word_char_ws.push(word_width);
                    } // S1059
                    word_natural_width += char_width + yakumono_saved;
                    if s809_hang || cjk_latin_period_hang {
                        // Existing Latin policy also admits comma/closing quotes;
                        // the separately measured mixed-text policy admits period.
                        word_trail_hang_w = if ch == '.'
                            || (s809_hang && ch == ',' && !s1630_no_comma)
                            || (s809_hang && s1262 && matches!(ch, '\u{201D}' | '\u{2019}'))
                        {
                            char_width
                        } else {
                            0.0
                        };
                    }
                    // S783 (2026-07-11): in a LATIN document a HYPHEN is a real
                    // word boundary — Word fills a partial line up to the '-'
                    // of a hyphenated compound (nyserda p13 'Other Co-' at line
                    // end x=512.66 < margin 522, 'funding' wraps). The
                    // opportunity-only model (derived on tokyoshugyo URLS after
                    // CJK text, where Word wraps the WHOLE token first) stays
                    // for '/'':' etc. and for CJK docs — hyphens there keep the
                    // URL behavior (JP byte-identical by construction).
                    // ★S1028 census (dcc, 2026-07-28): in the c14 MONOSPACE
                    // context Word does NOT fill up to the '-' of a compound —
                    // «majority-in-interest» / «152.201-152.206» wrap WHOLE
                    // (10/10 of the unified-rule misses were hyphen segments
                    // placed at the line end where Word wraps the compound).
                    // S783's fill-to-hyphen stays for the proportional class
                    // it was derived on (nyserda 'Other Co-'/'funding').
                    // S1399 (2026-09-14, opt-out OXI_S1399_DISABLE): the hyphen is a
                    // real word boundary in a CJK-body doc TOO. MEASURED
                    // (`_pb_urlbreak_gen.py`, Word COM per-char Information(6),
                    // MS Gothic 10.5 on a lines grid, «#　» + token): a 13-chunk
                    // «abcdefgh-…» token fills line 1 up to the 9th hyphen; the
                    // real URL fills line 1 to «notepad-plus-» with or without a
                    // kanji paragraph beside it; slash / dot / underscore tokens
                    // wrap WHOLE and margin-break (the tokyoshugyo URL model,
                    // which S783 mistook for a hyphen rule). technical__b80f6caa:
                    // the URL took 3 lines instead of 2, +18pt, one paragraph over.
                    let s1399 = std::env::var("OXI_S1399_DISABLE").is_err();
                    let s783_hyphen_flush = ch == '-'
                        && (s1399 || !self.doc_body_has_real_cjk
                            || (legacy_latin_hyphen
                                && !word.contains(':')
                                && !word.contains('/'))
                            || (self.compat_mode >= 15 && self.compat_mode_explicit
                                && std::env::var_os("OXI_CJK_HYPHEN_BREAK").is_some()))
                        && std::env::var("OXI_S783_DISABLE").is_err()
                        && !(c14_active
                            && c14_space_tw > 0
                            && std::env::var("OXI_S1028_HY_DISABLE").is_err());
                    // S1100: an EM/EN DASH is a REAL word boundary (Word fills a
                    // partial line up to it), exactly like S783's hyphen — the
                    // opportunity model below is the URL model («the whole token
                    // wraps to the next line first»), and the 76-arm sweep shows
                    // Word does NOT do that for a dash (NB3100: line 1 ends
                    // «…1300 mm —» while the NBSP-joined token would otherwise
                    // move whole). So dashes take the flush path.
                    let s1100_flush = s1100_dash_break && s801_latin_dash;
                    if latin_wordwrap && !s783_hyphen_flush && !s1100_flush {
                        // Record a break OPPORTUNITY (after this char) instead of forcing
                        // a flush — the maximal Latin token is kept together and only split
                        // by flush_word when it overflows a full line (Western word-wrap).
                        word_breaks.push((word.chars().count(), word_width));
                        seg_pending = true; // next char starts a new segment
                    } else {
                        flush_word!(style);
                    }
                } else if (kinsoku::is_cjk(ch) && !latin_ctx_quote
                    && !(ch as u32 == 0xFFE5 && chars_vec.get(char_index + 1).copied()
                        .or_else(|| fragments[frag_outer_idx + 1..].iter().find_map(|f| f.0.chars().next()))
                        .is_some_and(|next| next.is_ascii_digit())))
                    || (use_east_asia && matches!(ch, '\u{0370}'..='\u{04FF}'))
                {
                    // CJK characters always break at char boundaries (subject to kinsoku).
                    // ECMA-376 §17.3.1.40: wordWrap controls LATIN word-break only.
                    // V_JJ measurement (2026-05-02) confirmed: V_JJ2 (wordWrap=on) and
                    // V_JJ3 (wordWrap=off) produce identical CJK break points.
                    // Pre-2026-05-03: this branch was gated on `&& para_style.word_wrap`,
                    // causing CJK to accumulate as a single non-breakable word in
                    // wordWrap=off paragraphs (34 baseline docs / 108 instances).
                    // autoSpaceDE: add 2.5pt gap between Latin and CJK ideograph/kana.
                    // COM-confirmed (2026-04-07): only ideographs/kana trigger auto-space,
                    // not CJK punctuation (which gets no extra spacing from Latin).
                    // Session 95 (2026-05-18) split: autoSpaceDE gates ALPHABETIC
                    // boundaries, autoSpaceDN gates DIGIT boundaries. e3c545 has
                    // DE=on, DN=off → digits should NOT get the gap. Was previously
                    // gated on auto_space_de alone via is_ascii_alphanumeric().
                    let prev_alpha_local = !word.is_empty()
                        && word
                            .chars()
                            .last()
                            .map_or(false, |c| c.is_ascii_alphabetic());
                    let prev_digit_local = !word.is_empty()
                        && word.chars().last().map_or(false, |c| c.is_ascii_digit());
                    let (prev_frag_alpha, prev_frag_digit) = if word.is_empty() {
                        let last_char = current_line
                            .fragments
                            .last()
                            .and_then(|f| f.text.chars().last());
                        (
                            last_char.map_or(false, |c| c.is_ascii_alphabetic()),
                            last_char.map_or(false, |c| c.is_ascii_digit()),
                        )
                    } else {
                        (false, false)
                    };
                    let prev_is_alpha = prev_alpha_local || prev_frag_alpha;
                    let prev_is_digit = prev_digit_local || prev_frag_digit;
                    flush_word!(style);
                    let cur_is_cjk = kinsoku::is_cjk_ideograph_or_kana(ch);
                    // S1316 (2026-09-05, default ON, opt-out OXI_S1316_DISABLE): Word puts
                    // no auto-space on the side of a digit/Latin run that touches a
                    // RUBY field. DERIVED (`_pb_autospace_gen.py`, (資料代)50(円) with
                    // ruby on both / left / right / neither neighbour, body and cell):
                    // neither +2.52/+2.57, both -0.12/+0.05, left only 0/+2.57, right
                    // only +2.52/+0.17. correspondence__04a3e3e1's '(資料代50円)' cell
                    // line carries two ruby fields around '50': Word 0, Oxi +5.
                    let s1316_ruby_adjacent = std::env::var("OXI_S1316_DISABLE").is_err()
                        && (style.ruby_field
                            || current_line.fragments.last().map_or(false, |f| f.style.ruby_field));
                    if cur_is_cjk
                        && !s1316_ruby_adjacent
                        && ((prev_is_alpha && para_style.auto_space_de)
                            || (prev_is_digit && para_style.auto_space_dn))
                    {
                        // S546: gap = fs/4 true-space (old per-fontSize table = paint artifact).
                        let extra = current_line.fragments.last()
                            .map(|last| self.autospace_after_style(last.text.chars().last().unwrap_or(' '), &last.style, para_style))
                            .unwrap_or_else(|| s546_autospace_extra(font_size));
                        if let Some(last) = current_line.fragments.last_mut() {
                            last.width += extra;
                            last.natural_width += extra;
                            if legacy_gap_on || std::env::var_os("OXI_CJK_AUTOSPACE_COMPRESSION").is_some() {
                                last.auto_space_shrink += extra * 0.5;
                            }
                        }
                        // Keep quarter-em gaps fractional until the cumulative twip conversion.
                        let extra_tw = if std::env::var_os("OXI_CJK_AUTOSPACE_CUMULATIVE").is_some() {
                            pt_to_tw(current_width + extra) - pt_to_tw(current_width)
                        } else {
                            pt_to_tw(extra)
                        };
                        current_width += extra;
                        current_width_tw += extra_tw;
                        current_capw_tw += extra_tw; // S475: autoSpace, no punct capacity
                    }

                    // ORPHAN-OIKOMI gate (OXI_ORPHAN_OIKOMI=1, experiment, default OFF =
                    // byte-identical): on a paragraph's LAST line Word compresses 約物
                    // HARDER (raises the opening-bracket cap) to avoid orphaning the tail.
                    // nedo para 333 «…その子会社（…規定する子» is a ~1-line para ending «子»;
                    // Word fits «子» by compressing «、»(3.36)+«（»(3.31), while the
                    // MIDDLE-line wraps of 400/434/465 get NO extra compression (oidashi).
                    // Discriminator = paragraph-remaining-chars ≤ one line (last-line
                    // context): fires on «（» when it sits in the para's final line, not on
                    // a middle-line «、». OXI_ORPHAN_OPEN / OXI_ORPHAN_LINEMULT tune.
                    // S684 (2026-06-28, default ON, opt-out OXI_NPERIOD_DISABLE): the n_period
                    // opening gate — Word credits OPENING-bracket demand compression only when
                    // the line lacks cheap PERIOD (。．) half-em. Fixes nedo W1 (≥2 periods →
                    // openings light → 甲 wraps) AND 333 (0 periods → opening HI=3.4 → 子 fits).
                    // Pairs with OXI_TRAILBR (the trailing-<w:br/> empty line) to PASS nedo: the
                    // {−1:3} the memory called an 8-session char-budget cascade was the dropped
                    // <w:br/> empty line (found via the reliable article/page-top measurement).
                    // ★The all-types NPALL variant (≥2-period → comma/closing also light) was
                    // tried + FALSIFIED — it regressed b837/d77a/harassmanual/ohnoikuji; nedo
                    // only needs the OPENINGS gate (this) + TRAILBR. b837-SAFE (s475_break =
                    // type=lines only; b837 = linesAndChars). HI=3.4 fixes 333, LO=0 wraps W1.
                    let s475_open_eff =
                        if s475_break && std::env::var("OXI_NPERIOD_DISABLE").is_err() {
                            // n_period DISCRIMINATOR (2026-06-23, OXI_NPERIOD, default OFF):
                            // Word credits OPENING-bracket demand compression only when the line
                            // lacks cheap PERIOD (。．) half-em compression. MEASURED (nedo Word
                            // PDF, _nedo_open_trigger.py): of 35 demand-opening-compressed lines
                            // (aki>3.0), 31 have n_period=0, 4 have n_period=1, ZERO have
                            // n_period≥2. Periods supply 6.0pt "free" half-em, so Word uses them
                            // and leaves openings at baseline; with <2 periods Word dips into
                            // opening compression. nedo W1 (2 。, fit 甲 needs 18.8) wraps; i=334/
                            // para-333 (0 。, fit 子 needs the 社（ 3.31) compress the opening.
                            // hi = demand opening cap (fit 子/令), lo = period-rich opening cap
                            // (exclude → wrap, render-match). Env-tunable. See [[char_budget_wall]].
                            let line_periods = current_line
                                .fragments
                                .iter()
                                .flat_map(|f| f.text.chars())
                                .filter(|&c| matches!(c, '。' | '．'))
                                .count();
                            let hi: f32 = std::env::var("OXI_NPERIOD_HI")
                                .ok()
                                .and_then(|v| v.parse().ok())
                                .unwrap_or(3.4);
                            let lo: f32 = std::env::var("OXI_NPERIOD_LO")
                                .ok()
                                .and_then(|v| v.parse().ok())
                                .unwrap_or(0.0);
                            if line_periods >= 2 {
                                lo
                            } else {
                                hi
                            }
                        } else if s475_break && orphan_oikomi_on {
                            // SHORT-PARA discriminator: para 333 is a ~1-line para ending «子»;
                            // Word compresses 約物 to KEEP it 1 line (high value). The over-fit
                            // {400,434,465} cascade from a MULTI-line para's compression Word
                            // does NOT apply (losing a line there is fine). Gate the higher
                            // open cap to SHORT paragraphs (total ≤ ~1.x lines).
                            let para_total = chars_before_frag + chars_vec.len() + chars_after_frag;
                            let fw_tw = (font_size * 20.0) as i32;
                            let line_cap = if fw_tw > 0 {
                                ((available_tw as f32 / fw_tw as f32) * orphan_line_mult) as usize
                            } else {
                                0
                            };
                            if para_total <= line_cap.max(1) {
                                s475_open.max(orphan_open_cap)
                            } else {
                                s475_open
                            }
                        } else {
                            s475_open
                        };
                    let s475_capinc = if s475_break {
                        // Demand-aware breaker (2026-06-27, OXI_PERTYPE, default OFF =
                        // byte-identical): per-type 約物 break caps from the Word-PDF
                        // measurement — comma 3.4 / period 6.0 (half-em) / closing-solo
                        // 0.84 (light). b837-safe (s475_break = type=lines). Tunable.
                        let (period_pt, close_solo_pt, comma_pt) =
                            if std::env::var("OXI_PERTYPE").is_ok() {
                                (
                                    std::env::var("OXI_PT_PERIOD")
                                        .ok()
                                        .and_then(|v| v.parse().ok())
                                        .unwrap_or(6.0),
                                    std::env::var("OXI_PT_CLOSE")
                                        .ok()
                                        .and_then(|v| v.parse().ok())
                                        .unwrap_or(0.84),
                                    std::env::var("OXI_PT_COMMA")
                                        .ok()
                                        .and_then(|v| v.parse().ok())
                                        .unwrap_or(s475_solo),
                                )
                            } else if s1234_offdefault_light {
                                // S1235 (2026-08-26): per-CLASS caps for the legacy
                                // off-default regime, from the parttime Word-PDF
                                // squeezed-line census (~50 lines, 8pt body):
                                //   、/・ = 2.5@12 (1.67@8): packing lines take
                                //   ≤ −1.44/、 (第25条 L1 ×4) while ④解雇 L1
                                //   REFUSES a 1.83pt rescue from its single 、
                                //   (renders it natural, wraps 小 leaving 6.2pt)
                                //   → the solo cap sits in (2.16, 2.74)@12;
                                //   。 line-end = half-em (第24条 L2 4.26);
                                //   ）closing-solo = ~0.7em taken (5.66) → 3.6@12.
                                //   A comma cap of 0 was tried first and REFUTED:
                                //   ~50 lines under-packed (、 squeezes of −0.24
                                //   to −1.44 are everywhere in the census).
                                (6.0, 3.6, 2.5)
                            } else {
                                (s475_solo, s475_solo, s475_solo)
                            };
                        // S757 (2026-07-06, default ON, opt-out OXI_S757_DISABLE): a
                        // LINE-INITIAL opening bracket has NO compressible left-aki —
                        // the line-start position already absorbs it (JIS X 4051
                        // 行頭括弧). Crediting it over-fit kyodoken05's line-1
                        // «（所属機関名…» (fit け where Word wraps) = the S684-NPERIOD
                        // −0.0393 audit regression; nedo's fits (子/令/甲) are all
                        // MID-line openings → unaffected.
                        // Line-initial = nothing placed on this line yet (fragments
                        // empty + word buffer empty; current_width_tw is unusable —
                        // it is seeded with the first-line indent). The line-initial
                        // opening keeps the BASE cap (3.1) — exempt from the NPERIOD
                        // HI escalation, not zeroed: c7b9 outline_06 needs the base
                        // credit (a zero under-fit it −0.0065), kyodoken05 needs
                        // ≤ base (the HI 3.4 over-fit け, −0.0393).
                        let s757_open_eff = if kinsoku::is_yakumono_opening(ch)
                            && current_line.fragments.is_empty()
                            && word.is_empty()
                            && std::env::var("OXI_S757_DISABLE").is_err()
                        {
                            s475_open_eff.min(s475_open)
                        } else {
                            s475_open_eff
                        };
                        if std::env::var("OXI_DBG757").is_ok() && kinsoku::is_yakumono_opening(ch) {
                            eprintln!("[DBG757] ch={:?} cw_tw={} nfrag={} word_len={} capw_tw={} open_eff={:.2}",
                                ch, current_width_tw, current_line.fragments.len(), word.chars().count(), current_capw_tw, s757_open_eff);
                        }
                        let s1167_cap_raw = kinsoku::s475_max_compress_pt(
                            ch,
                            chars_vec.get(char_index + 1).copied(),
                            s475_pair,
                            comma_pt,
                            // S1235: no parttime line shows a compressed opening
                            // (（ renders 8.04 natural throughout) — zero the
                            // opening credit in the off-default regime.
                            if s1234_offdefault_light {
                                0.0
                            } else {
                                s757_open_eff
                            },
                            period_pt,
                            close_solo_pt,
                            font_size,
                        );
                        let s1167_cap = if s1167_cap_raw > 0.0 {
                            let em = s1167_em_ref(
                                &self.registry,
                                font_size,
                                &char_metrics,
                                gdi_map,
                            );
                            if std::env::var("OXI_DBG1167").is_ok()
                                && pre_yakumono_width < em - 0.01
                            {
                                eprintln!(
                                    "[DBG1167] ch={:?} nat={:.3} em_ref={:.3} fs={:.2} ratio={:.3}",
                                    ch, pre_yakumono_width, em, font_size,
                                    pre_yakumono_width / em
                                );
                            }
                            kinsoku::s475_aki_cap(s1167_cap_raw, pre_yakumono_width, em)
                        } else {
                            s1167_cap_raw
                        };
                        // S1235 (2026-08-26): the off-default regime carries a
                        // PER-LINE compression ceiling of 0.75em on top of the
                        // per-mark caps. FOUR hairline datapoints pin it
                        // (parttime Word PDF, 8pt → ceiling 6.0): 第24条 L1
                        // absorbs 5.90 (packs), two lines refuse at 6.01/6.04,
                        // and the 8.69-9.55 refusals follow. The earlier 0.6em
                        // (4.8@8) value was wrong — it would refuse 第24条's
                        // 5.90 — which is why the first budget attempt appeared
                        // refuted; the failure was the value plus the uniform
                        // comma cap, not the ceiling concept.
                        let s1167_cap = if s1234_offdefault_light && s1167_cap > 0.0 {
                            let budget_tw = pt_to_tw(0.75 * font_size);
                            let used_tw = current_width_tw - current_capw_tw;
                            s1167_cap.min(((budget_tw - used_tw).max(0) as f32) / 20.0)
                        } else {
                            s1167_cap
                        };
                        let capacity = pt_to_tw(pre_yakumono_width - s1167_cap);
                        if vertical && std::env::var("OXI_VERTICAL_PAIR_DISABLE").is_err() {
                            capacity.min(pt_to_tw(char_width))
                        } else {
                            capacity
                        }
                    } else {
                        0
                    };
                    // S725 (2026-07-03, default ON, opt-out OXI_S725_DISABLE):
                    // PARAGRAPH-FINAL char must fit NATURALLY — the s475 capacity
                    // credit models JUSTIFY compression, but a paragraph's LAST
                    // line is never justified (rendered left-aligned at natural
                    // advances), so a final char kept via credit produces a line
                    // the render can never compress: text overflows the right
                    // margin (probe twin L2 x_end 543.4 = 19pt past the 524.4
                    // margin) AND the paragraph under-counts a line vs Word
                    // (twin: Oxi 3 lines / Word 4 → ×50 paras = −2 pages).
                    // Word render-truth (proberubytwin): Word wraps the tail to
                    // a natural-width final line (ない。) with zero compression;
                    // OXI_S575_CAP=0 reproduces Word exactly — the credit is the
                    // sole cause. SCOPE = the MODERN (compat≥15 EXPLICIT)
                    // compressPunctuation s475 body arm only: legacy docs do the
                    // OPPOSITE (orphan-elimination tail compression, S721/S568 —
                    // Word 2010-mode pulls a short tail back via 約物 oikomi),
                    // and absent-compat docs lay out as legacy (S545). s476_grid
                    // (linesAndChars, b837) and s590 (legacy justified) keep
                    // their calibrated behavior.
                    let modern_cjk_line_end = std::env::var_os("OXI_CJK_MODERN_LINE_END").is_some()
                        && s476_body && !vertical && !lines_and_chars
                        && self.compat_mode_explicit && self.compat_mode >= 15
                        // Justified text may hang a mark when its preceding text
                        // fits at natural width. Compression credit must not also
                        // pull the preceding text beyond the line boundary.
                        && (!is_justified || current_width_tw > available_tw);
                    let s725_final_char = std::env::var("OXI_S725_DISABLE").is_err()
                        && !vertical
                        && self.compat_mode >= 15
                        && self.compat_mode_explicit
                        && !lines_and_chars
                        && !s476_grid
                        && !s590_legacy_just_cap
                        && s725_chars_after == 0
                        && char_index + 1 == chars_vec.len();
                    // S1317: this character's twips increment (cumulative for a
                    // grid-pitched character, per-char rounding otherwise).
                    let s1317_cw_tw = if s1317_grid_char {
                        s1317_inc(current_width, char_width)
                    } else {
                        pt_to_tw(char_width)
                    };
                    // S1318 v2: in the at-default legacy regime only a line-final
                    // mark may spend the line's mark capacity; any other
                    // overflowing character is refused (plain width test).
                    // S1318 v3: the cell is the grid pitch (char_width carries it
                    // for a grid-pitched character) or the em without a grid.
                    // S1346: half of the character's OWN cell -- a tracked (-8: 10.7) or
                    // scaled (w:w 90: 10.34) line's mark gives half of that, not half
                    // the section pitch (trackedge s-8_m22: む needs 5.4, the 、 has 5.35).
                    let s1318_half_cell_tw = pt_to_tw(0.5 * if char_width >= 0.75 * font_size {
                        char_width
                    } else {
                        grid_char_pitch.unwrap_or(font_size).max(font_size)
                    });
                    // S1318 v4 (2026-09-06): the at-default regime's pull-in, with the
                    // line's OWN elective capacity (every mark half an em, a pair's
                    // second member structural = free) instead of the S475 per-type
                    // caps. DERIVED with the faithful-slice probe
                    // `_pb_kinsokufinal_gen.py` (0ea3ec86's styles/settings, 2-column
                    // charSpace 2048, 17 arms) plus the document's lines:
                    //   normal X: pulled in iff the natural overflow is <= HALF a cell
                    //     and the marks cover it (20 kana + 字 with 1 or 3 、 = 0.51:
                    //     refused; 「介します。（各…等は各」 0.49 over 。+・: kept);
                    //   unit X+mark: pulled in iff EITHER the whole unit fits INSIDE
                    //     with at most ONE cell of elective compression, the final mark
                    //     at its natural advance (「…受理する。診察の結果、」 ）+。 =
                    //     1.0; 「…（障害者総合支援法）」とされた。」 （ + the pair), OR
                    //     X fits at NATURAL width and the mark hangs (19 kana + 「字、」
                    //     with no mark). Compression and hang are never combined: 20
                    //     kana + 1-4 、 + 「字、」, （）+「字、」, （）。+「字、」 are all
                    //     refused (the document's 「…がある。な|お、」, 「…などの家|事、」).
                    let s1318_regime = s1318_v2 && s1318_at_default_regime;
                    // S1346: the regime's fit tolerance -- the floor and the cumulative
                    // half-cell increments come from the same pitch through different
                    // roundings and can disagree by a twip.
                    const S1318_TOL_TW: i32 = 2;
                    let s1318_is_mark = |c: char| {
                        matches!(c, '、' | '。' | '，' | '．' | '・' | '：' | '；')
                            || kinsoku::is_yakumono_opening(c)
                            || kinsoku::is_yakumono_closing(c)
                    };
                    // S1318 v6 (2026-09-07): the break-time capacity, from
                    // `_pb_unitcap_gen.py` (52 arms on the faithful slice) and the
                    // document lines it explains:
                    //   - an ADJACENT PAIR of marks (）」 ）、 、「 。（ ...) collapses one
                    //     cell, both members at half, free: 「…（8字）」…字、」 keeps
                    //     字 (22 chars, ） 5.8 」 5.8), the same line without the
                    //     pair wraps it;
                    //   - STANDALONE marks together give strictly LESS than half a
                    //     cell: 4 、 or 4 standalone brackets on the line do not pull
                    //     in a character that needs 0.5 (m1-m4_h1, b2/b4), while
                    //     0.49 (「介します。（各」 after its pair, 06ee35d5's 「・・」
                    //     0.499) and 0.33 (「及び（公財）東京…」 tracked) are kept;
                    //   - the （ of 「…（障害者総合支援法）」とされた。」 is halved at
                    //     RENDER time only: 21 chars - the pair's cell = 20 cells and
                    //     the 。 hangs, so the break needs nothing from it.
                    // Oxi's pre-compression already halves a pair's second member
                    // (0.5); the other half is credited here.
                    let (s1318_pair_credit_tw, s1318_n_solo) = if s1318_regime {
                        let mut prev_mark = false;
                        let mut pairs = 0;
                        let mut solo = 0;
                        let mut run = 0usize;
                        for c in current_line
                            .fragments
                            .iter()
                            .flat_map(|f| f.text.chars())
                            .chain(word.chars())
                        {
                            let m = s1318_is_mark(c);
                            if m {
                                run += 1;
                            } else {
                                if run == 1 {
                                    solo += 1;
                                } else if run >= 2 {
                                    pairs += run - 1;
                                }
                                run = 0;
                            }
                            prev_mark = m;
                        }
                        let _ = prev_mark;
                        if run == 1 {
                            solo += 1;
                        } else if run >= 2 {
                            pairs += run - 1;
                        }
                        // S1346: a LINE-INITIAL opening bracket offers no blank -- Word
                        // draws 「（303以下は…であって行」 with the （ at full width and
                        // wraps the 末 that needed 0.5 (the line's only mark).
                        let first_two: Vec<char> = current_line
                            .fragments
                            .iter()
                            .flat_map(|f| f.text.chars())
                            .chain(word.chars())
                            .take(2)
                            .collect();
                        if first_two.first().map_or(false, |&c| kinsoku::is_yakumono_opening(c))
                            && !first_two.get(1).map_or(false, |&c| s1318_is_mark(c))
                        {
                            if solo > 0 {
                                solo -= 1;
                            }
                        }
                        // S1346 (2026-09-07, `_pb_unitcap_gen.py` 120 arms): a pair's
                        // first member is halved UNCONDITIONALLY (a 19-cell line still
                        // draws its ） at 5.8) -- that is natural width, which Oxi's
                        // pre-compression already carries by halving the second
                        // member -- and everything else (a standalone mark, the pair's
                        // other half) is ONE elective budget of half a cell per line,
                        // inclusive: 0.5 exactly is granted from any single mark
                        // (）（・、」。), 0.46 from one mark of a tracked line, while 0.64
                        // with 2 or 3 standalone marks and 1.0 with 2 are refused. The
                        // v6 "strictly less than 0.5" came from arms whose digit sat
                        // right before the overflowing character (see s1318_prev_latin).
                        (0, solo + pairs)
                    } else {
                        (0, 0)
                    };
                    let s1318_cap_tw: i32 = if s1318_n_solo > 0 { s1318_half_cell_tw } else { 0 };
                    // S1490: k/(k+1) cell minus 0.045 for a normal character; a
                    // unit X+mark needs 0.05 cell more (census: k=1 wraps at 0.457,
                    // k>=2 packs at 0.505).
                    let s1318_one_cell_tw = pt_to_tw(char_width.max(font_size));
                    // `compcap.py` (10.5pt, kinds identical): the capacity depends on
                    // the marks k AND the line's length n -- marks shrink further on a
                    // short line. Measured, cells of overflow a normal character may
                    // pull in (n = characters already on the line):
                    //   n=14: 0.495 0.838 0.886 0.910
                    //   n=26: 0.495 0.724 0.800 0.838
                    //   n=40: 0.452 0.624 0.714 0.762   (12pt n=35: +0.03)
                    // A unit X+mark (n=40): 0.429 0.538 0.614 -> -0.02/-0.09/-0.10.
                    // ：； (JIS 中点類) do not count: technical__978ec9c102290205 wraps
                    // a 0.58-cell overflow with 「：、」 on the line.
                    let s1490_colons = current_line.fragments.iter().flat_map(|f| f.text.chars()).filter(|&c| c == '：' || c == '；').count();
                    let punctuation_index = punctuation_offsets[frag_outer_idx] + char_index;
                    let following_punctuation = &punctuation_chars[punctuation_index + 1..];
                    let next_mark_compressed = yakumono_compressed
                        .get(char_index + 1).copied().unwrap_or_else(||
                            yakumono_pair_enabled && following_punctuation.first()
                                .copied().is_some_and(kinsoku::is_yakumono_closing)
                                && following_punctuation.get(1).copied()
                                    .is_some_and(kinsoku::is_yakumono_trigger));
                    let s1490_cap = |k: usize, unit: bool| -> i32 {
                        LayoutEngine::modern_punctuation_capacity_tw(
                            current_line.fragments.iter().map(|f| f.text.chars().count()).sum(),
                            k.saturating_sub(s1490_colons), s1318_one_cell_tw, unit,
                            ((japanese_language_oikomi && !unit) || (unit
                                && following_punctuation.iter().take_while(|&&nc|
                                    matches!(nc as u32, 0x3001 | 0x3002 | 0xFF0C | 0xFF0E | 0xFF1A | 0xFF1B | 0x30FB)
                                        || kinsoku::is_yakumono_closing(nc)).count() >= 2
                                && style.east_asia_lang.as_deref().is_some_and(|lang|
                                lang.eq_ignore_ascii_case("ja") || lang.to_ascii_lowercase().starts_with("ja-"))))
                                .then_some(s1318_half_cell_tw),
                        )
                    };
                    // the regime measures against the TRUE column edge, not the
                    // S1211C whole-cell floor (see `s1318_floor_slack`)
                    // S1318 v5 (2026-09-06): the capacity IS the S1211C floor -- Word's
                    // 2048-pitch columns hold 20 cells and its 42-cell single-column
                    // lines never reach the 43rd (the 43rd char would need 0.11 cell
                    // against the true edge and Word still wraps it); the justified
                    // line is then stretched to the true edge, which is what made the
                    // PDF's 20-cell lines read as pitch 11.78. A line-final mark hangs
                    // past the floor (p4 「…とされた。」: the 。 box ends at the true
                    // edge only because the text before it ends 5.8 short of the floor).
                    let s1318_avail_tw = available_tw;
                    let s1318_natural_over_tw = current_width_tw + s1317_cw_tw - s1318_avail_tw - s1318_pair_credit_tw;
                    // （：；） are line-final marks too: 0ea3ec86 p11 「…福祉手当（都・市町村：」
                    // holds 21 with the 「：」 at the end where Oxi stopped at 「市町」.
                    // S1346: ・ is a line-final mark too -- 0ea3ec86 「どに入院・入所中の
                    // 児童・生徒のために病院・」 holds 21 with the ・ hanging (the normal
                    // branch refused it and 追い出し took 院 down with it).
                    let s1318_next_mark = following_punctuation.first().map_or(false, |&nc| {
                        matches!(nc, '、' | '。' | '，' | '．' | '：' | '；' | '・') || kinsoku::is_yakumono_closing(nc)
                    });
                    let s1318_ch_mark = matches!(ch, '、' | '。' | '，' | '．' | '：' | '；' | '・')
                        || kinsoku::is_yakumono_closing(ch);
                    let s1318_prev_open = char_index > 0
                        && chars_vec
                            .get(char_index - 1)
                            .map_or(false, |&pc| kinsoku::is_yakumono_opening(pc));
                    // S1346: no elective rescue for a CJK character that directly
                    // follows a Latin-script element (digit, letter, ASCII space):
                    // 「…セ1|字」 「…セ）1|字」 「…）a|字」 「…）1 |字」 all wrap 字 with a
                    // standalone ） on the line (autoSpaceDE off changes nothing),
                    // while 「…1ル字」 「…）１字」 (full-width １) keep it, and a Latin
                    // WORD overflowing after CJK (「…）12」 「…）123」 「…の80」) is
                    // rescued. A pair's natural half still counts (「…（…）」…1字」 keeps 字).
                    let s1318_zw = |c: char| matches!(c, '\u{200B}' | '\u{200C}' | '\u{200D}' | '\u{FEFF}');
                    let s1318_prev_latin = {
                        let pc = chars_vec[..char_index]
                            .iter()
                            .rev()
                            .copied()
                            .chain(current_line.fragments.iter().rev().flat_map(|f| f.text.chars().rev()))
                            .find(|c| !s1318_zw(*c));
                        pc.map_or(false, |c| (c as u32) < 0x2E80 && !c.is_control() && !s1318_is_mark(c))
                    };
                    // S1346: a ZERO WIDTH JOINER after this character glues the next one
                    // to it (0ea3ec86 「…福祉ホ‍ーム」: Word sends ホ‍ー down together; the
                    // faithful slice the same, and without the joiner ホ stays and ー
                    // opens the next line -- ー and small kana are not line-start
                    // prohibited here).
                    let s1318_next_joined: Option<char> = {
                        let mut it = following_punctuation
                            .iter()
                            .copied()
;
                        match it.next() {
                            Some('\u{200D}') => it.find(|c| !s1318_zw(*c)),
                            _ => None,
                        }
                    };
                    let s1318_ch_cjk = (ch as u32) >= 0x2E80;
                    // S1490 v13: a unit whose extra marks push it past the cap is
                    // wrapped even when the character alone fits naturally.
                    let mut s1490_unit_refused = false;
                    let (mut s1318_force_fit, mut s1318_refuse_here) = if !s1318_regime {
                        (false, false)
                    } else if s1318_prev_open {
                        // S1346 (N_/Q_ arms): after an OPENING BRACKET the overflowing
                        // element may spend EVERY blank on the line, the bracket's own
                        // included (each half a cell): 0ea3ec86 p13 「…各島支庁（303」
                        // keeps 303 by halving 、 、 （ (1.5), 「…（3033」 takes 2.0 from
                        // 、、、（, while 、（ alone (1.0) sends 「（303」 down; 「（字」
                        // needing 1.0 is kept with 、、（ and goes down with （ alone.
                        let budget = s1318_n_solo as i32 * s1318_half_cell_tw;
                        let fit = s1318_natural_over_tw <= budget.max(S1318_TOL_TW);
                        (fit, !fit)
                    } else if s1318_ch_mark {
                        // the mark itself: inside at its natural advance (no charSpace
                        // after the line's last character) within the capacity, else
                        // the S601 hang decides (natural fit of the text before it).
                        // S1346 (d_ arms, both section pitches): a LINE-FINAL mark may
                        // spend up to ONE cell of elective compression, each mark's
                        // blank being half a cell (・ = two quarters): 「・・…、」 (21)
                        // keeps the 、 inside at 0.25 + 0.25, 「、、…・」 and 「・・…・」
                        // (21, p30 「…病院・」) give the final ・ 0.74 from the two mid
                        // marks. ・ never hangs whole -- 20 kana + ・ and 19.5 + ・ send
                        // the character before it down too (p4 「…平成27年１|月・」 with
                        // a 、 on the line) -- but its trailing quarter may overhang
                        // (the 21-char lines end 0.26 past the floor).
                        let s1318_cap_final_tw = (s1318_n_solo as i32 * s1318_half_cell_tw).min(s1318_one_cell_tw);
                        let over = current_width_tw + pt_to_tw(font_size) - s1318_avail_tw - s1318_pair_credit_tw;
                        if ch == '・' {
                            let fit = over - s1318_half_cell_tw / 2 <= s1318_cap_final_tw;
                            (fit, !fit)
                        } else {
                            // An unkerned final period/closing pair can spend its
                            // structural half-cell and the closing mark's hanging half.
                            let tail_pair_credit = if s1490_regime && !yakumono_pair_enabled
                                && kinsoku::is_yakumono_closing(ch)
                                && char_index > 0 && matches!(chars_vec[char_index - 1], '\u{3002}' | '\u{FF0E}')
                                && char_index + 1 == chars_vec.len()
                                && fragments[frag_outer_idx + 1..].iter().all(|f| f.0.trim().is_empty())
                                && style.east_asia_lang.as_deref().is_some_and(|lang|
                                    lang.eq_ignore_ascii_case("ja") || lang.to_ascii_lowercase().starts_with("ja-"))
                            { s1318_one_cell_tw } else { 0 };
                            (over <= s1318_cap_final_tw + tail_pair_credit, false)
                        }
                    } else if s1318_next_joined.is_some() {
                        let over = s1318_natural_over_tw + s1318_one_cell_tw;
                        // S1490: the joined mark hangs; X itself must fit the unit cap.
                        let fit = if s1490_regime {
                            s1318_natural_over_tw <= s1490_cap(s1318_n_solo, true).max(S1318_TOL_TW)
                        } else {
                            over <= s1318_half_cell_tw.min(s1318_cap_tw).max(S1318_TOL_TW)
                        };
                        (fit, !fit)
                    } else if s1318_prev_latin && s1318_ch_cjk {
                        // a natural fit needs no rescue; 2 twips absorb the cumulative
                        // rounding of half cells (3194 pitch: 17 kana + ） + ab + 字 =
                        // 4713 against a 4712 floor, Word keeps 字)
                        let fit = s1318_natural_over_tw <= S1318_TOL_TW;
                        (fit, !fit)
                    } else if s1318_next_mark {
                        // v5: X before a mark is a normal character against the floor
                        // (19 kana + 1 、 + 「字、」 keeps 字 with 0.5 of compression and
                        // hangs the 、; 20 kana + 1-4 、 + 「字、」 needs 1.0-1.5 and wraps;
                        // p4 「…（障害者総合支援法）」とされ|た。」 keeps た: 3 brackets
                        // minus the pair's free half = 0.5); the mark then hangs.
                        let over = s1318_natural_over_tw;
                        // S1490 v13: only ONE mark hangs past the character; every
                        // further line-start-prohibited mark of the unit adds its
                        // compressed half cell (technical__978ec9c102290205 p3
                        // 「…ご注意願いま|す。）」: す fits by 0.44 cell, yet Word wraps
                        // the unit す。） because the ） cannot hang behind the 。).
                        let s1490_extra_marks = following_punctuation
                            .iter()
                            .take_while(|&&nc| {
                                matches!(nc, '、' | '。' | '，' | '．' | '：' | '；' | '・') || kinsoku::is_yakumono_closing(nc)
                            })
                            .count()
                            .saturating_sub(1) as i32;
                        // Reserve the complete compressed unit before accepting
                        // its character. Structural pair compression does not make
                        // the remaining marks free. A period-led group may hang
                        // its last half-cell; a bracket-led group remains inside.
                        // Marks earlier on the line supply the final-mark pool,
                        // separately from the ordinary-character pull-in limit.
                        let paired_unit = next_mark_compressed && s1490_extra_marks > 0;
                        let unit_reserve = if paired_unit {
                            let count = s1490_extra_marks + 1;
                            let hang = following_punctuation.first().is_some_and(|c|
                                matches!(*c as u32, 0x3001 | 0x3002 | 0xFF0C | 0xFF0E));
                            count * s1318_half_cell_tw - if hang { s1318_half_cell_tw } else { 0 }
                        } else { s1490_extra_marks * s1318_half_cell_tw };
                        let unit_pool = if paired_unit {
                            (s1318_n_solo as i32 * s1318_half_cell_tw).min(2 * s1318_one_cell_tw)
                        } else { s1490_cap(s1318_n_solo, true) };
                        s1490_unit_refused = s1490_regime && s1490_extra_marks > 0
                            && over + unit_reserve > unit_pool.max(S1318_TOL_TW);
                        let fit = if s1490_regime {
                            !s1490_unit_refused && over <= s1490_cap(s1318_n_solo, true).max(S1318_TOL_TW)
                        } else {
                            over <= s1318_half_cell_tw.min(s1318_cap_tw).max(S1318_TOL_TW)
                        };
                        if s1490_regime && fit && paired_unit {
                            let count = following_punctuation.iter().take_while(|&&nc|
                                matches!(nc as u32, 0x3001 | 0x3002 | 0xFF0C | 0xFF0E | 0xFF1A | 0xFF1B | 0x30FB)
                                    || kinsoku::is_yakumono_closing(nc)).count();
                            accepted_punctuation_unit = Some((lines.len(), punctuation_index + count));
                        }
                        (fit, !fit)
                    } else {
                        let over = s1318_natural_over_tw;
                        let fit = if s1490_regime {
                            over <= s1490_cap(s1318_n_solo, false).max(S1318_TOL_TW)
                        } else {
                            over <= s1318_half_cell_tw.min(s1318_cap_tw).max(S1318_TOL_TW)
                        };
                        (fit, !fit)
                    };
                    // A reserved unit stays atomic while its marks are placed.
                    // A new line invalidates the reservation automatically.
                    if s1490_regime && s1318_ch_mark
                        && accepted_punctuation_unit.is_some_and(|(line, end)|
                            line == lines.len() && punctuation_index <= end) {
                        s1318_force_fit = true;
                        s1318_refuse_here = false;
                    }
                    // A closing mark can consume internal blank space after its
                    // preceding word has already passed the word-fit decision.
                    if !s1318_force_fit && !s1490_unit_refused && s1318_regime
                        && ideographic_closing_spacing && s1318_ch_mark && s1318_prev_latin
                        && s1318_natural_over_tw > 0
                        && {
                            let word_width: f32 = current_line.fragments.iter().rev()
                                .take_while(|f| !f.text.is_empty() && f.text.chars().all(|c|
                                    !c.is_whitespace() && (c.is_ascii() || matches!(c, '¥' | '￥'))))
                                .map(|f| f.width).sum();
                            let expansion = s1318_avail_tw - current_width_tw + pt_to_tw(word_width);
                            let compression = (s1318_natural_over_tw - s1318_n_solo as i32 * s1318_half_cell_tw).max(0);
                            word_width > 0.0 && expansion >= 0 && compression * 2 <= expansion
                        }
                        && s1318_natural_over_tw <= pt_to_tw(ideographic_space_capacity(&current_line))
                    {
                        s1318_force_fit = true;
                        s1318_refuse_here = false;
                        ideographic_closing_lines.insert(lines.len());
                    }
                    if s1318_regime && std::env::var("OXI_DBG1318").is_ok() {
                        eprintln!("[S1318] ch={:?} idx={} cur_tw={} cw_tw={} avail_tw={} over_tw={} cap_tw={} half_tw={} n_elect={} prev_latin={} next_mark={} ch_mark={} force_fit={} refuse={} s475={} r1490={} cap90={}/{} joined={:?} prev_open={} line={:?}",
                            ch, char_index, current_width_tw, s1317_cw_tw, s1318_avail_tw, s1318_natural_over_tw, s1318_cap_tw, s1318_half_cell_tw, s1318_n_solo, s1318_prev_latin, s1318_next_mark, s1318_ch_mark, s1318_force_fit, s1318_refuse_here, s475_break,
                            s1490_regime, s1490_cap(s1318_n_solo, false), s1490_cap(s1318_n_solo, true), s1318_next_joined, s1318_prev_open,
                            current_line.fragments.iter().map(|f| f.text.as_str()).collect::<String>());
                    }
                    let s1318_force_wrap = (s1318_regime && s1318_next_joined.is_some() && s1318_refuse_here) || s1490_unit_refused;
                    let s475_break = s475_break && !s1318_refuse_here;
                    let overflow_tw = if s475_break {
                        // S595 (2026-06-17): for s572 (jc=left legacy no-type oikomi),
                        // Word's per-line oikomi is a SMALL budget, not the per-約物 cap.
                        // DERIVED from the Word PDF (_ikuji_oikomi_derive.py, 176 oikomi
                        // lines): 90% overflow em-natural by ≤2.1pt (median 1.5); a char
                        // overflowing by a full em (para 152 «る» ~11pt) is WRAPPED
                        // (oidashi). The s476 capacity (6.0pt × n約物 ≈ 24pt+) over-credits
                        // → fits «る。» past the margin. Replace it with the natural
                        // (render-width) break + a small oikomi tolerance: keep the char
                        // only if it overflows by ≤ TOL (light 約物 compression absorbs it).
                        // Opt-out OXI_S595_DISABLE; tune OXI_S595_TOL.
                        if s572_legacy_notype_oikomi && std::env::var("OXI_S595_DISABLE").is_err() {
                            let tol = std::env::var("OXI_S595_TOL")
                                .ok()
                                .and_then(|v| v.parse::<f32>().ok())
                                .unwrap_or(4.0);
                            // S1487 (2026-09-19, default ON, opt-out OXI_S1487_DISABLE):
                            // the tolerance is what the line's own marks can absorb
                            // (the S1462 cap, marks x 0.5em) -- a mark-free line gets
                            // none. ikujidetail p1 「３ 配偶者が…」 line 1 (no mark):
                            // Word wraps at 42 chars, Oxi kept a 43rd 1.6pt past the
                            // margin; its lines 2-3 carry 、 and squeeze 42 in.
                            let tol = if std::env::var_os("OXI_S1487_DISABLE").is_none() {
                                let marks = current_line
                                    .fragments
                                    .iter()
                                    .flat_map(|f| f.text.chars())
                                    .filter(|&c| {
                                        matches!(c, '、' | '。' | '，' | '．' | '・' | '：' | '；')
                                            || kinsoku::is_yakumono_opening(c)
                                            || kinsoku::is_yakumono_closing(c)
                                    })
                                    .count() as f32;
                                tol.min(marks * font_size * 0.5)
                            } else {
                                tol
                            };
                            current_width_tw + pt_to_tw(char_width) - available_tw - pt_to_tw(tol)
                        } else if std::env::var("OXI_S601_DISABLE").is_err()
                            // S1429 (2026-09-16, default ON, opt-out OXI_S1429_DISABLE):
                            // no hang inside a table cell (policies__07543a6b p28
                            // 「服用した / 後、横紋筋」; see IN_TABLE_LAYOUT).
                            && !(IN_TABLE_LAYOUT.with(|c| c.get()) > 0
                                && std::env::var_os("OXI_S1429_DISABLE").is_none())
                            // S1334 (2026-09-06, default ON, opt-out OXI_S1334_DISABLE): the
                            // hang is granted at compat <= 14 whatever the alignment, and at
                            // compat 15 ONLY to a justified paragraph. DERIVED
                            // (_pb_hangpunct_gen.py, 36 arms, Word PDF): 40x国+。 at 10.5pt on
                            // a 425.2pt line -- jc=both hangs at c11 and c15, jc=left hangs
                            // at c11 (the 。 placed AT the margin, 510.0) and WRAPS 「国。」
                            // at c15; numPr / exact / doNotCompress / noPunctuationKerning /
                            // balance / useFELayout / doNotWrapTextWithPunct /
                            // doNotUseEastAsianBreakRules / useAltKinsoku / OpenType change
                            // nothing; the faithful slice of 08709ff2's numbered row wraps
                            // 「す。」 at c15 and hangs at c11. Unset compat = legacy.
                            && (self.compat_mode <= 14
                                || !self.compat_mode_explicit
                                || is_justified
                                || std::env::var("OXI_S1334_DISABLE").is_ok())
                            // S1346: under the regime a refused ・ is 追い出し, not hung
                            && !(s1318_regime && ch == '・' && s1318_refuse_here)
                            // S1346: under the regime a closing bracket hangs like 、
                            // (19 kana + 「字」」 keeps 字 with the 」 past the floor)
                            && (matches!(ch, '。' | '、' | '，' | '．' | '・')
                                || (s1318_regime && kinsoku::is_yakumono_closing(ch))
                                // S1220 (2026-08-25, opt-out `OXI_S1220_DISABLE`): the unified
                                // break-billing law from the regime probes
                                // (_pb_wrapbill/_pb_line2bill, 4 conditions):
                                //   the LINE-FINAL mark is free on every line --
                                //   including a closing bracket (kojin p1 keeps
                                //   「含む。）」 with the ） 10.4pt past the margin,
                                //   compat 15)...
                                || (std::env::var("OXI_S1220_DISABLE").is_err()
                                    && self.compat_mode >= 15
                                    && kinsoku::is_yakumono_closing(ch)))
                            // ...EXCEPT at compat 11, where a line-final PAIR cannot
                            // hang (明朝 c11: 。） at 372.95 in a 377.95 budget is
                            // still pushed, kinsoku dragging the 亜 before it;
                            // Ｐ明朝 c11 TRAIL=0 arms read the same).
                            && !(std::env::var("OXI_S1220_DISABLE").is_err()
                                && self.compat_mode < 15
                                && char_index > 0
                                && chars_vec
                                    .get(char_index - 1)
                                    .map_or(false, |&pc| {
                                        kinsoku::is_yakumono_closing(pc)
                                            || kinsoku::is_yakumono_opening(pc)
                                    }))
                            // S1219 (2026-08-25, opt-in `OXI_S1219=1`, HELD): the hang
                            // is granted ONLY on the paragraph's FIRST line. MEASURED
                            // with `tools/metrics/_pb_hangline2_gen.py`: a 36-char
                            // one-line arm ending in 。 outlasts its mark-free control
                            // by the mark's whole advance (both ＭＳ 明朝 and Ｐ明朝),
                            // but the 72-char two-line arm flips 2->3 lines at EXACTLY
                            // the control's r (47.00, both faces) -- a continuation
                            // line gets nothing. tokyoshugyo's two refusals (p20/p21,
                            // both second lines, overflow 1.8pt) are this: Word wraps
                            // them, Oxi hung the 。 and kept 38 characters.
                            // ★HELD OPT-IN (`OXI_S1219=1`): the rule itself is right --
                            // kojin p1's continuation line ends 「…において同」 in
                            // Word's own PDF (n=47), exactly what this gate produces,
                            // and the old behavior packs 「じ。）」 three characters
                            // past Word -- but the golden gate then loses kojin (9
                            // paragraphs, 0.9793) and nedocontract (1): the old
                            // over-pack was COMPENSATING an under-pack elsewhere.
                            // Same shape as S1218. Find the counter-defect first.
                            && (lines.is_empty()
                                || std::env::var("OXI_S1219").ok().as_deref() != Some("1"))
                            && (if s725_final_char {
                                current_width_tw
                            } else {
                                current_capw_tw
                            }) <= available_tw
                            // S1318 v4: in the at-default regime the hang is granted
                            // only when the text before the mark fits at NATURAL
                            // width (compression and hang are never combined).
                        {
                            // S601 (2026-06-18, default ON, opt-out OXI_S601_DISABLE;
                            // char-budget wall): line-end 約物 ぶら下げ
                            // (overflowPunct, default-ON for Japanese). A hangable 約物
                            // at the line end hangs PAST the right margin when the
                            // PRECEDING content fits (current_capw_tw ≤ available) — its
                            // width is NOT counted, so the line is not broken before it.
                            // DERIVED from Word PDF render-truth (_yak_stat.py,
                            // ohnoshugyo): Word fits 42-char ２． lines via mid-約物
                            // compression + the line-end 。 hanging to x521 (11pt past
                            // content-right 510). Oxi's s475 break counted the 。 width →
                            // over-wrapped. Hanging makes overflow_tw ≤ 0 → 約物 placed.
                            current_capw_tw - available_tw
                        } else if s725_final_char {
                            // S725: the paragraph's final char gets a BOUNDED
                            // tolerance instead of the full capacity credit — the
                            // line it lands on is the (never-justified) last line,
                            // so the render can deliver only its unconditional
                            // 約物詰め, not the justify water-fill. Empirical
                            // window (7-doc sweep incl 3a4f/d77a/nedo/ohnoshugyo
                            // + the controlled twin): flat-pt [8.0, 8.46); em-
                            // scaled (caps convention, fs/12) [8.0, 9.67) — the
                            // corpus tails Word KEEPS overflow ≤ ~0.67em, the
                            // twin tail Word WRAPS overflows 0.81em. Default 8.5
                            // (12pt reference), override OXI_S725_TOL.
                            let tol = std::env::var("OXI_S725_TOL")
                                .ok()
                                .and_then(|v| v.parse::<f32>().ok())
                                .unwrap_or(8.5)
                                * font_size
                                / 12.0;
                            // S1462 (2026-09-17, default ON, opt-out
                            // OXI_S1462_DISABLE): the S725 tolerance is a stand-in
                            // for the render's unconditional 約物詰め, so it cannot
                            // exceed what this line can actually give back. A line
                            // with NO compressible mark has nothing to squeeze and
                            // gets no tolerance at all. golden ohnoikuji p7
                            // 「子の出生日の翌日…８週間を経過した日」 (38 chars,
                            // ＭＳ 明朝 10.5, avail 396.85, first-line indent 2.65)
                            // has not one 、。・：； or bracket: Word breaks it
                            // 37 + 1, Oxi kept all 38 on a line that renders to
                            // 515.05 against a content right edge of 510.25 -- the
                            // 4.8pt sat inside the flat 7.44pt window. That one
                            // line was the 18pt Oxi lost on page 7, which let the
                            // page's last paragraph keep both of its lines and
                            // shifted every page from 8 to 13.
                            let tol = if std::env::var("OXI_S1462_DISABLE").is_err() {
                                let marks = current_line
                                    .fragments
                                    .iter()
                                    .flat_map(|f| f.text.chars())
                                    .filter(|&c| {
                                        matches!(c, '、' | '。' | '，' | '．' | '・' | '：' | '；')
                                            || kinsoku::is_yakumono_opening(c)
                                            || kinsoku::is_yakumono_closing(c)
                                    })
                                    .count() as f32;
                                tol.min(marks * font_size * 0.5)
                            } else {
                                tol
                            };
                            current_width_tw + s1317_cw_tw - available_tw - pt_to_tw(tol)
                        } else {
                            current_capw_tw + s475_capinc - available_tw
                        }
                    } else {
                        current_width_tw + s1317_cw_tw - available_tw
                    };
                    let overflow_tw = if modern_cjk_line_end && kinsoku::is_hangable_punct(ch) {
                        current_width_tw + s1317_cw_tw - available_tw
                    } else {
                        overflow_tw
                    };
                    // S1318 v4: a pull-in the regime granted is placed whatever the
                    // S475 per-type caps say.
                    let overflow_tw = if natural_punctuation_boundary {
                        current_width_tw + s1317_cw_tw - available_tw
                    } else if s1318_force_fit {
                        overflow_tw.min(-1)
                    } else if s1318_force_wrap {
                        overflow_tw.max(1)
                    } else {
                        overflow_tw
                    };
                    // Vertical font advances already include the substituted
                    // punctuation metrics. The experimental boundary uses those
                    // advances consistently instead of horizontal hang credits.
                    let mut overflow_tw = if vertical_natural_boundary {
                        current_width_tw + s1317_cw_tw - available_tw
                    } else { overflow_tw };
                    // S1585: the CJK-unit arm (the gap term alone).
                    if overflow_tw > 0 && s1585_on && kinsoku::is_cjk(ch) && ch != '\u{3000}' {
                        let line_chars: Vec<char> = current_line.fragments.iter().flat_map(|f| f.text.chars()).collect();
                        let (g0, _) = s1585_counts(&line_chars);
                        let gap_before = line_chars.last().map_or(false, |c| c.is_ascii_alphanumeric());
                        overflow_tw -= pt_to_tw(s1585_gap_part(g0 + gap_before as usize, font_size));
                    }
                    // S1499 (2026-09-20, default ON, opt-out OXI_S1499_DISABLE): a
                    // compat-15 doNotCompress body line set jc=both absorbs a small
                    // overflow through its CJK<->Latin auto-space gaps. MEASURED on
                    // the faithful slice of c5bb00 paragraph 54 (`slice54_probe`
                    // v14-v23, left-vs-justified right-indent thresholds, 10 tw
                    // steps, Meiryo 10.5pt): no gap -> 0; a line-start Latin island
                    // -> 0; 2 gaps -> 30 tw; 4 -> 50; 6 -> 60; a longer Latin run
                    // adds ~10 (2 gaps + 3-5 letters -> 40). Modelled as fs/14 per
                    // gap capped at fs/4 (30 / 52.5 / 52.5), the conservative side
                    // of every arm (21pt reads 100/140/170 for 2/3/4 gaps: the
                    // model under-credits there and Oxi keeps its current break).
                    // The document's line 1 (48 chars, 5 gaps) overflows 10.8 tw
                    // in Oxi and Word keeps it; the pure-CJK arm refuses 22 tw.
                    if overflow_tw > 0
                        && !s475_break
                        && std::env::var_os("OXI_S1499_DISABLE").is_none()
                        && !self.compress_punctuation
                        && is_justified
                        && self.compat_mode >= 15
                        && !vertical
                        && !lines_and_chars
                        && s476_body
                        && IN_TABLE_LAYOUT.with(|c| c.get()) == 0
                    {
                        let line_chars: Vec<char> = current_line
                            .fragments
                            .iter()
                            .flat_map(|f| f.text.chars())
                            .chain(word.chars())
                            .chain(std::iter::once(ch))
                            .collect();
                        let is_lat = |c: char| {
                            (c.is_ascii_alphabetic() && para_style.auto_space_de)
                                || (c.is_ascii_digit() && para_style.auto_space_dn)
                        };
                        let mut gaps = 0usize;
                        let mut island_from_start = line_chars.first().map_or(false, |&c| is_lat(c));
                        for w in line_chars.windows(2) {
                            let (a, b) = (w[0], w[1]);
                            let a_cjk = kinsoku::is_cjk_ideograph_or_kana(a);
                            let b_cjk = kinsoku::is_cjk_ideograph_or_kana(b);
                            if a_cjk && is_lat(b) {
                                gaps += 1;
                                island_from_start = false;
                            } else if is_lat(a) && b_cjk {
                                if !island_from_start {
                                    gaps += 1;
                                }
                                island_from_start = false;
                            } else if !is_lat(b) {
                                island_from_start = false;
                            }
                        }
                        let credit_tw = pt_to_tw(font_size / 14.0)
                            .saturating_mul(gaps as i32)
                            .min(pt_to_tw(font_size / 4.0));
                        if std::env::var_os("OXI_DBG1499").is_some() {
                            eprintln!("[S1499] ch={:?} over_tw={} gaps={} credit_tw={}", ch, overflow_tw, gaps, credit_tw);
                        }
                        overflow_tw -= credit_tw;
                    }
                    // Legacy gap squeeze (see legacy_gap_on): a CJK char within
                    // min(fs/2, floor), a line-start-prohibited char within the floor.
                    if overflow_tw > 0
                        && legacy_gap_on
                        && IN_TABLE_LAYOUT.with(|c| c.get()) == 0
                        && (kinsoku::is_cjk_ideograph_or_kana(ch) || kinsoku::is_line_start_prohibited(ch))
                    {
                        let line_chars: Vec<char> = current_line
                            .fragments
                            .iter()
                            .flat_map(|f| f.text.chars())
                            .chain(word.chars())
                            .chain(std::iter::once(ch))
                            .collect();
                        let floor = legacy_gap_floor(&line_chars, font_size);
                        let credit = if kinsoku::is_line_start_prohibited(ch) {
                            floor
                        } else {
                            floor.min(font_size / 2.0)
                        };
                        overflow_tw -= pt_to_tw(credit);
                    }
                    if overflow_tw > 0 && s476_body && is_justified
                        && !vertical && !lines_and_chars
                        && kinsoku::is_line_start_prohibited(ch)
                        && std::env::var_os("OXI_CJK_AUTOSPACE_COMPRESSION").is_some()
                    {
                        let capacity: f32 = current_line.fragments.iter()
                            .map(|fragment| fragment.auto_space_shrink).sum();
                        let needed = current_width + char_width - available_tw as f32 / 20.0;
                        if needed > 0.0 && needed <= capacity && capacity > 0.0 {
                            let fraction = needed / capacity;
                            for fragment in &mut current_line.fragments {
                                let reduction = fragment.auto_space_shrink * fraction;
                                fragment.width -= reduction;
                                fragment.natural_width -= reduction;
                                fragment.auto_space_shrink -= reduction;
                            }
                            current_width -= needed;
                            current_width_tw = pt_to_tw(current_width);
                            current_capw_tw -= pt_to_tw(needed);
                            overflow_tw = 0;
                        }
                    }
                    if std::env::var("OXI_DBG721").is_ok()
                        && text.contains("三六協定で定める時間数")
                    {
                        eprintln!("[DBG721] ch={:?} idx={} cw_tw={} capw_tw={} avail={} ovf={} s475={} s590={} lrpb={}",
                            ch, char_index, current_width_tw, current_capw_tw, available_tw, overflow_tw,
                            s475_break, s590_legacy_just_cap, para_has_lrpb);
                    }
                    if let Ok(needle) = std::env::var("OXI_DBG721_TEXT") {
                        if !needle.is_empty() && text.contains(&needle) {
                            eprintln!(
                                "[DBG721T] ch={:?} idx={} cw_tw={} capw_tw={} avail={} ovf={}",
                                ch,
                                char_index,
                                current_width_tw,
                                current_capw_tw,
                                available_tw,
                                overflow_tw
                            );
                            eprintln!("[DBG721V] vertical={} char_width={} cw_inc={} cap_inc={} s475={} s476grid={} s568={} s572={} s590={} force_fit={} force_wrap={}",
                                vertical, char_width, s1317_cw_tw, s475_capinc, s475_break,
                                s476_grid, s568_legacy_oikomi, s572_legacy_notype_oikomi,
                                s590_legacy_just_cap, s1318_force_fit, s1318_force_wrap);
                        }
                    }
                    // PER-LINE TOTAL compression budget FALSIFIED (2026-06-23): tested
                    // OXI_LINE_BUDGET (wrap iff nat_overflow > budget). It correctly wrapped
                    // nedo's preamble 甲 (nat_ovf 18.8, Word compresses only 15.36 & wraps)
                    // but (a) was pagination-NEUTRAL for nedo (the over-fit chars only
                    // redistribute WITHIN paras, no line-count change) and (b) no fixed
                    // budget discriminates: Word's per-line total compression reaches 26pt
                    // (period-rich lines, each 。=6.0 half-em "free") yet wraps W1 at 15.36.
                    // The discriminator is per-約物-TYPE cost (periods free→6.0, closings
                    // light~0.84, openings demand-gated/"expensive") NOT a line total — the
                    // documented "guessing exhausted" wall needing statistical derivation.
                    // See [[char_budget_wall]].
                    // HALF-EM 二分 oikomi cap ATTEMPTED + REVERTED (2026-06-22, OXI_HALFEM): the
                    // best-fit "oikomi iff boundary natural-overflow ≤ half-em" matched the 2
                    // decisive fail lines (nedo 子 3.4<6 OIKOMI, tks 務 5.4>5.25 OIDASHI) and
                    // held on 72-86% of the tokyoshugyo dataset, BUT a hard cap OVER-CORRECTS:
                    // nedocontract 0.9979→0.9522 (+1→+23 — Word oikomi's MANY 12pt lines with
                    // overflow > half-em, the 28% the dataset flagged), 3a4f/model 1.0→0.9994.
                    // So half-em is a strong TENDENCY, not the exact rule; the 28% over-half-em
                    // oikomi (multi-約物 lines / 約物 boundary units) carry an additional factor.
                    // See [[char_budget_wall]] / tokyoshugyo memory. Reverted (env-gate removed).
                    // 82de3fa REVERTED 2026-05-03 (independently confirmed by
                    // Session 52 + Session 51 oxi-3 branch). The trailing-U+3000
                    // immune-from-wrap rule (originally added for ed025 p.1 +0.042)
                    // caused d77a p.10 -0.054, p.9 -0.037, p.8 -0.008 (net -0.099)
                    // because d77a has a paragraph with 142×U+3000 decorative run
                    // (verified by OOXML walk) that got marked immune. Mid-text
                    // U+3000s (used for indentation) were also affected — both
                    // per-fragment and per-paragraph trailing-run-length
                    // threshold gates failed to discriminate d77a's mid-text
                    // U+3000s from ed025's true trailing run, because the
                    // immune flag propagates to ALL U+3000s in a paragraph,
                    // not just the trailing run. Word's actual gate is likely
                    // line-fill aware (only elide when the line is already
                    // near-full).
                    // Net trade: +0.099 d77a recovery / -0.042 ed025 loss
                    // = +0.057 net on bottom-bucket. d77a min p.7=0.6268
                    // unchanged so bottom-5 floor is preserved (3.2377 →
                    // 3.2646, +0.0269 Path A strict positive). ed025 stays
                    // rank 18 (out of bottom-5).
                    // Future: re-attempt with line-fill-aware gate (e.g. only
                    // immune when the line is already at >95% of available_tw).
                    //
                    // R7.62 (Day 36 part 9, 2026-05-14): re-attempt the trailing-U+3000
                    // immunity rule with TWO conditions tightened to avoid the d77a
                    // regression of the previous 82de3fa attempt:
                    //   1. ch == U+3000 AND
                    //   2. ALL remaining chars (char_index+1..end) are also U+3000
                    //      (true trailing-U+3000 run — distinguishes ed025c wi=10's
                    //      5 trailing U+3000 from d77a's 142-char mid-text U+3000
                    //      decorative run where non-U+3000 chars follow) AND
                    //   3. current line is already at ≥95% of available_tw (near-full
                    //      — Word collapses trailing U+3000 only when the line has
                    //      legitimately filled with content first; this excludes
                    //      degenerate "empty + trailing" lines).
                    // ed025c wi=10: "32×U+3000 + 法人番号： + 5×U+3000". Char 38 (1st
                    // trailing U+3000 that overflows) has remaining 4 chars all
                    // U+3000 + line at 99.1% full → immune. Resolves the +16pt
                    // drift jump that cascades 149 paras +1, 69 paras +2/+3.
                    let trailing_u3000 = ch == '\u{3000}'
                        && chars_vec
                            .iter()
                            .skip(char_index + 1)
                            .all(|&c| c == '\u{3000}');
                    let line_near_full =
                        available_tw > 0 && (current_width_tw * 100) >= (available_tw * 95);
                    // With a nonpositive measure, leading fullwidth spaces can
                    // overhang before the first visible glyph. A negative measure
                    // limits each such row to six spaces; zero permits the run.
                    let degenerate_space_overhang = std::env::var("OXI_DEGENERATE_CJK_SPACES").is_ok()
                        && available_tw <= 0
                        && ch == '\u{3000}'
                        && current_line.fragments.iter().all(|f| f.text.chars().all(|c| c == '\u{3000}'))
                        && (available_tw == 0 || current_line.fragments.iter().map(|f| f.text.chars().count()).sum::<usize>() < 6);
                    // S1436 (2026-09-16, default ON, opt-out OXI_S1436_DISABLE): an
                    // ideographic space never starts a new line by itself -- like an
                    // ASCII space it hangs past the edge and the next non-space
                    // character wraps. `_pb_wideindent_gen.py` (tests/fixtures/
                    // wideindent, PREFIX=ideo/ascii): 26 leading U+3000 before 8
                    // characters in a 0-width column give Word 10 lines (all 26
                    // spaces on line 1), Oxi 35. reports__28abf02c p2 「　×26 調査
                    // 平成30年７月」 overflowed a page (+1 x14).
                    let s1436_space_hang = ch == '\u{3000}'
                        && std::env::var_os("OXI_S1436_DISABLE").is_none();
                    // S1488 (2026-09-19, default ON, opt-out OXI_S1488_DISABLE): an
                    // ASCII space reaching this loop (after a CJK character) hangs
                    // the same way -- ikujidetail p9 「…はない。」 + two U+0020: Word's
                    // line ends at x=549.6 with the 。 hung and the spaces past it,
                    // Oxi put the spaces on a 5th line.
                    let s1488_ascii_space_hang = ch == ' '
                        && std::env::var_os("OXI_S1488_DISABLE").is_none();
                    let is_immune_space = (trailing_u3000 && line_near_full) || degenerate_space_overhang || s1436_space_hang || s1488_ascii_space_hang;
                    let line_compress_count = current_line
                        .fragments
                        .iter()
                        .flat_map(|f| f.text.chars())
                        .filter(|&c| kinsoku::is_cjk_compressible(c))
                        .count();
                    // EVEN-DISTRIBUTION break ATTEMPTED + FALSIFIED on the gate (2026-06-23,
                    // OXI_EVENDIST): the hypothesis was that the flat-cap s475 break wrongly
                    // wraps nedocontract word_i=333 «イ…規定する子» because the OPENING
                    // bracket 社（ is capped at 3.1 < its needed even-share 3.335, where
                    // Word distributes the line's demand EVENLY across the 約物 (、3.36 +
                    // （3.31 ≈ demand/2). Implemented `fit iff (natural demand)/(n約物) ≤ aki`
                    // and swept OXI_EVENDIST_AKI on the Phase-1 gate: NO clean window exists.
                    // aki≤3.1 → 333 still +1 AND a new line over-fits (-1); aki 3.2-3.3 →
                    // 4 over-fits, +1 remains; aki≥3.33 → 333 fits but 3 lines over-fit
                    // ({-1:3,0:478}=0.9938 < the flat-cap {0:480,1:1}=0.9979). 333's
                    // even-share (≈3.33) is HIGHER than the over-fit lines' → non-monotonic:
                    // any aki that fits 333 over-fits the others. Even-distribution ≡ the
                    // flat-cap at the equivalent cap (the {-1:3} EXACTLY matches flat
                    // open=3.34) — it is the SAME fit-iff decision function. CONFIRMS the
                    // char-budget wall on the gate: the nedo +1×1 residual is a per-line
                    // oikomi/oidashi BADNESS remainder, not a per-約物 cap/distribution
                    // issue (needs Word's per-line wrap penalty, not a cap). Reverted.
                    // See [[char_budget_wall]].
                    // Phase 2 pair-yakumono compression for compressPunctuation docs.
                    // COM-refined 2026-04-17: Word only absorbs overflow when a
                    // PAIR of adjacent yakumono is present that can actually
                    // compress. Previous `count * font_size * 0.5` formula
                    // overestimated available savings, letting Oxi fit 1-2 extra
                    // chars/line on long paragraphs. Restrict absorption to small
                    // overflow (≤ 10tw = 0.5pt) and only when we have evidence of
                    // pair-compressible yakumono on the line.
                    let has_pair = current_line
                        .fragments
                        .iter()
                        .flat_map(|f| f.text.chars())
                        .collect::<Vec<_>>()
                        .windows(2)
                        .any(|w| {
                            kinsoku::is_yakumono_trigger(w[0]) && kinsoku::is_yakumono_trigger(w[1])
                        });
                    // 2026-04-21: allow absorb when line starts with narrow yakumono
                    // (・/、/。/，/．). COM-verified on d77a pi=24-27 — single-yakumono
                    // line-start needs -2.5pt compression to fit +1 char/line. Without
                    // this extension the compression applied above still leaves ~0.1pt
                    // residual tw overflow that breaks the line 1 char early.
                    let has_linestart_narrow_yakumono = current_line
                        .fragments
                        .first()
                        .and_then(|f| f.text.chars().next())
                        .map_or(false, |c| matches!(c, '・' | '、' | '。' | '，' | '．'));
                    // Threshold raised 10→50tw (2026-04-18) per
                    // project_wrap_overflow_analyzer_e3c545.md analysis:
                    // e3c545 idx=29 at +18tw triggers 20.5pt cascade; 4 of its
                    // 18 over-wraps cluster in 10-50tw and are gated by
                    // has_pair so d77a over-wraps (has_pair=false) remain
                    // unaffected.
                    // S472 demand-driven absorb: a line carrying standalone 、，(left
                    // at NATURAL width upstream when S472 is on) can absorb overflow up
                    // to (count of such 、)×(fontSize/3 ≈4pt each at 12pt) by compressing
                    // them — Word's per-line justify-demand compression. On absorb the
                    // 、 fragments already on the line are retroactively shrunk by the
                    // absorbed overflow so the line fits exactly (break count AND render
                    // both match Word). This replaces the over-eager flat pre-compress.
                    // S543 (2026-06-11): demand oikomi for the NON-justified natural
                    // path. Word compresses each compressible yakumono on the line by
                    // a LIGHT -0.75pt (at fs=10.5; scaled fs*0.75/10.5) to fit one
                    // more char when the natural overflow is within that budget;
                    // otherwise oidashi with zero compression (S492's zero-compression
                    // observation was the oidashi branch only). Repro-confirmed
                    // (tools/metrics/repro_s542_width.py, verbatim 7f272a ３． para:
                    // ．（、 all 9.75 mid-line, 45-char L1; short lines stay 10.50 =
                    // demand-gated). Opening brackets DO compress in this tier
                    // S545 (2026-06-11) THE GATE: the demand oikomi is a
                    // compatibilityMode ≤ 14 (Word 2010) layout behavior.
                    // Bidirectionally repro-confirmed: the isolated real ed025c
                    // para (compat 15, refuses) FIRES when flipped to 14; the
                    // synthetic (compat 14, fires) STOPS when flipped to 15.
                    // ABSENT compatSetting = legacy doc = Word lays out ≤14
                    // (d77a/34140b/04b88e/fded6 have none and oikomi in Word),
                    // but parse_compat_mode reports 15 for them → use the
                    // explicit flag. ed025c (explicit 15) is correctly excluded.
                    // Default ON (spec complete); opt-out OXI_S543_DISABLE.
                    let s543_oikomi = natural_break_jc
                        && std::env::var("OXI_S543_DISABLE").is_err()
                        && self.compress_punctuation
                        && (self.compat_mode <= 14 || !self.compat_mode_explicit);
                    // S556 scaffold (opt-IN OXI_S556_JUST15; default OFF =
                    // byte-identical): the c15 justified pack tier, re-applied
                    // for integration debugging. Slack-table rule from
                    // S551-S555 (T={1:6.5,2:4.15,3:3.3,4:2.65,5+:2.15}).
                    // OXI_S556_DEBUG=1 prints each candidate decision.
                    let s556_just15 = is_justified
                        && std::env::var("OXI_S556_JUST15").is_ok()
                        && self.compress_punctuation
                        && self.compat_mode >= 15 && self.compat_mode_explicit
                        && !lines_and_chars
                        // PLAIN pulls only: a line-start-prohibited overflow
                        // char is the KINSOKU path's case (S550 K matrix:
                        // c15 = oidashi, NO compression) — the pack matrices
                        // all pulled plain 国.
                        && !kinsoku::is_line_start_prohibited(ch)
                        && current_line.fragments.iter()
                            .all(|f| f.width >= f.natural_width - 0.001)
                        && {
                            let cs: Vec<char> = current_line.fragments.iter()
                                .flat_map(|f| f.text.chars()).collect();
                            !cs.windows(2).any(|w|
                                kinsoku::is_yakumono_trigger(w[0])
                                    && kinsoku::is_yakumono_trigger(w[1]))
                        };
                    let mut s472_absorb = false;
                    if !natural_punctuation_boundary && !vertical_natural_boundary && (s472_demand || s543_oikomi || s556_just15)
                        && overflow_tw > 0
                        && self.compress_punctuation
                        && (self.compat_mode >= 15 || s543_oikomi)
                    {
                        let nat = font_size;
                        // Cap-aware compressibles for the whole-line FIT budget:
                        // 、,，→fs/3 (8.0pt floor); 。．and CLOSING brackets→fs/2 (6.0pt
                        // floor; opening brackets never compress). Including closing
                        // brackets fixes bracket-heavy lines (d77a p1 「…規約」) that Word
                        // packs by compressing 」 but Oxi's 、-only budget under-packed.
                        let mut comps: Vec<(usize, f32)> = Vec::new(); // (fi, removable)
                        for (i, f) in current_line.fragments.iter().enumerate() {
                            if f.text.chars().count() != 1 {
                                continue;
                            }
                            let c = f.text.chars().next().unwrap_or(' ');
                            if s556_just15 && !s543_oikomi && !s472_demand {
                                // indices only; rem=0 keeps legacy budget at 0.
                                if !kinsoku::is_s473_compressible(c) {
                                    continue;
                                }
                                comps.push((i, 0.0));
                            } else if s543_oikomi && !s472_demand {
                                // S543 light tier compressibles (、，。．+ opening AND
                                // closing brackets; ！？ never compress).
                                // S546 (2026-06-12): each punct can compress down to its
                                // HALVING floor (fs/2) — the margin-sweep fit boundaries
                                // (_s546e/_s546f: single 、 fits overflow 5.10 not 5.35;
                                // 3-punct line frees −1.5/−1.5/−2.25 at need 5.05)
                                // falsified the S543b flat 0.75 cap, which was a painted
                                // (1px) artifact of small demands. The LINE-TOTAL budget
                                // is fs/2 (see below). Pre-S546 (OXI_S546_DISABLE):
                                // flat 0.75/punct.
                                if !kinsoku::is_s473_compressible(c) {
                                    continue;
                                }
                                let cap = if crate::font::s546_exact_halfwidth() {
                                    font_size / 2.0
                                } else {
                                    0.75
                                };
                                let floor = font_size - cap;
                                let rem = (f.width - floor).max(0.0);
                                if rem > 0.001 {
                                    comps.push((i, rem));
                                }
                            } else if s473_locomp {
                                // S473: break budget = Σ cap over ALL compressibles
                                // (、。，．+ closing AND opening brackets — break-flip
                                // showed opening （ compresses at break too), cap =
                                // s473_cap (≈3.25pt = fs×0.27). NO 0.95 exclusion:
                                // removable = how much THIS fragment can still lose down
                                // to its floor (font_size − cap); = cap when at natural,
                                // tapering for already-compressed fragments. This is the
                                // remaining-capacity model that fixes d77a p1 under-pack
                                // (37→38) and b837 p5 (40→39 rows) without over-packing
                                // p9 (39 needs 3.6pt/、 > cap → still wraps).
                                if !kinsoku::is_s473_compressible(c) {
                                    continue;
                                }
                                let cap_pt = if s473_asym {
                                    match c {
                                        '、' | '，' => s473_cc,
                                        '。' | '．' => s473_cp,
                                        _ if kinsoku::is_yakumono_opening(c) => s473_cop,
                                        _ => s473_ccl, // closing brackets
                                    }
                                } else {
                                    s473_cap
                                };
                                let cap = cap_pt * (font_size / 12.0);
                                let floor = font_size - cap;
                                let rem = (f.width - floor).max(0.0);
                                if rem > 0.001 {
                                    comps.push((i, rem));
                                }
                            } else {
                                if f.width < nat * 0.95 {
                                    continue;
                                }
                                let cap = match c {
                                    '、' | '，' => font_size / 3.0,
                                    '。' | '．' => font_size / 2.0,
                                    '」' | '』' | '】' | '〕' | '》' | '〉' | '｝' | '］'
                                    | '）' => font_size / 2.0,
                                    _ => continue,
                                };
                                comps.push((i, cap));
                            }
                        }
                        // S556: justified-c15 slack-table pack + quanta distribution.
                        if s556_just15 && !s543_oikomi && !s472_demand && !comps.is_empty() {
                            let n = comps.len();
                            let t_n = match n {
                                1 => 6.5f32,
                                2 => 4.15,
                                3 => 3.3,
                                4 => 2.65,
                                _ => 2.15,
                            };
                            let need = overflow_tw as f32 / 20.0;
                            let slack = char_width - need;
                            let fire = need > 0.0 && slack >= t_n;
                            if std::env::var("OXI_S556_DEBUG").is_ok() {
                                // S557 (2026-06-13): the width components expose that
                                // `overflow_tw` in the s475_break (justified c15)
                                // regime is a CAPACITY overflow (capw + capinc −
                                // avail), measured against s475's shallow 2.5pt/punct
                                // compression — NOT a natural overflow. wid−capw is
                                // the s475 compression already baked in. The d77a
                                // para9 "counterexample" is a CASCADE artifact (L3
                                // under-packs 39 vs Word 40 → all later windows shift
                                // +1; Word's に…あらか line is a different 38-char
                                // window than Oxi's drifted 拠…あら). NOT a pack-rule
                                // gap. See [[session557_*]].
                                let head: String = current_line
                                    .fragments
                                    .iter()
                                    .flat_map(|f| f.text.chars())
                                    .take(12)
                                    .collect();
                                eprintln!("S556 cand ch={} need={:.2} slack={:.2} n={} t={:.2} fire={} | s475brk={} wid={:.2} capw={:.2} avail={:.2} capinc={:.2} cw={:.2} head={}",
                                    ch, need, slack, n, t_n, fire,
                                    s475_break, current_width_tw as f32/20.0, current_capw_tw as f32/20.0,
                                    available_tw as f32/20.0, s475_capinc as f32/20.0, char_width, head);
                            }
                            if fire {
                                let q = (need / 0.75).round() as i32;
                                if q >= 1 {
                                    let mut order: Vec<usize> =
                                        comps.iter().map(|(fi, _)| *fi).collect();
                                    order.sort_by_key(|fi| {
                                        let c = current_line.fragments[*fi]
                                            .text
                                            .chars()
                                            .next()
                                            .unwrap_or(' ');
                                        let class = if matches!(c, '、' | '，' | '。' | '．') {
                                            0usize
                                        } else {
                                            1
                                        };
                                        (class, *fi)
                                    });
                                    let floor = font_size / 2.0;
                                    let mut remaining = q;
                                    'rr: loop {
                                        let mut placed = false;
                                        for fi in order.iter() {
                                            if remaining == 0 {
                                                break 'rr;
                                            }
                                            if current_line.fragments[*fi].width - 0.75
                                                >= floor - 0.001
                                            {
                                                current_line.fragments[*fi].width -= 0.75;
                                                remaining -= 1;
                                                placed = true;
                                            }
                                        }
                                        if !placed {
                                            break;
                                        }
                                    }
                                    let saved = (q - remaining) as f32 * 0.75;
                                    current_width -= saved;
                                    current_width_tw =
                                        current_width_tw.saturating_sub(pt_to_tw(saved));
                                    current_capw_tw =
                                        current_capw_tw.saturating_sub(pt_to_tw(saved));
                                    s472_absorb = true;
                                }
                            }
                        }
                        // S546 (2026-06-12): for the S543 light tier the fit budget is
                        // LINE-TOTAL fs/2 — one halfwidth char worth — independent of
                        // punct count (margin-sweep boundaries: single 、 [5.10, 5.35],
                        // 4-punct line [5.10, 5.60], both bracketing 5.25 at fs=10.5;
                        // 0.75×count would cap the 4-punct line at 3.0). Capped by Σrem
                        // (puncts already compressed, e.g. S532 pairs, contribute less).
                        let budget_tw =
                            if s543_oikomi && !s472_demand && crate::font::s546_exact_halfwidth() {
                                pt_to_tw((font_size / 2.0).min(comps.iter().map(|(_, c)| *c).sum()))
                            } else {
                                pt_to_tw(comps.iter().map(|(_, c)| *c).sum())
                            };
                        if !comps.is_empty()
                            && overflow_tw <= budget_tw
                            && s543_oikomi
                            && !s472_demand
                        {
                            // S545/S546 Word distribution rule: comma/period class
                            // (、，。．) before brackets, left-to-right within a class.
                            // S546 deep-demand refinement: the freed amount is assigned
                            // in 0.75pt QUANTA, round-robin across the ordered puncts,
                            // n_quanta = round(overflow/0.75) — reproduces BOTH the
                            // S545 min-count×0.75 observations (small demand: need 1.5
                            // /3 puncts → 2 quanta → 2 puncts −0.75, 1 natural) AND the
                            // deep-demand split (_s546d r=1358: need 5.05 → 7 quanta →
                            // 、−2.25 （−1.5 ）−1.5 = the COM-painted advances exactly).
                            // Pre-S546: full-rem greedy (cap 0.75 ⇒ identical behavior).
                            let mut order: Vec<(usize, f32)> = comps.clone();
                            order.sort_by_key(|(fi, _)| {
                                let c = current_line.fragments[*fi]
                                    .text
                                    .chars()
                                    .next()
                                    .unwrap_or(' ');
                                let class = if matches!(c, '、' | '，' | '。' | '．') {
                                    0usize
                                } else {
                                    1
                                };
                                (class, *fi)
                            });
                            let mut saved = 0.0f32;
                            if crate::font::s546_exact_halfwidth() {
                                let quantum = 0.75f32;
                                let mut n_quanta =
                                    ((overflow_tw as f32 / 20.0) / quantum).round() as i32;
                                let mut rem_cap: Vec<f32> = order.iter().map(|(_, r)| *r).collect();
                                'outer: loop {
                                    let mut assigned_any = false;
                                    for (oi, (fi, _)) in order.iter().enumerate() {
                                        if n_quanta <= 0 {
                                            break 'outer;
                                        }
                                        if rem_cap[oi] >= quantum - 0.001 {
                                            current_line.fragments[*fi].width -= quantum;
                                            rem_cap[oi] -= quantum;
                                            saved += quantum;
                                            n_quanta -= 1;
                                            assigned_any = true;
                                        }
                                    }
                                    if !assigned_any {
                                        break;
                                    }
                                }
                            } else {
                                let mut needed = (overflow_tw as f32) / 20.0;
                                for (fi, rem) in &order {
                                    if needed <= 0.001 {
                                        break;
                                    }
                                    current_line.fragments[*fi].width -= *rem;
                                    saved += *rem;
                                    needed -= *rem;
                                }
                            }
                            current_width -= saved;
                            current_width_tw = current_width_tw.saturating_sub(pt_to_tw(saved));
                            s472_absorb = true;
                        } else if !comps.is_empty() && overflow_tw <= budget_tw {
                            // water-fill the overflow across comps (cap-aware), reducing
                            // current_width so accumulation stays coherent.
                            let mut needed = (overflow_tw as f32) / 20.0;
                            let mut active = comps.clone();
                            let mut amt = vec![0.0f32; current_line.fragments.len()];
                            loop {
                                if active.is_empty() || needed <= 0.001 {
                                    break;
                                }
                                let share = needed / active.len() as f32;
                                let capped: Vec<(usize, f32)> = active
                                    .iter()
                                    .cloned()
                                    .filter(|(_, c)| *c <= share)
                                    .collect();
                                if capped.is_empty() {
                                    for (fi, _) in &active {
                                        amt[*fi] = share;
                                    }
                                    break;
                                }
                                for (fi, c) in &capped {
                                    amt[*fi] = *c;
                                    needed -= c;
                                }
                                active.retain(|(_, c)| *c > share);
                            }
                            let mut saved = 0.0f32;
                            for (i, _) in &comps {
                                current_line.fragments[*i].width -= amt[*i];
                                saved += amt[*i];
                            }
                            current_width -= saved;
                            current_width_tw = current_width_tw.saturating_sub(pt_to_tw(saved));
                            s472_absorb = true;
                        }
                    }
                    let absorb = if vertical_natural_boundary {
                        false
                    } else if s472_absorb {
                        true
                    } else if !s474_natural
                        // Preserve the explicit refusal from unit admission.
                        && !s1318_force_wrap
                        && !s475_break
                        && overflow_tw > 0
                        && overflow_tw <= 50
                        && self.compress_punctuation
                        && self.compat_mode >= 15
                        && (has_pair || has_linestart_narrow_yakumono)
                    {
                        true
                    } else {
                        false
                    };
                    let _ = line_compress_count;
                    if absorb {
                        compress_used = true;
                    }
                    if overflow_tw > 0
                        && !absorb
                        && !is_immune_space
                        && !current_line.fragments.is_empty()
                        && !para_all_whitespace
                    {
                        // Word CJK hybrid hang/oikomi rule — COM-confirmed 2026-04-08.
                        // See memory/hangable_oikomi_rule.md.
                        //
                        // Hang ch on current line (burasagari) only if:
                        //   1. ch is a hangable CJK punct (、。）」 etc.), AND
                        //   2. next char (if any) is NOT line-start-prohibited
                        //      (otherwise hanging would push a still-prohibited char to L2 head).
                        //
                        // S228 (2026-05-23) v2: block hang ONLY when the
                        // current line has already absorbed earlier overflow
                        // (compress_used=true) AND the hang would compound
                        // the cheat. c7b923 wi=43 line 3: char ん at +36tw
                        // overflow gets absorbed (has_pair from mid-line 。）),
                        // then 。 at +129tw hangs. Both cheats stack to
                        // produce 46 chars instead of Word's 33-34.
                        // Word likely refuses the second cheat: if a line
                        // already absorbed overflow, hanging a further
                        // sentence-terminator past the right margin is
                        // disallowed.
                        // Gate fires ONLY when:
                        //   - compress_used = true (line already cheated), AND
                        //   - ch is `。` or `．` (sentence terminator), AND
                        //   - last char of last fragment (= paragraph end)
                        // OXI_LEGACY_HANG_NO_S228_GATE=1 disables.
                        let next_ch = chars_vec.get(char_index + 1).copied();
                        let next_is_proh = next_ch.map_or(false, kinsoku::is_line_start_prohibited);
                        let legacy_s228 = std::env::var("OXI_LEGACY_HANG_NO_S228_GATE").is_ok();
                        let is_para_last_char =
                            frag_outer_idx + 1 == n_fragments && char_index + 1 == chars_vec.len();
                        let is_sentence_terminator = matches!(ch, '。' | '．');
                        // S472h: the S228 hang-block fires when a line already "cheated"
                        // (compress_used). Under the S472 demand model, a line legitimately
                        // compresses 、 by a small justify-demand amount (compress_used=true)
                        // yet should STILL let a trailing 。 hang (b837 para13 L4: る fits via
                        // 、-absorb, then 。 must hang to keep 38/line = Word). So exempt the
                        // S472 path from the S228 hang-block.
                        let s228_block_hang = !legacy_s228
                            && compress_used
                            && is_para_last_char
                            && is_sentence_terminator
                            && !s472_demand;
                        // S492: burasagari (ぶら下げ) is NOT cleanly justify-gated.
                        // The synthetic jc=left repro (国、×30 = 36, oidashi) does NOT
                        // hang, but e3c545 (doNotCompress, type=lines) DOES hang on
                        // its jc=left lines — disabling the hang there cascaded its
                        // pagination (0.9997->0.245, the sole S492 Phase-1 regression).
                        // Burasagari is a doc-level HangingPunctuation behaviour, not a
                        // justify effect; leave it ON for the non-justified path. (The
                        // synthetic comma jc=left then over-hangs +2 vs Word, an
                        // accepted edge case — real docs don't carry 50%-punct lines.
                        // Re-deriving the exact jc/HangingPunctuation gate is next-session
                        // work — see docs/spec/cjk_break_refactor_s492.md.)
                        // S506 (2026-06-08, opt-in OXI_S506_OIDASHI scaffold; default OFF =
                        // byte-identical) — the CORRECT gate STRUCTURE (compat≥15 oidashi tied
                        // to the F1 natural-break path), but still cascades pending the
                        // footnote-vs-body grid distinction. compat≥15 (Word 2013+) does
                        // OIDASHI not burasagari at line-end (S506 repro: compat 12/14 HANG,
                        // 15 OIDASHI; b837=15 / e3c545=14). TEST (OXI_S492_JCNATURAL +
                        // OXI_S506_OIDASHI): b837 STILL cascaded 7→8. ROOT: natural_break_jc
                        // fires for b837's BODY (fs12, linesAndChars) too — disabling its grid
                        // count — and scoping it out via OXI_S492_LINESONLY also kills it for
                        // the FOOTNOTE (fs11) which NEEDS oidashi. break_into_lines cannot tell
                        // the off-grid footnote (→ natural+oidashi) from the on-grid body (→
                        // grid count) within one linesAndChars doc. That distinction (does the
                        // para's font align to the docGrid char pitch?) is the IRREDUCIBLE core
                        // of the S492 Step-2 refactor — see docs/spec/cjk_break_refactor_s492.md
                        // §8 and session505_b837_kinsoku_oidashi. compat_mode>=15 correctly
                        // leaves compat-14 (e3c545) hanging.
                        // S506 (2026-06-08, opt-in OXI_S506_OIDASHI scaffold, default OFF =
                        // byte-identical) — compat≥15 (Word 2013+) does OIDASHI not burasagari
                        // at line-end (S506 repro: compat 12/14 HANG, 15 OIDASHI; b837=15 /
                        // e3c545=14). DEFINITIVE CONCLUSION (3 gate conditions tried —
                        // lines_and_chars / !s476_grid / !s476_body — ALL cascade b837 7→8):
                        // the cascade is NOT a gate-scope problem. The b837 footnote MUST grow
                        // 2→3 lines (oidashi, to match Word's char positions), but growing it
                        // overflows Oxi's layout where Word fits the SAME 3-line footnote in 7
                        // pages — b837 carries a SECOND, vertical, compensating error (Oxi ~1
                        // line of vertical space taller than Word; the wrong 2-line hang
                        // footnote was offsetting it). So the kinsoku oidashi and the vertical
                        // over-height MUST land TOGETHER — b837 is multiply-compensated, the
                        // S492 multi-session refactor. compat_mode>=15 correctly leaves
                        // compat-14 (e3c545) hanging. See docs/spec/cjk_break_refactor_s492.md
                        // §8 / session505_b837_kinsoku_oidashi.
                        // S507 (reverted): an oidashi gate for type=lines compat-15 docs
                        // (683f/0e7af/d77a) was a NO-OP — 0 glyphs changed on 683f. Their
                        // S492 §2 over-pack is the S475 CAPACITY break, not burasagari, so it
                        // is addressed by F1 (OXI_S492_JCNATURAL, disables S475) — a
                        // coverage-track fix that §6 found SSIM-swamped by their structural /
                        // weight-AA errors. The burasagari/oidashi gate (S506) is for the
                        // hang-overflow case (b837 footnote), which these docs don't hit.
                        let s506_oidashi = std::env::var("OXI_S506_OIDASHI").is_ok()
                            && !is_justified
                            && self.compat_mode >= 15
                            && !s476_body;
                        // S548 (2026-06-12, default ON, opt-out OXI_S548_DISABLE):
                        // the S506-confirmed rule shipped with its correct scope.
                        // compat≥15 (Word 2013+) does OIDASHI, not burasagari, at
                        // line-end (S506 repro: compat 12/14 HANG, 15 OIDASHI).
                        // EXPLICIT compat only — absent compatSetting = legacy doc
                        // = Word 2010 layout = burasagari stays (d77a/34140b/
                        // 04b88e/fded6; same semantics as the S545 oikomi gate).
                        // BODY paragraphs included (S506's !s476_body excluded
                        // them = no-op for the 3a4f class: its kern=0 注釈 para
                        // 「…就業規則は、」hung the 、 at natural width 5.3pt PAST
                        // the right margin where Word compat-15 pushes は、 down
                        // → every 注釈 para one line short → the 5 delta=-1
                        // boundary paras = the Phase-1 sole FAIL).
                        // linesAndChars (b837) excluded: its footnote oidashi is
                        // coupled to a compensating vertical over-height (S506
                        // definitive conclusion) — needs the S492 §8 refactor.
                        // s476_body: MAIN BODY flow only — applying the oidashi to
                        // CELL/textbox paragraphs regressed ed025c p4 −0.0685 (its
                        // narrow fitText cells re-wrapped); Word's in-cell hang
                        // behavior is unmeasured — body is the COM-pinned scope.
                        let s548_oidashi = std::env::var("OXI_S548_DISABLE").is_err()
                            && !is_justified
                            && s476_body
                            && self.compat_mode >= 15
                            && self.compat_mode_explicit
                            && !lines_and_chars;
                        // S725: a PARAGRAPH-FINAL hangable punct may hang only when
                        // the preceding content fits NATURALLY — the line it closes
                        // is the (never-justified) last line, so a credit-fit
                        // predecessor would render past the margin (the second
                        // hang site; the S601 site above has the same gate).
                        // S725: same bounded tail tolerance as the overflow site
                        // (em-scaled, 12pt reference, default 8.5).
                        let s725_hang_tol = std::env::var("OXI_S725_TOL")
                            .ok()
                            .and_then(|v| v.parse::<f32>().ok())
                            .unwrap_or(8.5)
                            * font_size
                            / 12.0;
                        let vertical_grid_hang = vertical_natural_boundary && quantized_char_grid
                            && crate::font::vertical_char_grid_on()
                            && current_width_tw <= available_tw;
                        let s1429_cell_oidashi = IN_TABLE_LAYOUT.with(|c| c.get()) > 0
                            && std::env::var_os("OXI_S1429_DISABLE").is_none();
                        let can_hang = kinsoku::is_hangable_punct(ch)
                            && !s1429_cell_oidashi
                            && (!vertical_natural_boundary || vertical_grid_hang)
                            && !modern_cjk_line_end
                            && !next_is_proh
                            && !s228_block_hang
                            && !s506_oidashi
                            && !s548_oidashi
                            && !(s725_final_char
                                && current_width_tw > available_tw + pt_to_tw(s725_hang_tol));

                        if can_hang {
                            current_line.fragments.push(LineFragment {
                                auto_space_shrink: 0.0,
                                text: char_to_string(ch),
                                width: char_width,
                                natural_width: char_width + yakumono_saved,
                                style: style.clone(),
                                tab_alignment: None,
                                tab_position: None,
                                field_type: frag_field_type,
                                run_index: frag_run_index,
                                char_offset: char_pos_in_run,
                            });
                            lines.push(std::mem::take(&mut current_line));
                            current_width = 0.0;
                            current_width_tw = 0;
                            current_capw_tw = 0;
                            latin_space_credit_tw = 0;
                    latin_space_credit_remainder = 0.0;
                            right_tab_slack_tw = 0;
                            center_tab_stop_tw = None;
                            compress_used = false;
                            s1488_after_hang = true;
                            continue;
                        }

                        // Oikomi (押し下げ): pop fragments from end of current line until
                        // both conditions are satisfied:
                        //   - first char of next line is NOT line-start-prohibited
                        //   - last char of current line is NOT line-end-prohibited
                        if std::env::var("OXI_DBGWRAP").is_ok() && text.contains("区市は福祉事務所等") {
                            let tail: Vec<String> = current_line.fragments.iter().rev().take(3).map(|f| f.text.clone()).collect();
                            eprintln!("[WRAP-CJK] at ch={:?} tail={:?} nfrag={}", ch, tail, current_line.fragments.len());
                        }
                        let mut popped: Vec<LineFragment> = Vec::new();
                        loop {
                            let last_of_curr = current_line
                                .fragments
                                .last()
                                .and_then(|f| f.text.chars().last());
                            let next_first = if let Some(p) = popped.last() {
                                p.text.chars().next().unwrap_or(ch)
                            } else {
                                ch
                            };
                            let bad = kinsoku::is_line_start_prohibited(next_first)
                                || last_of_curr.map_or(false, kinsoku::is_line_end_prohibited);
                            if !bad {
                                break;
                            }
                            if current_line.fragments.len() <= 1 {
                                if std::env::var("OXI_KINSOKU_EMERGENCY_BREAK").is_ok() {
                                    // When the entire prefix has no legal boundary,
                                    // retain the fitted line for an emergency break.
                                    current_line.fragments.extend(popped.drain(..).rev());
                                }
                                break;
                            }
                            let f = current_line.fragments.pop().unwrap();
                            current_width -= f.width;
                            popped.push(f);
                        }
                        lines.push(std::mem::take(&mut current_line));
                        current_width = 0.0;
                        current_width_tw = 0;
                        current_capw_tw = 0;
                        latin_space_credit_tw = 0;
                    latin_space_credit_remainder = 0.0;
                        right_tab_slack_tw = 0;
                        center_tab_stop_tw = None;
                        compress_used = false;
                        for f in popped.into_iter().rev() {
                            current_width += f.width;
                            current_width_tw += pt_to_tw(f.width);
                            // S475: re-added oikomi frag. Approximate its break capacity
                            // from its first char's natural − max_compress (edge path).
                            if s475_break {
                                let fc = f.text.chars().next().unwrap_or(' ');
                                let fnext = f.text.chars().nth(1);
                                current_capw_tw += pt_to_tw(
                                    f.natural_width
                                        - kinsoku::s475_aki_cap(
                                            kinsoku::s475_max_compress(
                                                fc, fnext, s475_pair, s475_solo, s475_open,
                                                font_size,
                                            ),
                                            f.natural_width,
                                            // The popped fragments come from the run
                                            // being broken, so its metrics are the ones
                                            // in hand.
                                            s1167_em_ref(
                                                &self.registry,
                                                font_size,
                                                &char_metrics,
                                                gdi_map,
                                            ),
                                        ),
                                );
                            } else {
                                current_capw_tw += pt_to_tw(f.width);
                            }
                            current_line.fragments.push(f);
                        }
                    }

                    current_line.fragments.push(LineFragment {
                        auto_space_shrink: 0.0,
                        text: char_to_string(ch),
                        width: char_width,
                        natural_width: char_width + yakumono_saved,
                        style: style.clone(),
                        tab_alignment: None,
                        tab_position: None,
                        field_type: frag_field_type,
                        run_index: frag_run_index,
                        char_offset: char_pos_in_run,
                    });
                    // S1317: the accumulator follows the cumulative rounding for a
                    // grid-pitched character (see the flag at the grid fold).
                    let s1317_acc_tw = if s1317_grid_char {
                        s1317_inc(current_width, char_width)
                    } else {
                        pt_to_tw(char_width)
                    };
                    current_width += char_width;
                    current_width_tw += s1317_acc_tw;
                    current_capw_tw += if s475_break {
                        s475_capinc
                    } else {
                        s1317_acc_tw
                    };
                } else {
                    // Regular word character — accumulate
                    // autoSpaceDE: add 2.5pt gap when transitioning from CJK ideograph/kana to Latin.
                    // COM-confirmed (2026-04-07): Word only adds auto-space between Latin and
                    // CJK ideographs/kana, NOT between Latin and CJK punctuation.
                    // Session 95 (2026-05-18) split DE (alpha) vs DN (digit).
                    let ch_is_alpha = ch.is_ascii_alphabetic();
                    let ch_is_digit = ch.is_ascii_digit();
                    let s95_de_fires = ch_is_alpha && para_style.auto_space_de;
                    let s95_dn_fires = ch_is_digit && para_style.auto_space_dn;
                    if word_style.is_none() && (s95_de_fires || s95_dn_fires) {
                        // S1316: a ruby field on either side suppresses the auto-space.
                        let s1316_adj = std::env::var("OXI_S1316_DISABLE").is_err()
                            && (style.ruby_field
                                || current_line.fragments.last().map_or(false, |f| f.style.ruby_field));
                        let prev_is_cjk_ideo = !s1316_adj && current_line.fragments.last().map_or(false, |f| {
                            f.text
                                .chars()
                                .last()
                                .map_or(false, |c| kinsoku::is_cjk_ideograph_or_kana(c))
                        });
                        if prev_is_cjk_ideo {
                            // S546: gap = fs/4 true-space (old per-fontSize table = paint artifact).
                            let extra = current_line.fragments.last()
                            .map(|last| self.autospace_after_style(last.text.chars().last().unwrap_or(' '), &last.style, para_style))
                            .unwrap_or_else(|| s546_autospace_extra(font_size));
                            if let Some(last) = current_line.fragments.last_mut() {
                                last.width += extra;
                                last.natural_width += extra;
                            if legacy_gap_on || std::env::var_os("OXI_CJK_AUTOSPACE_COMPRESSION").is_some() {
                                last.auto_space_shrink += extra * 0.5;
                            }
                            }
                            // Keep quarter-em gaps fractional until the cumulative twip conversion.
                            let extra_tw = if std::env::var_os("OXI_CJK_AUTOSPACE_CUMULATIVE").is_some() {
                                pt_to_tw(current_width + extra) - pt_to_tw(current_width)
                            } else {
                                pt_to_tw(extra)
                            };
                            current_width += extra;
                            current_width_tw += extra_tw;
                            current_capw_tw += extra_tw; // S475: autoSpace, no punct capacity
                        }
                    }
                    if word_style.is_none() {
                        word_first_width_tw = pt_to_tw(char_width);
                        word_style = Some(style.clone());
                        word_field_type = frag_field_type;
                        word_run_index = frag_run_index;
                        word_char_offset = char_pos_in_run;
                    }
                    if latin_wordwrap && (word.is_empty() || seg_pending) {
                        word_seg_meta.push((word.chars().count(), frag_run_index, char_pos_in_run));
                        seg_pending = false;
                    }
                    word.push(ch);
                    if !ch.is_whitespace() {
                        s1026_nonws_consumed += 1;
                    } // S1026 final-token
                    word_width += char_width;
                    if latin_wordwrap {
                        word_char_ws.push(word_width);
                    } // S1059
                      // S1022 (2026-07-27, opt-in OXI_S1022): credit an NBSP (U+00A0) as a
                      // fully-equal compressible gap (Word compresses the marker NBSP
                      // proportionally-identically to regular spaces, marker-gap ==
                      // reg-min-space across 3.6..11.76). Oxi glued it into the non-
                      // breaking word token with no compression credit. Includes the
                      // line-START marker NBSPs (before any regular space) by using the
                      // monospace check directly + setting c14_space_tw. Paired with the
                      // corrected s1022_badness_wrap (which wraps the lines Word wraps).
                    if ch == '\u{00A0}'
                        && is_justified
                        && c14_active
                        && (char_metrics.char_width_em('i') - char_metrics.char_width_em('M')).abs()
                            < 0.001
                        && std::env::var("OXI_S1022_DISABLE").is_err()
                    {
                        latin_space_credit_tw += pt_to_tw(char_width * 0.5);
                        if c14_space_tw == 0 {
                            c14_space_tw = pt_to_tw(char_width);
                        }
                    }
                    word_natural_width += char_width + yakumono_saved;
                    if s809_hang || cjk_latin_period_hang {
                        // Existing Latin policy also admits comma/closing quotes;
                        // the separately measured mixed-text policy admits period.
                        word_trail_hang_w = if ch == '.'
                            || (s809_hang && ch == ',' && !s1630_no_comma)
                            || (s809_hang && s1262 && matches!(ch, '\u{201D}' | '\u{2019}'))
                        {
                            char_width
                        } else {
                            0.0
                        };
                    }
                    // S745 (2026-07-04, default ON, opt-out OXI_S745_DISABLE):
                    // `<w:wordWrap w:val="0"/>` = Latin text may break at ANY
                    // character (ECMA-376 §17.3.1.40 — "allow line breaking at
                    // the character level"). Record a break OPPORTUNITY after
                    // every Latin char (the latin_wordwrap mechanism), so a
                    // token crossing the line end splits at the last fitting
                    // char instead of wrapping whole (probezwordwrap: Word
                    // char-breaks the LONGWORDS tokens → Oxi's whole-token
                    // wrap left the line short → +1×20). Fitting tokens are
                    // unaffected (opportunities only matter on overflow).
                    if latin_wordwrap
                        && !para_style.word_wrap
                        && std::env::var("OXI_S745_DISABLE").is_err()
                    {
                        word_breaks.push((word.chars().count(), word_width));
                        seg_pending = true;
                    }
                }
                char_pos_in_run += 1; // character index (not byte offset) for JS compatibility
            }
            // Do NOT flush word here — it may continue in the next fragment
        }

        // Flush any remaining word after all fragments. Route through flush_word! so a
        // LATIN-WORDWRAP over-long token at the paragraph end is split too; the default
        // path inside the macro is the same break/push logic as the original inline flush.
        if !word.is_empty() {
            let fallback_style = fragments.last().map(|f| f.1.clone()).unwrap_or_default();
            flush_word!(fallback_style);
        }

        // Flush last line
        if !current_line.fragments.is_empty() {
            lines.push(current_line);
        } else if std::env::var("OXI_TRAILBR_DISABLE").is_err()
            && lines
                .last()
                .map_or(false, |l| l.break_type == LineBreakType::SoftBreak)
        {
            // S684/TRAILBR (2026-06-28, default ON, opt-out OXI_TRAILBR_DISABLE): a TRAILING
            // <w:br/> (soft break at the paragraph end) leaves an EMPTY current_line that the
            // !is_empty() check above drops. Word renders it as an empty line (+1 line height)
            // — the author's idiom for adding space before the next heading. nedo:
            // «…譲り渡されるものとする。<w:br/>» before «（知的財産権放棄の届出）» → Word +18pt,
            // Oxi dropped it → Oxi over-fit 1 line at the p22 bottom → the {−1:3} "cascade".
            // ★This is the REAL root of nedo's 8-session "char-budget wall" — a dropped
            // trailing <w:br/>, not 約物. Push the empty line. GATE: full Phase-1 85/87→86/87
            // (nedo PASS, 0 regress); SSIM byte-identical (0 word_png docs have a trailing
            // <w:br/>; canaries 3a4f/b837/d77a/kojin/683f byte-identical); lib 142/0/6.
            lines.push(current_line);
        }

        // Ensure at least one empty line for empty paragraphs
        if lines.is_empty() {
            lines.push(Line {
            empty_break_style: None,
                seg2_at: None,
                fragments: vec![],
                ..Default::default()
            });
        }

        // Reconcile only the lines admitted by the closing-mark space reserve.
        for (line_index, line) in lines.iter_mut().enumerate() {
            if !ideographic_closing_lines.contains(&line_index) { continue; }
            let capacity = ideographic_space_capacity(line);
            let target = available_width - if line_index == 0 { first_line_indent } else { 0.0 };
            let natural: f32 = line.fragments.iter().map(|f| f.natural_width).sum();
            let needed = (natural - target).max(0.0).min(capacity);
            if needed <= 0.0 || capacity <= 0.0 { continue; }
            let fraction = needed / capacity;
            let mut leading = true;
            for fragment in &mut line.fragments {
                if leading && fragment.text.chars().all(char::is_whitespace) { continue; }
                leading = false;
                if !fragment.text.is_empty() && fragment.text.chars().all(|c| c == '\u{3000}') {
                    let reduction = (fragment.width - fragment.natural_width * 0.25).max(0.0) * fraction;
                    fragment.width -= reduction;
                    fragment.natural_width -= reduction;
                }
            }
        }

        // 2-pass wrap (Stage 1): compute per-line natural_total_width and
        // was_compressed flag. These are consumed by Stage 2+ for context-aware
        // yakumono handling (loose vs tight line).
        for line in &mut lines {
            let nat: f32 = line.fragments.iter().map(|f| f.natural_width).sum();
            let comp: f32 = line.fragments.iter().map(|f| f.width).sum();
            line.natural_total_width = nat;
            line.was_compressed = (nat - comp) > 0.5;
        }

        // 2-pass wrap (Stage 2): demand-scaled compression revert.
        //   - Full revert: natural fits within available → revert all compression
        //     (Word's loose-line rule: no compression when line has slack)
        //   - Partial revert: natural slightly exceeds available → keep just enough
        //     compression to make line fit, scale rest back toward natural
        //     (Word's demand-driven rule: compression amount matches actual overflow)
        //   - No revert: natural greatly exceeds available (demand ≥ total savings) →
        //     keep full compression
        // S532 (2026-06-10): PAIR-compressed yakumono (。」/）」 adjacency) is
        // EXCLUDED from the revert — Word compresses adjacent-pair punctuation
        // UNCONDITIONALLY (minimal repro _s532_pair_repro.py: 。 advance = 6.0pt
        // exactly, in centered, loose-justified AND wrapping-justified lines
        // alike; d77a title/body 。」 gate pixels agree). Only the demand-driven
        // compressions (standalone 、。 ×0.6667, line-start narrow yakumono)
        // revert on loose lines. The pair members are re-identified here by
        // mirroring the break-time pair rule over the line's char sequence
        // (a fragment is typically one CJK char). Fragments the break never
        // compressed have width==natural, so over-marking is a no-op.
        // opt-out OXI_S532_DISABLE.
        // S547: the pair revert-protection only applies where the pair rule
        // itself applies (w:kern docs). Without this, a kern-less doc whose
        // standalone 、 was pre-compressed by OTHER rules and happens to sit
        // before an opener would be wrongly protected from the loose-line
        // revert (Word keeps it natural — kern0 sweep had zero halved pairs).
        let s532_keep_pairs =
            std::env::var("OXI_S532_DISABLE").is_err() && (!s547_kern_gate || para_kern_on);
        for line in &mut lines {
            if !line.was_compressed {
                continue;
            }
            let pair_frag: Vec<bool> = if s532_keep_pairs {
                let line_chars: Vec<(usize, char)> = line
                    .fragments
                    .iter()
                    .enumerate()
                    .flat_map(|(i, f)| f.text.chars().map(move |c| (i, c)))
                    .collect();
                let n = line_chars.len();
                let mut v = vec![false; n];
                for k in 0..n {
                    let c = line_chars[k].1;
                    if kinsoku::is_yakumono_closing(c) {
                        if k + 1 < n && kinsoku::is_yakumono_trigger(line_chars[k + 1].1) {
                            v[k] = true;
                        }
                    } else if kinsoku::is_yakumono_opening(c) {
                        // S1217: mirror of the break-side rule -- an opening bracket
                        // followed by another opening bracket compresses ITSELF.
                        if std::env::var("OXI_S1217_DISABLE").is_err()
                            && k + 1 < n
                            && kinsoku::is_yakumono_opening(line_chars[k + 1].1)
                        {
                            v[k] = true;
                        } else if k > 0
                            && kinsoku::is_yakumono_trigger(line_chars[k - 1].1)
                            && !v[k - 1]
                        {
                            v[k] = true;
                        }
                    }
                }
                let mut mask = vec![false; line.fragments.len()];
                for k in 0..n {
                    let (fi, c) = line_chars[k];
                    let is_opening = matches!(
                        c,
                        '（' | '「' | '『' | '〔' | '【' | '《' | '〈' | '｛' | '［'
                    );
                    // Only the FIRST char of an adjacent pair compresses (S532
                    // measurement); the second keeps natural advance, so only
                    // v[k] members need revert protection.
                    // S1217: an opening bracket that gave up its half em to a
                    // following opening bracket needs the same revert protection as
                    // a closing one.
                    let s1217_next_open = std::env::var("OXI_S1217_DISABLE").is_err()
                        && k + 1 < n
                        && kinsoku::is_yakumono_opening(line_chars[k + 1].1);
                    if v[k] && (!is_opening || s1217_next_open) {
                        mask[fi] = true;
                    }
                }
                mask
            } else {
                vec![false; line.fragments.len()]
            };
            let savings: f32 = line
                .fragments
                .iter()
                .enumerate()
                .filter(|(fi, _)| !pair_frag[*fi])
                .map(|(_, f)| (f.natural_width - f.width).max(0.0))
                .sum();
            if savings <= 0.5 {
                continue;
            }
            let demand = (line.natural_total_width - available_width).max(0.0);
            if demand <= 0.5 {
                // Full revert: loose line, no compression needed
                for (fi, f) in line.fragments.iter_mut().enumerate() {
                    if !pair_frag[fi] {
                        f.width = f.natural_width;
                    }
                }
                line.was_compressed = line
                    .fragments
                    .iter()
                    .enumerate()
                    .any(|(fi, f)| pair_frag[fi] && (f.natural_width - f.width) > 0.5);
            } else if demand < savings {
                // Partial revert: demand-scaled. Release (savings - demand) back to
                // compressed fragments proportionally, matching Word's per-line
                // demand-driven compression on line-start yakumono (d77a pi=24-27
                // COM: ・ compresses -0.5 to -2.5pt based on line overflow demand).
                let keep_ratio = demand / savings;
                for (fi, f) in line.fragments.iter_mut().enumerate() {
                    if pair_frag[fi] {
                        continue;
                    }
                    let f_saving = (f.natural_width - f.width).max(0.0);
                    if f_saving > 0.0 {
                        f.width = f.natural_width - f_saving * keep_ratio;
                    }
                }
            }
        }

        // A line admitted by the legacy auto-space capacity must also paint
        // with those narrower gaps. Apply only the remaining demand AFTER
        // punctuation reconciliation; glyph advances and source/style metadata
        // stay intact. Capacity is recorded from each actual inserted gap,
        // so mixed run sizes do not turn into one paragraph-wide estimate.
        if legacy_gap_on && IN_TABLE_LAYOUT.with(|c| c.get()) == 0 {
            let is_lat = |c: char| {
                (c.is_ascii_alphabetic() && para_style.auto_space_de)
                    || (c.is_ascii_digit() && para_style.auto_space_dn)
            };
            for (line_index, line) in lines.iter_mut().enumerate() {
                let target = available_width
                    - if line_index == 0 { first_line_indent } else { 0.0 };
                let needed = (line.fragments.iter().map(|f| f.width).sum::<f32>() - target).max(0.0);
                if needed <= 0.0 { continue; }
                // A Latin island touching line start has no shrinkable return
                // gap, matching the boundary model used to admit this line.
                let mut initial_island = line.fragments.iter()
                    .flat_map(|f| f.text.chars()).next().is_some_and(is_lat);
                let capacities: Vec<f32> = line.fragments.iter().map(|f| {
                    if f.text.chars().any(|c| !is_lat(c)) { initial_island = false; }
                    if initial_island { 0.0 } else { f.auto_space_shrink.max(0.0) }
                }).collect();
                let capacity: f32 = capacities.iter().sum();
                if capacity <= 0.0 { continue; }
                let fraction = (needed / capacity).min(1.0);
                for (fragment, cap) in line.fragments.iter_mut().zip(capacities) {
                    let reduction = cap * fraction;
                    fragment.width -= reduction;
                    fragment.natural_width -= reduction;
                    fragment.auto_space_shrink -= reduction;
                }
                line.natural_total_width = line.fragments.iter().map(|f| f.natural_width).sum();
            }
        }

        // Post-process: adjust tab fragment widths for Center/Right/Decimal alignment.
        // ECMA-376 §17.3.1.38: Center tabs center the following segment on the tab position,
        // Right tabs right-align, Decimal tabs align at the decimal point.
        // S841 (2026-07-14, opt-out OXI_S841_DISABLE): each tab targets its
        // ABSOLUTE stop. The raw tab widths were computed during break with
        // UNADJUSTED prior widths, so shrinking tab1 by segment/2 shifted
        // every later tab's landing point left by the same amount (hmrc para
        // B: tab2's strip centered at 361.7 instead of the 413.85 stop =
        // Word x330.5 vs Oxi 278.3). Track the cumulative shrink and restore
        // it on each subsequent tab before applying its own alignment.
        let s841_on = std::env::var("OXI_S841_DISABLE").is_err();
        for line in &mut lines {
            let frag_count = line.fragments.len();
            let mut s841_cum_delta: f32 = 0.0;
            let mut i = 0;
            while i < frag_count {
                if let Some(align) = line.fragments[i].tab_alignment {
                    if align == TabStopAlignment::Left {
                        i += 1;
                        continue;
                    }
                    let _tab_pos = line.fragments[i].tab_position.unwrap_or(0.0);
                    // Measure the segment width after this tab until next tab or end of line
                    let mut segment_width: f32 = 0.0;
                    let mut decimal_offset: Option<f32> = None;
                    let mut j = i + 1;
                    while j < frag_count {
                        if line.fragments[j].tab_alignment.is_some() {
                            break;
                        }
                        if align == TabStopAlignment::Decimal && decimal_offset.is_none() {
                            // Find decimal point position within this fragment
                            let mut char_offset: f32 = 0.0;
                            let fs = line.fragments[j].style.font_size.unwrap_or(11.0);
                            let metrics = self.registry.default_metrics();
                            for ch in line.fragments[j].text.chars() {
                                if ch == '.' || ch == ',' {
                                    decimal_offset = Some(segment_width + char_offset);
                                    break;
                                }
                                char_offset +=
                                    self.registry.char_width_pt_with_fallback(ch, fs, &metrics);
                            }
                        }
                        segment_width += line.fragments[j].width;
                        j += 1;
                    }

                    // Calculate the desired tab width so the segment aligns correctly
                    // Current tab width advances cursor to tab_pos. We need to adjust it
                    // so the segment is positioned according to the alignment type.
                    // S841: re-anchor to the absolute stop first (undo the
                    // cumulative left-shift produced by earlier tabs' shrink).
                    let current_tab_width =
                        line.fragments[i].width + if s841_on { s841_cum_delta } else { 0.0 };
                    let adjustment = match align {
                        TabStopAlignment::Center => segment_width / 2.0,
                        TabStopAlignment::Right => segment_width,
                        TabStopAlignment::Decimal => decimal_offset.unwrap_or(segment_width),
                        TabStopAlignment::Left => 0.0,
                    };
                    // New tab width = original width - adjustment (shift left by adjustment)
                    let new_width = (current_tab_width - adjustment).max(0.0);
                    s841_cum_delta += line.fragments[i].width - new_width;
                    line.fragments[i].width = new_width;
                }
                i += 1;
            }
        }

        // S492 (2026-06-03) — paragraph-level DEMAND break optimizer (env OXI_S492_OPT,
        // default OFF = byte-identical). Replaces the char-greedy break for JUSTIFIED
        // linesAndChars paragraphs with a Knuth-Plass DP that minimizes per-line
        // underfull² with free residual compression (= fill each line to ~avail with
        // LIGHT compression, Word's demand behaviour). Validated 72% per-line match vs
        // Word on b837 (vs 58-62% for any per-line greedy; greedy+maxcomp over-packs at
        // 31%). Render unchanged (decides COUNTS only; render water-fill re-justifies).
        // Scope: justified + linesAndChars + all-Normal-break + no tab/field fragments;
        // re-derive scope before extending. See docs/spec/cjk_break_optimizer_design.md.
        let s492_opt = std::env::var("OXI_S492_OPT").is_ok();
        if s492_opt
            && is_justified
            && lines_and_chars
            && lines.len() > 1
            && lines.iter().all(|l| {
                l.break_type == LineBreakType::Normal
                    && l.fragments
                        .iter()
                        .all(|f| f.tab_alignment.is_none() && f.field_type.is_none())
            })
        {
            let flat: Vec<LineFragment> = lines
                .iter()
                .flat_map(|l| l.fragments.iter().cloned())
                .collect();
            let n = flat.len();
            if n > 1 {
                let avail_l0 = (available_width - first_line_indent).max(0.0);
                let avail_cont = available_width;
                let mut pn = vec![0.0f32; n + 1];
                let mut pc = vec![0.0f32; n + 1];
                for k in 0..n {
                    let mc = if flat[k].text.chars().count() == 1 {
                        let fs = self.resolve_font_size(&flat[k].style, para_style);
                        kinsoku::s492_max_compress(flat[k].text.chars().next().unwrap(), fs)
                    } else {
                        0.0
                    };
                    pn[k + 1] = pn[k] + flat[k].natural_width;
                    pc[k + 1] = pc[k] + mc;
                }
                // Cost weights (env-tunable during the canary). Defaults from the
                // Python fit; w_line breaks ties toward packing (the Python's implicit
                // tie order favoured packing; Rust needs it explicit, else lines
                // under-pack once slack hits 0).
                let w_slack: f32 = std::env::var("OXI_S492_WSLACK")
                    .ok()
                    .and_then(|v| v.parse().ok())
                    .unwrap_or(1.0);
                let w_comp: f32 = std::env::var("OXI_S492_WCOMP")
                    .ok()
                    .and_then(|v| v.parse().ok())
                    .unwrap_or(0.0);
                let w_line: f32 = std::env::var("OXI_S492_WLINE")
                    .ok()
                    .and_then(|v| v.parse().ok())
                    .unwrap_or(0.0);
                let inf = f32::INFINITY;
                let mut best = vec![inf; n + 1];
                let mut prev = vec![0usize; n + 1];
                best[0] = 0.0;
                // Overflow tolerance (env-tunable): Word's linesAndChars grid fits a
                // partial trailing cell / hangs a punct, so a line may exceed avail by
                // up to ~half a cell even with little compression. Default 0.6; sweep.
                let tol: f32 = std::env::var("OXI_S492_TOL")
                    .ok()
                    .and_then(|v| v.parse().ok())
                    .unwrap_or(0.6);
                for j in 1..=n {
                    // kinsoku: break after frag j-1 is invalid if the next frag (j) would
                    // start a line with a line-start-prohibited char, or frag j-1 ends
                    // with a line-end-prohibited char.
                    if j < n {
                        if let Some(c0) = flat[j].text.chars().next() {
                            if kinsoku::is_line_start_prohibited(c0) {
                                continue;
                            }
                        }
                    }
                    if let Some(cl) = flat[j - 1].text.chars().last() {
                        if kinsoku::is_line_end_prohibited(cl) {
                            continue;
                        }
                    }
                    for i in 0..j {
                        if !best[i].is_finite() {
                            continue;
                        }
                        let avail = if i == 0 { avail_l0 } else { avail_cont };
                        let natural = pn[j] - pn[i];
                        let comp = pc[j] - pc[i];
                        if natural - comp > avail + tol {
                            continue;
                        } // infeasible even compressed
                        let lc = if j == n {
                            0.0 // last line: ragged-right, free
                        } else {
                            let slack = (avail - natural).max(0.0);
                            let used = (natural - avail).max(0.0);
                            w_slack * slack * slack + w_comp * used * used
                        };
                        let t = best[i] + lc + w_line;
                        if t < best[j] {
                            best[j] = t;
                            prev[j] = i;
                        }
                    }
                }
                if best[n].is_finite() {
                    let mut bounds = vec![n];
                    let mut j = n;
                    while j > 0 {
                        j = prev[j];
                        bounds.push(j);
                    }
                    bounds.reverse();
                    let mut new_lines: Vec<Line> = Vec::with_capacity(bounds.len());
                    let mut flat_iter = flat.into_iter();
                    let mut taken = 0usize;
                    for w in bounds.windows(2) {
                        let count = w[1] - w[0];
                        let mut frags: Vec<LineFragment> = Vec::with_capacity(count);
                        for _ in 0..count {
                            frags.push(flat_iter.next().unwrap());
                        }
                        let _ = taken;
                        taken += count;
                        let nat: f32 = frags.iter().map(|f| f.natural_width).sum();
                        let comp: f32 = frags.iter().map(|f| f.width).sum();
                        new_lines.push(Line {
            break_source: None,
            empty_break_style: None,
                            seg2_at: None,
                            whitespace_paragraph: para_all_whitespace,
                            emergency_word_break: false,
                            fragments: frags,
                            break_type: LineBreakType::Normal,
                            natural_total_width: nat,
                            was_compressed: (nat - comp) > 0.5,
                        });
                    }
                    if !new_lines.is_empty() {
                        lines = new_lines;
                    }
                }
            }
        }

        // OXI_DUMP_LINEW=1: what the breaker decided, for EVERY line it produced.
        // The over-wide CJK lines the census finds are pushed at a dozen different
        // sites, so the dump belongs here, where the paragraph's lines are final,
        // rather than at any one flush -- an instrument that covers three of the
        // document's lines invites reading one line's breaker width against
        // another line's rendered width.
        if std::env::var("OXI_DUMP_LINEW").is_ok() {
            for line in &lines {
                let fw: f32 = line.fragments.iter().map(|f| f.width).sum();
                let nw: f32 = line.fragments.iter().map(|f| f.natural_width).sum();
                let nch: usize = line.fragments.iter().map(|f| f.text.chars().count()).sum();
                if nch == 0 {
                    continue;
                }
                let head: String =
                    line.fragments.iter().flat_map(|f| f.text.chars()).take(14).collect();
                eprintln!(
                    "[LINEW] nch={} frag_sum={:.2} nat_sum={:.2} avail={:.2} over={:.2} head={}",
                    nch, fw, nw, available_tw as f32 / 20.0,
                    fw - available_tw as f32 / 20.0, head
                );
            }
        }

        for line in &mut lines {
            line.whitespace_paragraph = para_all_whitespace;
        }
        lines
    }
}
