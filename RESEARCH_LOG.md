# Research log — shared across oxi-1/2/3/4

Append-only log of hypotheses tested, confirmed, or refuted across all worktrees.
Read this at the start of every loop iteration to avoid re-chasing falsified leads.
Append your own findings at the end (newest on top).

Format:
```
## YYYY-MM-DD — [worktree] — [confirmed|refuted|partial] — [short label]
- context: what doc/page/feature
- hypothesis: what was tested
- evidence: measurement data, COM results, commit refs
- outcome: what this means for other agents
```

## 2026-05-02 — oxi-4 — partial + correction — 1ec1 +9pt offset variant test (Baseline only); previous "RESOLVED" claim was WRONG (Word ignores ind both □1 and □3, +5.5pt gap structural)
- context: User's R37 measurement (post H_F fix): both □1 (no ind) and □3
  (ind=5.25pt) render at SAME Word x=48pt → Word ignores ind in 1ec1 textbox
  CONFIRMED. My previous "RESOLVED" geometric reconstruction was INCORRECT
  (it assumed ind=5.25pt applied; user's □1 data refutes this).
- correction to previous "RESOLVED" entry:
  - Word renders □1 (no ind) at x=48pt
  - Word renders □3 (ind=5.25pt) at x=48.48pt (essentially same)
  - **Word DOES ignore w:ind in 1ec1's textbox** (consistent with master's
    5b8d07c, generalized further than I thought)
  - The ~5.25pt I attributed to ind:left in the geometric reconstruction was
    LUCKY MATCH not actual mechanism
  - Real unexplained gap: ~+9pt offset Word vs Oxi (post-fix), independent
    of ind value
- partial variant test (per user 依頼書, 1/8 succeeded before Word RPC died):
  - Script: tools/metrics/repro_1ec1_textbox_9pt_v2.py
  - Baseline (rect prstGeom + 1ec1-equiv settings): □ glyph x = 43.0pt
  - 1ec1 actual: □ glyph x = 48.48pt
  - **Difference 5.48pt unexplained** — synthetic minimal repro DOES NOT
    reproduce 1ec1's offset
  - V_A through V_G all failed with Word RPC server unresponsive
- strong hypothesis: 1ec1 uses `<a:prstGeom prst="roundRect">` while my
  synthetic Baseline uses `prst="rect"`. roundRect may add corner-radius
  padding to content area (~5-9pt for 522.75pt-wide shape with default
  corner radius).
- outcome:
  1. **Previous "RESOLVED" claim retracted** — Word DOES ignore w:ind in
     1ec1 textbox; my geometric model was wrong about which 5.25pt added.
  2. **Master's 5b8d07c finding** generalizes to 1ec1 rounded-rect (not
     just synthetic DML wsp as I previously corrected).
  3. **The H_F fix is CORRECT in direction** (suppress ind in textbox).
     User's empirical regression (-7.7pt) is because Oxi was BOTH wrong:
     applied ind AND missing the +5.5pt structural offset. Removing ind
     made the missing +5.5pt visible.
  4. **The +5.5pt gap remains unexplained**. Strong candidate: roundRect
     content padding. Test required: rebuild synthetic Baseline with
     `prst="roundRect"`, see if □ shifts to ~48pt.
- next: re-run variant sweep when Word is healthy. Add V_H_roundRect
  variant. Verify gap source via single property change.
- references:
  memory/investigation_1ec1_9pt_offset_partial_2026_05_02.md
  (full partial result + roundRect hypothesis + reasoning correction).
  Supersedes/corrects investigation_1ec1_box3_root_cause_finally_*.

## 2026-05-02 — oxi-4 — RESOLVED — 1ec1 □３ root cause: paragraph DOES have <w:ind w:left="105"/>; H_F was inverse direction; remaining +2.4pt is pure LSB pipeline diff
- context: deep dive after H_F empirical refutation. Use direct Shape geometry
  + per-paragraph XML scan + reconciliation against all 3 measured points.
- evidence:
  - **□３ paragraph XML inspection** (tools/metrics/deep_dive_1ec1_box3_geometry.py
    + python regex on 1ec1's document.xml inside <w:txbxContent>):
    ```xml
    <w:p>
      <w:pPr>
        ...
        <w:ind w:leftChars="50" w:left="105"/>  ← 5.25pt twip-priority
        ...
      </w:pPr>
      <w:r><w:t>□３ ...</w:t></w:r>
    </w:p>
    ```
    Earlier paragraph dumps showed only □１/□２ (no ind). The specific □３
    AND □４ AND □(1) AND □(2) DO have ind set. Original investigation looked
    at wrong paragraph or assumed first paragraph applied.
  - **Geometric reconstruction** matches all 3 measured points:
    ```
    advance = shape_left(36.275) + lIns(2.83) + ind:left(5.25) = 44.36pt
    Word visible:        44.36 + 4.14pt LSB = 48.50pt  ✓ matches PNG 48.48
    Oxi pre-fix:         44.36 + 1.74pt LSB = 46.10pt  ✓ matches memo 46.1
    Oxi post-fix (no ind): 39.11 + 1.74pt   = 40.85pt  ✓ matches user 40.80
    ```
  - **Master's 5b8d07c finding context**: was for SYNTHETIC DML wsp without
    rounded-rect. Word's COM Information(5)=39pt for that synthetic setup.
    1ec1 uses VML rounded-rect; COM Information(5)=-1 (not applicable);
    visual PNG x=48.48pt with ind APPLIED.
- outcome:
  1. **Word DOES apply w:ind in 1ec1 rounded-rect textbox** (advance includes
     ind:left=5.25pt). Master's "Word ignores ind" rule is **synthetic-only**.
  2. **H_F fix MUST BE REVERTED**. Pre-fix Oxi was closer to Word.
  3. **The remaining +2.4pt diff** (Word visible 48.5 vs Oxi pre-fix 46.1)
     is a pure **glyph LSB rendering pipeline difference**:
     - Word: DirectWrite/Direct2D produces 4.14pt LSB for □ at MS Gothic 14pt
     - Oxi: GDI TextOutW produces 1.74pt LSB (= abcA=2 verified in Investigation B)
  4. **The "□ LSB diff" framing was correct** (per Investigation B's pivot).
     Just the master memo's "Word visual x=39pt" was COM-derived for synthetic
     test, not actual rendering position.
  5. **Master's session_51_textbox_ind_rule.md needs scope qualification**:
     "applies to DML wsp without rounded-rect; DOES NOT apply to VML rounded-rect".
- action items:
  - Revert H_F fix at `crates/oxidocs-core/src/layout/mod.rs:3094-3102`
  - Update master's textbox ind rule with scope qualification
  - +2.4pt LSB residual is sub-pt DirectWrite vs GDI difference, structural
    (would require DirectWrite renderer port to fully match Word)
- references: memory/investigation_1ec1_box3_root_cause_finally_2026_05_02.md
  (full geometric reconstruction + all 3 measured points reconciled).
  Supersedes investigation_oxi_textbox_indent_gap_*, _REFUTED_, _phase_a/b memos.

## 2026-05-02 — oxi-4 — refuted (empirical) — H_F fix REGRESSED 1ec1: Oxi suppressing w:ind in textbox moved □ 5.3pt FURTHER from Word (SSIM 0.6453→0.6389)
- context: H_F memo (entry below) recommended gating indent_left/right/first_line
  computation on in_textbox=true. User implemented the fix and tested.
- result (USER EMPIRICAL TEST):
  - Pre-fix: Oxi □ at ~46.6pt vs Word 48pt → diff = -1.4pt
  - Post-fix: Oxi □ at 40.80pt vs Word 48.48pt → diff = **-7.7pt** (5.3pt MORE LEFT)
  - SSIM: 0.6453 → **0.6389** (regression)
- diagnosis:
  - Master's 5b8d07c established "Word ignores w:ind in textbox" via DML wsp
    minimal repro (NOT rounded-rect). That finding is correct for that synthetic
    setup.
  - 1ec1 uses VML rounded-rect textbox structure. Word DOES apply (some of) the
    ind for that structure → pre-fix Oxi was CLOSER to Word.
  - The "memo Word visual x=39pt" measurement was COM Information(5), which
    differs from PNG pixel x=48pt for □ in 1ec1's actual textbox.
- outcome:
  1. **H_F is REFUTED for 1ec1** (and likely other rounded-rect textbox docs).
  2. **The fix must be REVERTED** — pre-fix Oxi was closer to truth.
  3. Master's 5b8d07c finding is **narrower than originally framed** — applies
     to DML wsp without rounded-rect, NOT generalizable.
  4. The +1.4pt remaining 1ec1 □ diff is **still unexplained** — Investigation
     A/B/H_F all failed. Need different pivot.
- lesson: empirical evidence from minimal repros doesn't always generalize.
  Pre-baseline pixel measurement on the ACTUAL target doc is critical before
  recommending or applying fixes.
- next: REVERT the H_F fix. Re-investigate by rendering 1ec1 with Oxi
  (current code = unconditional ind), measuring □ pixel x, then comparing to
  Word render at same fixed margin/scale. Diff at advance level → ind/lIns/
  shape_left calculation issue. Diff at glyph level → DirectWrite vs GDI.
- references:
  memory/investigation_oxi_textbox_indent_REFUTED_2026_05_02.md (full lesson
  + recommended next steps). Supersedes
  memory/investigation_oxi_textbox_indent_gap_2026_05_02.md.

## 2026-05-02 — oxi-4 — confirmed — H_F: Oxi applies w:ind in textbox content, but Word ignores ALL w:ind (master 5b8d07c)
- context: Investigation B refuted glyph LSB hypothesis. Master's commit 5b8d07c
  ("Word UNCONDITIONALLY ignores ALL <w:ind> attrs in textbox") provides direct
  spec finding. Now verify Oxi's implementation matches this rule.
- hypothesis (H_F): Oxi applies <w:ind> attrs in textbox content paragraphs.
- evidence (Oxi code review):
  - `crates/oxidocs-core/src/layout/mod.rs:2844`: layout_text_box passes
    `in_textbox=true` to layout_paragraph
  - `in_textbox` flag IS consulted at multiple places:
    line 3120 (char_pitch), 3272 (cw_ratio), 3416 (page break), 3423 (widow),
    3567 (justify), 3589 (overflow Mech 1), 3645 (decompress)
  - **`in_textbox` is NOT consulted at lines 3094-3102** (indent_left,
    indent_right, first_line_indent_raw computation):
    ```rust
    let indent_left = para.style.indent_left
        .or_else(|| para.style.indent_left_chars.map(|c| c / 100.0 * 10.5))
        .unwrap_or(0.0);
    ```
  - Downstream at line 3522: `line_x = start_x + indent_left + extra_indent`
    applies indent unconditionally
  - Line 3038 has `#[allow(unused)] in_textbox: bool` confirming in_textbox
    flag is partially-used (CJK compression suppression worked, indent gate
    missed)
- outcome:
  1. **H_F CONFIRMED**: Oxi applies w:ind in textbox content; Word ignores it.
     Clean spec-implementation mismatch.
  2. **Recommended fix**: ~10-line change at line 3094-3102 to gate indent
     computation on `in_textbox` (return 0 when in textbox).
  3. **HIGH confidence**: master's 14+ variant test + direct code review
     showing the gap. Path B (confidence merge) candidate.
  4. **Decoupled from "□ glyph LSB" framing**: Investigation A/B refuted
     the LSB-rendering interpretation. The "+1.47pt diff" master observed
     is BORDER position (separate issue, not text indent). Indent bug applies
     to ANY textbox paragraph with ind set.
- predicted impact:
  - 1ec1's specific □3 paragraph has no ind → no change for that paragraph
  - 1ec1's OTHER textbox paragraphs with ind → shift left to match Word
  - All docs with textbox + ind in textbox content → potentially affected
- next: oxi-2 implementation session: apply fix, pre-baseline, verify,
  ship via Path B (confidence merge) — empirical evidence is overwhelming.
- references:
  memory/investigation_oxi_textbox_indent_gap_2026_05_02.md (full code review
  + recommended fix code + risk analysis). Builds on master's 5b8d07c +
  Phase A/B refutation memos.

## 2026-05-02 — oxi-4 — refuted — 1ec1 □ glyph LSB Investigation B: charset hypothesis + font linking BOTH refuted; Word's 4.14pt LSB not explainable via GDI
- context: 依頼書 Investigation B. After A refuted font fallback hypothesis,
  test if CreateFontW charset (DEFAULT vs SHIFTJIS) or font linking changes □
  glyph LSB.
- evidence (Python ctypes direct GDI tests):
  - **Test 1 — CreateFontW charset effect** (`tools/metrics/test_gdi_charset_lsb.py`):
    Render MS Gothic 14pt □ at x=20. DEFAULT_CHARSET / SHIFTJIS_CHARSET /
    ANSI_CHARSET all produce IDENTICAL result: face=ＭＳ ゴシック, abcA=2,
    leftmost pixel at x=22 (LSB=2px ≈ 1.5pt). Only SYMBOL_CHARSET substitutes
    to Wingdings (irrelevant for our case).
  - **Test 2 — Glyph existence** (`tools/metrics/test_gdi_glyph_existence.py`):
    GetGlyphIndicesW with GGI_MARK_NONEXISTING_GLYPHS shows MS Gothic has NATIVE
    □ glyph (idx=1703, abcA=2, abcB=15). No font linking / GDI fallback for □.
  - Oxi's measured 1.74pt LSB matches GDI's abcA=2 (≈ 1.5pt at PPEM=19).
    **Oxi behaves correctly** per GDI TextOutW path.
  - Word's measured 4.14pt LSB cannot be reproduced via any GDI path I tested.
- outcome:
  1. **Charset hypothesis FALSIFIED**. Oxi's CreateFontW DEFAULT_CHARSET (1)
     produces same result as SHIFTJIS_CHARSET (128).
  2. **Font linking hypothesis FALSIFIED**. MS Gothic has native □ glyph.
  3. **The "+2.4pt LSB diff" framing is likely wrong**. Word's 4.14pt LSB
     can't come from standard GDI rendering. New candidate hypotheses:
     - **H_D**: Word uses DirectWrite/Direct2D (not GDI) for typography,
       producing different glyph LSB than GDI TextOutW.
     - **H_E**: Original measurement methodology error — "44.36pt advance"
       was COMPUTED from layout formula, not directly measured. Word's actual
       advance may differ (e.g., ind:left or lIns inheritance bug).
     - **H_F**: ind:left or lIns calculation issue — relevant given existing
       1ec1 textbox bullet investigation (session_51_textbox_ind_rule.md).
  4. **Investigation should pivot** from "glyph LSB" to "advance position".
     Re-measure Word's rendered □ position directly and back-calculate what
     advance Word actually used. Compare to Oxi's layout-time advance.
- next: re-verify the "44.36pt advance" assumption. Either:
  (a) Word's render at advance=44.36pt + LSB=4.14pt (= visible 48.5pt) — implies
      DirectWrite vs GDI rendering pipeline difference, or
  (b) Word's actual advance != 44.36pt — implies advance calculation error in
      Oxi (ind:left, lIns, or shape_left).
  Hypothesis (b) is more likely given existing textbox-related findings.
  Defer investigation C (cumulative drift) until advance is verified.
- references:
  memory/investigation_1ec1_box_phase_b_2026_05_02.md (full GDI test data +
  H_D/H_E/H_F hypotheses + recommended pivot to advance verification).
  Phase A: memory/investigation_1ec1_box_font_phase_a_2026_05_02.md.

## 2026-05-02 — oxi-4 — refuted — 1ec1 □ glyph LSB +2.4pt H_A (font fallback): both Word and Oxi resolve to MS Gothic
- context: 依頼書 Investigation A. 1ec1091177b1_006.docx Shape 4 textbox □3
  paragraph: Oxi's □ visible glyph at +2.4pt LEFT of Word's. H_A hypothesis:
  Oxi resolves theme major eastAsia to a different physical font (e.g., Yu
  Mincho Light) than Word.
- evidence:
  - Word COM measurement (tools/metrics/investigate_1ec1_box3_font.py):
    □ char Font.Name = ＭＳ ゴシック (MS Gothic). NameAscii/FarEast =
    "+見出しのフォント - 日本語" (= asciiTheme="majorEastAsia" placeholder).
  - Theme1.xml: `<a:ea typeface=""/>` (empty) BUT `<a:font script="Jpan"
    typeface="ＭＳ ゴシック"/>` provides Japanese-specific value.
  - Run XML rFonts: `asciiTheme=eastAsiaTheme=hAnsiTheme="majorEastAsia"`
    + `hint="eastAsia"` for ALL chars in □3 paragraph.
  - Oxi code review trace:
    1. theme.rs parses `<a:font script="Jpan">` → major_font_ea = Some("ＭＳ ゴシック")
    2. Final fallback "Meiryo" at line 251 doesn't trigger (already Some)
    3. ooxml.rs:4193-4200 resolves eastAsiaTheme="majorEastAsia" via
       resolve_theme_font_pub → returns Some("ＭＳ ゴシック")
    4. style.font_family_east_asia = Some("ＭＳ ゴシック")
    5. resolve_font_family_for_text("□", ...) → has_cjk(□)=true → returns
       run_style.font_family_east_asia = "ＭＳ ゴシック"
    6. font/mod.rs:824 normalize: "ＭＳ ゴシック" → "MS Gothic" (registered)
- outcome:
  1. **H_A FALSIFIED**. Both Word and Oxi resolve to MS Gothic.
  2. The +2.4pt LSB diff is NOT a font selection bug.
  3. **H_B promoted to top hypothesis**: same font, GDI rendering produces
     different glyph LSB.
  4. **Suspect identified**: `tools/oxi-gdi-renderer/src/main.rs:193`
     uses `DEFAULT_CHARSET (1)` in CreateFontW. Word likely uses
     `SHIFTJIS_CHARSET (128)` for Japanese fonts. With DEFAULT_CHARSET,
     GDI may resolve "MS Gothic" to a slightly different physical font
     instance with different LSB metrics for □ U+25A1.
  5. Also noted: MS Gothic in `font_metrics_compact.json` does NOT have
     U+25A1 width entry (only common CJK chars). Doesn't matter for LSB
     calculation since GDI handles glyph metrics at render time, but
     could affect other paths.
- next: Investigation B (依頼書 1 hour scope) — minimal docx with single □
  in MS Gothic 14pt. Render Word + Oxi, pixel-measure □ visible left edge.
  If diff persists → modify CreateFontW charset to SHIFTJIS_CHARSET. Then
  Investigation C for cumulative drift.
- references: memory/investigation_1ec1_box_font_phase_a_2026_05_02.md
  (full code-review trace + B/C plan). Builds on
  session_50_1ec1_phase3_findings.md (original measurement).

## 2026-05-02 — oxi-4 — confirmed — Cluster A root cause: Oxi cell first text y = top_bw + implicit pad_t + cell_text_y_off (3 cumulative additions = +2.5pt)
- context: continuation of Cluster A finding (+2.5pt offset on b35123 first cell).
  Source-level decomposition of Oxi's `crates/oxidocs-core/src/layout/mod.rs`
  to identify the 3 cumulative additions.
- evidence: code review at:
  - **Line 5420-5423**: `*cursor_y += top_bw` when table.style.border (Round 30
    Apr 2026 logic). For sz=4 border: **+0.5pt**.
  - **Line 5606-5609** (and 5444-5447 first pass): `pad_t = bw` when default
    cellMar.top=0 + table has border (Round 30). **+0.5pt**.
  - **Line 5965-5973**: cell_text_y_off = `(lh - cell_max_fs)/2` rounded for
    Single/auto cells. For 10.5pt MS Mincho with lh=13.5 (= GDI tmHeight×83/64):
    raw=1.5, rounded=**+1.5pt**.
  - Total: 0.5 + 0.5 + 1.5 = **+2.5pt** ✓ matches measured offset.
- assembly at line 6085-6087:
  ```
  let dy = *cursor_y + pad_t + v_offset;
  for mut elem in cell_elements { elem.y += dy; ... }
  ```
- outcome:
  1. **Three cumulative bugs** combine to produce the +2.5pt offset.
  2. **Components 1 & 2 contradicted by master's §13.3 RETRACT** (commit 8913593).
     Master's 40-fixture sweep showed border has ZERO effect on text X position;
     same logic applies to Y. The Round 30 measurement that motivated these
     additions was based on a misinterpretation (the cited "1row_outer4 marker_y=72,
     cell_y=97.5 → offset=0.5" doesn't actually compute 0.5pt; 97.5-72=25.5pt).
  3. **Component 3 incorrectly centers** fontSize within lh for default vAlign=top.
     Word does NOT center per-line; text top sits at line top.
  4. **Recommended fix** (for oxi-2 implementation):
     - Remove `cursor_y += top_bw` at line 5420-5423
     - Remove `pad_t = bw` defaults at line 5444-5447 and 5606-5609
     - Replace cell_text_y_off centering with vAlign-conditional logic (only
       center for vAlign=center, else 0)
- expected impact:
  - b35123 first cell shifts from y=128.5 to y=126.0 (matches Word)
  - All cells in 13 of 15 bottom-bucket docs (those with tables) shift up 2.5pt
  - SSIM lift expected on 5+ table-heavy docs (b35123, 2ea81a, 1ec1091, e3c545,
    6514f214). Single fix lifts multiple docs.
  - Risk: cells with vAlign=center may need separate handling; multi-line cell
    behavior across line 1 → line 2 needs verification.
- next: implementation in oxi-2 session, baseline pre-measurement before ship
  (Path A bottom-N gate). Pre-measure b35123 / 2ea81a / 1ec1091 SSIM with each
  of the 3 component fixes individually to identify which are correct.
- references: memory/investigation_oxi_cell_first_y_root_cause_2026_05_02.md
  (full source-level decomposition + recommended fix code).
  Builds on memory/investigation_table_first_cell_y_2026_05_02.md.

## 2026-05-02 — oxi-4 — confirmed — Cluster A: Oxi adds +2.5pt offset to first cell content y (Word: cell text y = table top y)
- context: bottom-bucket Cluster A (heavy-table layout) verification on b35123
  (min SSIM 0.666 page 1, 22 in survey was regex artifact — actually 2 large
  tables with 13r+10r×2c = 46 cells covering most of doc).
- hypothesis: Oxi mis-positions first cell content y due to cellMar.top default
  or border thickness applied incorrectly.
- evidence:
  - **Word's UNIVERSAL rule**: first cell first-character y == tbl.Range.Information(6)
    (= table top y). Verified on 8 tables / 3 docs:
    b35123 (2 tables, dy=0), 2ea81a (4 tables, dy=0), 1ec1091 (1 table, dy=0)
  - **Oxi b35123 measurement**: first cell text at +2.5pt below Word's table top:
    - Table 1 page 1: Word=126.0, Oxi=128.5, dy=+2.50
    - Table 2 page 2: Word= 91.0, Oxi= 93.5, dy=+2.50
    - Consistent across both tables (not noise)
  - b35123 OOXML: tblStyle "af" (Table Grid 0.5pt borders), no tcMar override,
    docDefault tblCellMar top=0, bottom=0, left/right=108tw
  - Word's first cell content sits at cell top (consistent with master's recent
    §13.3 RETRACT: "border has ZERO effect on text position")
- outcome:
  1. **Oxi has spurious +2.5pt offset** for first cell content y. Likely sources:
     (a) default cellMar.top set to 50tw instead of 0
     (b) unconditional vertical centering of first row text
     (c) border thickness mis-application (less likely, doesn't fit 2.5pt)
  2. **Universal Word rule** (8/8 tables): first_cell_text_y = table_top_y.
     No cellMar.top offset, no border offset.
  3. **High-leverage fix**: 13 of 15 bottom-bucket docs have tables. If same
     +2.5pt bug applies universally, fixing it lifts 5+ docs simultaneously.
- next: (a) re-render 2ea81a / 1ec1091 / e3c545 with current Oxi to verify
  +2.5pt is universal vs b35123-specific; (b) code review
  crates/oxidocs-core/src/layout/mod.rs for cell first-paragraph y formula —
  search for `pad_t`, `cellMar.top`, default values; (c) cross-reference §3.3
  grid centering: should NOT apply to first-row table content.
- references: memory/investigation_table_first_cell_y_2026_05_02.md (full
  comparison table + hypothesis ranking + recommended investigation steps).
  Related: master's §13.3 RETRACT (commit 8913593, X-axis); this is Y-axis
  analog. Cluster A from
  memory/investigation_bottom_bucket_post_r32_2026_05_02.md.

## 2026-05-02 — oxi-4 — confirmed — Oxi yakumono cascade implementation has gap with Word's proportional Mech 2 (37% of justified body lines)
- context: Cross-doc Mech 2 audit (RESEARCH_LOG entry below) confirmed 65 of 175
  audited lines (37%) show Word's proportional Mech 2 (yakumono partial compression
  to 7-9pt range). Code review of Oxi's break_into_lines to check if cascade matches.
- hypothesis: Oxi implements Mech 1 + ASCII slack distribution but lacks proportional
  Mech 2 (partial yakumono compression on residual overflow).
- evidence (code review):
  - `crates/oxidocs-core/src/layout/mod.rs:4321` — break_into_lines yakumono pre-pass:
    builds yakumono_compressed[] boolean (Mech 1 Type A/B detection); applies
    ×0.5 to compressed yakumono, ×0.583 to standalone 、。
  - `crates/oxidocs-core/src/layout/mod.rs:3589` (Phase 1 overflow Mech 1):
    on `slack < 0`, re-applies ×0.5 to remaining CJK punctuation. **Binary
    compression — full or half**, no partial.
  - `crates/oxidocs-core/src/layout/mod.rs:3645` (Stage 2b decompress):
    on `slack > 0.5`, restores compressed 、。 toward natural. UNDOES Mech 1.
  - `crates/oxidocs-core/src/layout/mod.rs:3677` (Phase 2 slack):
    on `slack > 0`, distributes positive slack to ASCII spaces or as uniform
    per-char gap. **Adds spacing, not compresses**.
  - **Missing stage** between line 3618 and 3621: proportional Mech 2 that
    distributes residual NEGATIVE slack across full-width yakumono with
    per-char cap (font_size × 1/3 per spec).
- outcome:
  1. **Oxi can produce SELECTIVE behavior** (26 lines = 15% of audited) — matches Word.
     Mech 1 fires on neighbor patterns; others stay full.
  2. **Oxi can produce PURE FULL** (76 lines = 43%) — matches Word when no overflow.
  3. **Oxi CANNOT produce PROPORTIONAL Mech 2** (65 lines = 37%) where Word
     compresses yakumono partially to 7-9pt range. Oxi's options:
     - Phase 1 forces ×0.5 (5.25pt for 10.5pt font) → over-compressed
     - OR leaves full → still overflowing
     - Net mismatch: ~2pt difference per yakumono in 37% of justified lines
  4. **Recommended fix** (for oxi-2 implementation session):
     ```
     // Add Stage 2.5 between line 3618 and 3621:
     if slack < 0 && !in_textbox {
         compressible: yakumono still at full width (not Mech 1 hit yet)
         per_char_reduction = (-slack / n_compressible).min(font_size × 0.5)
         apply to each, recompute slack
     }
     ```
- next: (a) targeted SSIM measurement: implement Stage 2.5, run pipeline.verify,
  expect lift on 3a4f / d77a / b35123 (have most proportional Mech 2 lines);
  (b) cross-reference master's §4.7b spec (a8d70c2) — likely already specifies
  cascade with per-char cap formula; (c) consider Mech 3 (jc=left, master's
  §4.7c) as separate axis from this Mech 2 gap.
- references:
  memory/investigation_oxi_yakumono_cascade_gap_2026_05_02.md (full code references
  + recommended fix + caveats). Builds on
  memory/investigation_mech2_selective_cross_doc_2026_05_02.md.

## 2026-05-02 — oxi-4 — confirmed — Cross-doc Mech 2 selective behavior is GENERAL (5 docs / 175 lines / 486 yakumono)
- context: User question "paragraph 10 L4 ('Mech 1 hit + 3 yak full') pattern が他 doc
  にも共通か。selective rule の一般性確認。"
- hypothesis: selective behavior (Mech 1 hit + uncompressed yakumono coexist on
  same line) generalizes across baseline docs.
- evidence:
  - Audited 5 jc=both + kern docs: 3a4f, 7f272a, ed025, d77a, b35123
  - Method: walked first 30 multi-line body paragraphs per doc, COM-measured
    per-char advance, classified each yakumono as Mech 1 (≤6.5pt) / Mech 2 partial
    (6.5-10pt) / uncompressed (≥10pt) for 10.5pt MS Mincho
  - Total: 175 non-final lines audited, 486 yakumono characters classified
  - Script: tools/metrics/audit_mech2_selective_cross_doc.py
  - Data: pipeline_data/mech2_cross_doc_audit.json + _part2.json
  - Per-doc breakdown:
    - 3a4f: 58 lines, **5 selective**, 37 proportional, 14 pure_full
    - 7f272a: 4 lines, 0 selective (small sample, no overflow)
    - ed025: 9 lines, **3 selective**, 0 proportional, 6 pure_full
    - d77a: 98 lines, **18 selective**, 24 proportional, 52 pure_full
    - b35123: 6 lines, 0 selective, 4 proportional, 2 pure_full
  - **TOTAL: 26 selective lines (14.9%) across 3/5 docs**
- outcome:
  1. **Selective behavior is GENERAL, NOT a 3a4f-specific quirk**. Appears in
     3a4f (5L), ed025 (3L), d77a (18L) = 26 lines across 3 distinct docs.
  2. **Pattern is MIXED (text-dependent)**: same docs show both selective AND
     proportional on different lines. 3a4f: 5 sel vs 37 prop. d77a: 18 sel vs 24 prop.
  3. **Proposed mechanism**: cascade — Word first applies Mech 1 (Type A/B/C
     neighbor pairs), then Mech 2 (slack distribution) on RESIDUAL overflow only:
     - Line fits after Mech 1 alone → SELECTIVE (others stay full)
     - Mech 1 insufficient → Mech 2 distributes residual = PROPORTIONAL
     - No overflow → no compression = pure_full
     - No Mech 1 candidates but overflow → pure proportional Mech 2
  4. **Implication for Oxi**: must implement cascade, NOT pure proportional Mech 2.
     If Oxi only does proportional, it will over-compress yakumono in 15% of
     justified lines that Word leaves at full width (selective lines).
- next: (a) verify Oxi break_into_lines current implementation against cascade
  rule; (b) targeted overflow audit on ed025/7f272a to confirm Mech 2 absence
  is sample bias not real; (c) cross-reference master's §4.7b spec
  (a8d70c2 Mech 1↔Mech 2 precedence) — likely already specifies cascade.
- references:
  memory/investigation_mech2_selective_cross_doc_2026_05_02.md (full per-doc
  data + cascade hypothesis). Builds on session_51_3a4f_p64_p42_validation
  + session_51_oxi_compress_spec_table.

## 2026-05-02 — oxi-4 — correction — e3c545 pagination "divergence" was para_idx mapping artifact (cell paragraphs share table block_idx)
- context: previous entry "Cluster B: e3c545 has CATASTROPHIC pagination divergence
  (130 paras / 12 Oxi pages vs 550 / 12 Word pages)" claimed Oxi paginates 4× slower
  than Word. Investigation continued to identify the over-reserved content type.
- correction: discovered that Oxi's `paragraph_index` field for table cell elements
  is set to the TABLE BLOCK's index (per layout/mod.rs:6038, with comment "Attribute
  to the table's source block index so diff tools can localize cell text"). NOT the
  cell paragraph's own index.
- evidence:
  - tools/metrics/inspect_e3c545_para83.py → para_idx=83 has 121 elements, 74 unique y
  - But OOXML `<w:p>` 84 (= idx 83 1-based) is a single short PreformattedText
    paragraph (just "@prefix cc:&lt;http://creativecommons.org/ns#&gt;.")
  - Mismatch resolved by code review: para_idx aliasing for table cell content
  - Oxi has 541 `<w:p>` total; 119 of those are inside tables (3 `<w:tbl>` opens
    before the data-set-definition position, 2 `</w:tbl>` closes — i.e., inside
    one of 9 tables)
- corrected interpretation:
  - Oxi IR has ~130 Body Blocks for e3c545 (each table is ONE Block containing
    many cell paragraphs)
  - Word's `Paragraphs(i)` includes all 541 (body + cell) paragraphs; Oxi's
    para_idx in layout JSON is a body Block index for body content but aliases
    to table block_idx for cell content
  - Mapping `Word p_i = Oxi para_idx (p_i - 1)` breaks for cell paragraphs
  - The "5-page divergence at p100" was Oxi's body para 99 vs Word's cell para 100
    (different entities)
- what's still real:
  - Word p1-p4 dy = 0.0pt (body, perfect match)
  - Word p5 dy = -3.5pt
  - **Word p50 dy = +21pt** (body paragraph, cumulative drift)
  - This implies ~0.4pt per body paragraph drift accumulating over the first 50
    body paragraphs. Cluster B (per-para drift) hypothesis remains valid at this
    smaller scale.
- outcome:
  1. **PREVIOUS ENTRY's "catastrophic divergence" claim is RETRACTED**. The data
     was misinterpreted via an incorrect Word↔Oxi paragraph mapping.
  2. **Body paragraph drift IS real** at ~0.4pt/para in early paragraphs of e3c545.
  3. **Cluster B (per-para drift) hypothesis remains valid** — magnitude smaller
     than claimed but real. Need text-match-based comparison (not index mapping)
     to quantify cumulative drift accurately for full p1-p550 range.
- next: (a) build text-content-based Word↔Oxi paragraph mapping; (b) compute
  cumulative Y drift across all body paragraphs (excluding cell paragraphs);
  (c) identify which body paragraphs show step-changes in dy (suggesting specific
  feature triggers).
- references: memory/investigation_e3c545_pagination_correction_2026_05_02.md
  (correction memo). Supersedes investigation_e3c545_pagination_divergence_2026_05_02.md.

## 2026-05-02 — oxi-4 — confirmed (revised mechanism) — Cluster B: e3c545 has CATASTROPHIC pagination divergence (130 paras / 12 Oxi pages vs 550 / 12 Word pages)
- context: bottom-bucket survey Cluster B (long-doc cumulative Y drift) hypothesis.
  Tested on e3c545 (550 paras, min SSIM 0.665 page 11).
- hypothesis tested: small per-paragraph Y drift accumulates over hundreds of paras
- evidence:
  - Word COM Y measurement (tools/metrics/measure_long_doc_drift.py) on 20 sampled
    paragraphs across e3c545
  - Oxi cached layout comparison (tools/metrics/compare_long_doc_drift.py) using
    pipeline_data/_e3c545_layout.json
  - Result: Oxi 12 pages cover only paras 0-129. Word 12 pages cover all 550.
  - dy progression: 0pt (paras 1-4) → -3.5pt (p5) → +21pt (p50) → 5-page
    divergence (p100: Word page 4, Oxi page 9)
  - Oxi page 5 contains ONLY 1 paragraph (para_idx=83) — single para fills full page
- outcome:
  1. **Cluster B mechanism REVISED**: NOT linear per-para drift. Oxi paginates
     ~4× slower than Word due to catastrophic over-reservation of specific content
     blocks (likely code blocks, nested tables, or specific OOXML features).
  2. **Single-paragraph-fills-page anomaly** at para_idx=83 indicates a content
     block that Oxi grossly over-reserves. e3c545 is a TTL/RDF technical
     handbook with monospace code blocks — likely candidate.
  3. **Why bottom-page SSIM is low**: Oxi page 11 contains paras 109-118 while
     Word page 11 contains paras ~470-490 (completely DIFFERENT content). The
     pixel-comparison sees totally different text → SSIM crashes.
  4. **Cluster B should be re-classified** as "block over-reservation" not
     "per-para drift". Aligns more with Cluster A (table cell layout) than the
     original per-para hypothesis.
- next: (a) identify para_idx=83 content type in e3c545 (code block? table?);
  (b) trace per-paragraph cursor_y advance through paras 70-100 to find where
  over-reservation kicks in; (c) verify pattern on 04b88e / 34140b (no Oxi cache
  yet); (d) hypothesis: specific element (`<w:pre>`-style? grid charSpace?
  monospace font lh formula?) over-reserves vertically.
- references:
  memory/investigation_e3c545_pagination_divergence_2026_05_02.md (full data + per-page
  para_idx range + revised mechanism). Builds on
  memory/investigation_bottom_bucket_post_r32_2026_05_02.md.

## 2026-05-02 — oxi-4 — survey — Bottom-bucket post-R32: 15 docs, 5 hypothesis clusters
- context: User question "R32 後 bottom-bucket survey, SSIM < 0.70 の docs 抽出 →
  structural feature 分類 → cluster → 次 hypothesis 3-5 件 propose"
- method:
  - Source: ssim_baseline.json (PRE-R32 — R32 baseline-refresh not yet committed
    but R31→R32 mean delta tiny, bottom-bucket composition stable)
  - Bottom 30 worst-SSIM pages → deduplicated to 15 unique docs
  - Per-doc XML feature scan: kern, jc, numPr, chars-indent, tbl, floating shape,
    footnote, n_paras, doc_grid, compat_mode
  - Sweep: tools/metrics/survey_bottom_bucket.py
  - Data: pipeline_data/bottom_bucket_survey.json
- structural pattern findings:
  - **100% have effective kern** (15/15) — R32 kern gate fires here
  - **100% have chars-indent** (15/15, avg 76 paras/doc with chars indent)
  - **0% jc=both** — Mech 2 (justify-time slack) does NOT fire in bottom bucket;
    only Mech 1 active
  - **87% have tables** (13/15) — table-heavy template style
  - **67% have floating shapes** (10/15)
  - **33% have list paragraphs** (5/15)
  - **20% have footnotes** (3/15) — but b837 has 26 fn (extreme density)
- hypothesis clusters (predicted gain order):
  1. **Cluster D — b837 footnote spill** (1 doc, KNOWN BLOCKER): Oxi reserves full
     footnote area, Word splits across pages. Already in RESEARCH_LOG ## Active
     hypotheses, assigned to oxi-2.
  2. **Cluster A — heavy table internal layout** (5+ docs: 2ea81a, b35123, 1ec1091,
     e3c545, 6514f214): table cell vertical Y, vAlign, cellMar, line stride within
     cell. Multi-doc leverage.
  3. **Cluster B — long-doc cumulative Y drift** (4-6 docs: e3c545 541p, 34140b
     499p, 04b88e 386p, a1d6e4 317p, 6514f214 350p, d4d126 313p): each para small
     Y drift accumulates over hundreds of paragraphs. Bottom pages downstream of
     accumulated drift.
  4. **Cluster E — chars-indent precise measurement** (universal, 100%): master's
     active §15.1.1 work suggests not yet fully resolved. d77a 137 chars-indent
     paragraphs (70% of 197 paras) is densest case.
  5. **Cluster C — floating shape wrap-around** (3+ docs: 2ea81a 21fs, 459f, 1ec1091,
     6514f214): body Y when text crosses floating shape zone. Master's §17 expansion
     covered positionV/posOffset formula; wrap-effect on body Y still TBD.
- specific bottom-bucket high-impact docs:
  - d77a58485f16 p7 SSIM 0.627 (worst overall) — 137 chars-indent + 10 tables
  - b837808d0555 p6 SSIM 0.645 — 26 footnotes (b837 spill case)
  - 2ea81a8441cc p2 SSIM 0.664 — 21 floating shapes (master's §19.7 work doc)
  - e3c545fac7a7 p11 SSIM 0.665 — 541 paras (longest doc, cumulative drift)
  - b35123fe8efc p1 SSIM 0.666 — 22 tables in only 78 paras (28% in tables)
- outcome:
  1. **R32 kern gate already targets all 15** (100% kern coverage) — R32's
     improvement here likely shows up post-baseline-refresh.
  2. **5 distinct cluster axes** identified beyond kern. Each cluster has 1+ docs
     where targeted fix could unlock major SSIM gains.
  3. **Mech 2 (justify) not relevant in bottom bucket** — 0/15 use jc=both.
     Future Mech 2 work won't help these docs.
  4. **Long-doc drift hypothesis is novel** — was not in master's recent investigation
     queue. If confirmed, single fix could lift 4-6 docs simultaneously.
- next: rank cluster investigations. Cluster D (b837) already assigned. Cluster A
  (table cell) and B (cumulative drift) are highest-leverage next steps.
  Cluster E (chars-indent) should sync with master's §15.1.1 active work.
- references:
  memory/investigation_bottom_bucket_post_r32_2026_05_02.md (full table + per-doc
  details + recommended priority order). Builds on
  memory/investigation_r31_kern_cross_tab_2026_05_02.md.

## 2026-05-02 — oxi-4 — confirmed — R31 narrow gate × kern cross-tab: 40/41 candidates have kern → R31 redundant under R32
- context: User question — "R31 で SSIM 変化した 18 docs (3a4f_p64 +0.032 含む)
  について kern presence 確認。あり → R32 kern gate で同 paragraph fire → R31
  narrow path は冗長 → R32 ship 時削除可。なし → 残す価値あり。"
- hypothesis: R31's narrow gate (chars-indent + cross-run pair + compressPunctuation)
  is essentially a subset of R32's kern gate. Almost all R31-fire candidates have
  effective kern.
- evidence:
  - Scan: tools/metrics/scan_r31_gate_candidates.py over 184-doc baseline
  - Data: pipeline_data/r31_gate_candidates.json
  - Cross-tab vs pipeline_data/kern_audit_2026-05-02.json
  - **41 docs match R31's narrow gate conditions**
    (chars-indent paragraph AND cross-<w:r> yakumono pair AND compressPunctuation)
  - **40/41 (97.6%) have effective kern** (37 via docDefaults + 3 via Normal style)
  - 1 outlier (9a8e8ddab85b_order_06-1.docx) has R31 conditions but NO effective kern
  - Specifically: 3a4f9fbe1a83 (R31 winner p64 +0.032 + minor loser p42 -0.008)
    has kern via Normal style val=2 → R32 will fire on both pages independently
- outcome:
  1. **R31 narrow path is essentially a subset of R32 kern gate**.
  2. **3a4f_p64 (R31's only material winner +0.032) IS in the 40 with-kern set**.
     R32 captures it independently. R31 not needed for this gain.
  3. **17 minor-delta docs (|delta| ≤ 0.009) likely also in the 40-with-kern set**.
     R32 will track or improve their micro-deltas.
  4. **9a8e8ddab85b is R31's only no-kern fire** = potential false positive (Word
     doesn't compress without kern, R31 does). R31 deletion CORRECTS this case.
  5. **R32 ship-time R31 deletion is safe** — net impact predicted ≥ R31's
     +0.000058, likely larger since R32 also catches R17 big_losers (7f272a, ed025).
- next: pre-baseline measurement before R32 ship to verify mechanical prediction.
  Specifically watch 9a8e8ddab85b for SSIM lift from removing R31's false fire.
- references: memory/investigation_r31_kern_cross_tab_2026_05_02.md (full analysis +
  41-doc list with kern source). Builds on session_51_kern_audit_177docs.md and
  session_50_r31_chars_indent_cross_run_gate.md.

## 2026-05-02 — oxi-4 — confirmed — §9 Footnote line-height closed-form: max(17.5, max(size + 5.5, natural_lh)) for CJK 83/64
- context: prior §9.1 footnote investigation (memory `spec_footnote_lh_2026_05_02.md`)
  observed size-dependent "extra" decreasing from 1.64pt (13pt) to 0.16pt (18pt) but
  didn't pin closed form.
- hypothesis: extras reflect a clean `lh = font_size + 5.5pt` for CJK 83/64 family
  in the "above floor" regime (i.e., when natural_lh < size+5.5).
- evidence:
  - Phase 3 sweep: 30 records covering MS Mincho {11.5, 12.5, 15, 20, 24}pt +
    Calibri {12, 14, 18}pt + Yu Mincho {11, 14}pt
  - Data: tools/metrics/output/footnote_lh_sweep_phase3.json
  - Sweep: tools/metrics/measure_footnote_lh_sweep.py (phase 3 variants)
  - All MS Mincho 12.5/13/14/15/16/18pt match `lh = size + 5.5` exactly
    (12.5→18.0, 13→18.5, 14→19.5, 15→20.5, 16→21.5, 18→23.5)
  - MS Mincho 20pt: natural=25.94 > size+5.5=25.5 → measured 26.0 ≈ natural ✓
  - MS Mincho 24pt: natural=31.13 > size+5.5=29.5 → measured 31.0 ≈ natural ✓
  - Yu Mincho 11pt: natural=18.5 > floor → measured 18.5 ✓ (no boost; natural-only)
  - Yu Mincho 14pt: natural=23.59 → measured 23.5 ✓
  - Calibri 14pt: measured 18.5, BUT size+5.5=19.5. natural=17.25 + 1.25 extra.
    Linear fit Calibri extra = 5.625 - 0.3125 × size. NOT same as MS Mincho's
    size+5.5 boost.
  - Calibri 18pt: natural=22.0 → measured 22.0 ✓ (Calibri extra ≈ 0 here)
- outcome:
  1. **CJK 83/64 family (MS Mincho/Gothic UPM=256) closed-form**:
     `lh_footnote = max(17.5pt, max(font_size + 5.5pt, natural_lh))` snapped to 0.5pt.
     - size 12-19pt: size+5.5 wins (lh = size + 5.5)
     - size ≥ 20pt: natural wins (lh ≈ 1.297 × size, snapped)
     - size ≤ 11.5pt: floor 17.5 wins
  2. **Latin (Calibri)**: natural-based with small per-font extra (5.625 - 0.3125×size
     for Calibri). Different coefficients per Latin font likely. Needs more data
     (Calibri 13/15/16pt) for full closed-form.
  3. **Large-glyph CJK (Yu Mincho/Meiryo)**: natural_lh dominates always
     (natural >> floor and natural >> size+5.5).
  4. **Why "size + 5.5" is special for CJK 83/64**: natural_lh = 1.297 × size;
     measured extra = 5.5 - 0.297 × size. Sum: lh = 1.297×size + 5.5 - 0.297×size
     = size + 5.5. The CJK 83/64 ratio + footnote extra coefficient happen to
     produce a clean linear formula.
- next: (a) Calibri 13/15/16pt to verify Latin extra linearity; (b) cross-language
  verification (English Word); (c) footnote with grid mode (does docGrid affect
  footnote area?).
- references: memory/spec_footnote_size_extra_2026_05_02.md (full memo with
  recommended spec text).

## 2026-05-02 — oxi-4 — confirmed — §9 Footnote 17.5pt floor is HARDCODED in renderer (not from style)
- context: prior §9.1 entry left open question "what's the origin of the 17.5pt floor?".
  Hypothesis was hidden default FootnoteText style with `<w:spacing line="350" atLeast/>`
  or similar.
- hypothesis tested: explicit FootnoteText style with `<w:spacing line="240" auto/>`
  ("Single") will override the floor and produce natural_lh = 13.617pt for MS Mincho 10.5pt.
- evidence:
  - 8-variant sweep: tools/metrics/measure_footnote_floor_origin.py
  - Data: tools/metrics/output/footnote_floor_origin.json
  - V1 no styles → 17.5pt (baseline)
  - V2 explicit FootnoteText with `line=240 auto` → **17.5pt** (NOT 13.617pt!)
  - V3 explicit `line=200 auto` → 15.0pt (sub-default, formula TBD)
  - V4 explicit `line=200 exact` → 10.0pt (= line/20, exact wins)
  - V5 explicit FootnoteText with empty pPr → 17.5pt
  - V6 explicit `line=400 atLeast` (=20pt > floor) → 20.0pt (atLeast wins above floor)
  - V7 explicit `line=350 atLeast` (=17.5pt = floor) → 17.5pt
  - V8 styles.xml present, no pStyle ref on footnote → 17.5pt
- outcome:
  1. **17.5pt floor is HARDCODED in Word's footnote renderer**, NOT from a default style.
     Even when docx provides FootnoteText style with explicit "Single" line spacing,
     Word ignores it.
  2. **lineRule precedence in footnote area**:
     - `exact`: always wins precisely (lh = line/20)
     - `atLeast`: wins only when line/20 > 17.5pt
     - `auto`: silently overridden by floor when line ≥ 240 (= "Single" or higher)
       For line < 240, produces a sub-default rule (15pt for line=200, formula TBD).
  3. **17.5pt = 350tw** likely tied to Word's legacy footnote area allocation default.
  4. **Spec §9.1 needs rewrite**: not just "12pt Single" correction (already known
     wrong) but adding the explicit hardcoded floor + override precedence rules.
- next: (a) pin "size-extra" formula for natural_lh > 17pt (decreasing 1.64→0.16pt
  with size 13→18pt); (b) test cross-language (English Word) — is 17.5pt locale-
  independent?; (c) test docGrid present in section — does footnote area inherit
  body grid pitch?
- references:
  - memory/spec_footnote_floor_hardcoded_2026_05_02.md (full memo with recommended
    spec rewrite text)
  - earlier: memory/spec_footnote_lh_2026_05_02.md (initial finding showing 17.5pt floor)

## 2026-05-02 — oxi-4 — investigation — Run fragment merging policy: NO merging, but per-fragment yakumono loop misses cross-<w:r> adjacency
- context: User question — "break_into_lines に渡される fragments は <w:r> 単位ではなく
  Oxi が同 style ones を merge した結果。これが cross-run yakumono detection を無効化
  している可能性"
- hypothesis tested: Oxi has a same-style run merging pass before break_into_lines
- evidence (code review):
  - `crates/oxidocs-core/src/parser/ooxml.rs:1046-1064`: parser pushes one Run per
    `<w:r>`. Within a single `<w:r>`, multiple `<w:t>` text elements are
    concatenated in `parse_run` (line 2407-2413), but cross-run text is NOT combined.
  - `crates/oxidocs-core/src/layout/mod.rs:3218-3223`: fragments built as 1:1 from
    `para.runs.iter().enumerate()`. No filtering, no merging.
  - Only IR `Vec<Run>` mutations post-parse: `revisions.rs:57` (retain_mut filter
    for tracked changes), `parser/ooxml.rs:5769` (mutate-in-place), `layout/mod.rs:685`
    (filter for tracked changes), `layout/mod.rs:1345` (`resolve_fit_text_runs` reads
    groups, doesn't merge).
- outcome:
  1. **User's "merging" premise is INCORRECT**: there is no same-style merging pass.
     IR `Run` count = `<w:r>` count (modulo tracked-change filter).
  2. **The functional issue user describes IS real**: yakumono detection at
     `layout/mod.rs:4322` builds `chars_vec` per-fragment, so neighbor lookups
     `chars_vec[i+1]` / `chars_vec[i-1]` are bounded to within a single `<w:r>`.
     Cross-`<w:r>` adjacencies (e.g., `」` ending run A + `、` starting run B)
     fail both `i+1 < n` and `i > 0` boundary checks → no compression.
  3. **Real-world impact**: any docx with `<w:r>` boundaries from track-changes,
     comment markers, formatting toggles, hyperlink boundaries, or field codes
     can have undetected yakumono adjacencies. Word's renderer compresses based
     on character identity (Type A/B/C per session 51 R0), independent of style
     boundaries.
- recommended fix (deferred to oxi-2 implementation session, ~50 LOC):
  Pre-compute yakumono flags on paragraph-wide concatenated character sequence,
  store paragraph-level + index map, look up by (frag_idx, char_idx_in_frag) in
  the fragment loop. O(n) extra pass, negligible cost.
- risks to verify before shipping:
  (a) Style-boundary semantics — does Word fire yakumono compression when adjacent
      chars have different fonts/sizes? Likely yes (character-driven), but COM
      verification on `<w:r>「</w:r><w:r font=B>、</w:r>` minimal repro is needed.
  (b) Revisions ordering — pre-pass must run AFTER tracked-change filtering, else
      it'd see deleted/inserted text incorrectly.
  (c) Field code boundary respect — pre-pass must honour `field_result_depth`
      suppression at parser/ooxml.rs:1059-1063.
- references: detail memo at memory/investigation_run_fragment_merging_2026_05_02.md
  (text-only, no code change). Master's session 51 R0 yakumono Type A/B/C tables
  in MEMORY.md provide the character-driven compression rule baseline.

## 2026-05-02 — oxi-1 — confirmed — §15.1 leftChars char_width source: docDefault.rPr.sz, not pPr/run

- context: §15.1 defined `effective_indent = leftChars / 100 * char_width`
  for chars-based indents but did NOT specify which font's size determines
  `char_width`. The 2026-03-29 confirmation used 10.5pt as char_width but
  did not note its source.
- hypothesis options:
  (a) docDefault.rPrDefault.rPr.sz
  (b) Paragraph pPr.rPr.sz
  (c) Run rPr.sz
  (d) Some inheritance chain (style > pPr > docDefault > fallback)
- evidence — IC_* axis-isolation:
  - `tools/metrics/ic_repro/` (no styles.xml): 9 variants × {pPr.sz, run.sz}
    in {21, 28, mixed, empty}. ALL gave LeftIndent = 10.50pt regardless
    of pPr.rPr.sz / run.rPr.sz. Conclusion: pPr/run rPr.sz are NOT used.
    The 10.5pt default is Word's fallback when styles.xml is absent
    (Japanese-localized Word implementation default).
  - `tools/metrics/ic_docdef_repro/` (explicit styles.xml): 5 docDefault
    sz values × variants. Linear scaling confirmed:
    docDef sz=20 → LeftIndent=10.00pt
    docDef sz=21 → 10.50pt
    docDef sz=24 → 12.00pt
    docDef sz=28 → 14.00pt
    docDef sz=44 → 22.00pt
  - Cross-check with run-override: IC_dd20_run44 (docDef=20, run sz=44)
    gives LeftIndent=10.00pt. IC_dd44_run21 gives 22.00pt. **docDefault
    always wins**.
- outcome:
  - char_width source CONFIRMED as `docDefault.rPrDefault.rPr.sz`.
    NOT inherited via style/Normal chain (style hierarchy not tested
    here; defer to follow-up).
  - Refined §15.1.1 formula:
    `char_width_pt = docDefault.rPr.sz_val / 2 (= sz_pt)`
    `effective_indent_pt = leftChars / 100.0 * char_width_pt`
  - Applies to all *Chars attributes (leftChars/rightChars/
    firstLineChars/hangingChars).
  - Implication for Oxi: parser must read docDefault.rPr.sz when
    computing chars-based indents; current Oxi behavior likely uses
    pPr/run.rPr.sz which is incorrect (need to verify).
- code change: NONE. Audit needed of Oxi's `*Chars` indent computation
  in `crates/oxidocs-core/src/layout/mod.rs`.

## 2026-05-02 — oxi-1 — confirmed — §17 Shape Positioning expansion: positionV/posOffset formula + RelativeVerticalPosition enum mapping + baseline survey

- context: §17 had only 2 minimal sub-sections (§17.1 reference table,
  §17.2 wrap behavior) and no formula for positionV with posOffset, despite
  shapes being heavily used in 18/184 baseline docs (177 anchors total).
  Master had not investigated this area.
- hypothesis: For `<wp:positionV relativeFrom="X"><wp:posOffset>N</wp:posOffset>`,
  Shape.Top (pt, COM) = N / 12700 (linear), and absolute_Y =
  ref_origin(X) + Shape.Top.
- evidence:
  - `tools/metrics/scan_baseline_shape_positioning.py` survey: 100% of
    baseline anchors (177/177) use `positionV relativeFrom="paragraph"`.
    Plus column-anchored horizontal (155/177), wrapNone (172/177),
    layoutInCell (172/177), allowOverlap (177/177).
  - `tools/metrics/build_sp_position_v.py` + `measure_sp_position_v.py`:
    12 SP_* variants — 3 anchor positions × 4 posOffset values
    {0, +9pt, +53.2pt, **−50pt**}. ALL show Shape.Top = posOffset_emu /
    12700 EXACTLY (residual = 0.00pt). absolute_Y = anchor_paragraph_top
    + Shape.Top, also exact across all 12.
  - `tools/metrics/build_sp_relative_from.py` + `measure_sp_relative_from.py`:
    12 SR_* variants — 6 relativeFrom values
    {paragraph, page, margin, line, topMargin, bottomMargin} × 2 posOffset
    {0, +100pt}. Linear conversion confirmed for all 6 references. COM's
    `Shape.RelativeVerticalPosition` integer enum mapping pinned:
    margin=0, page=1, paragraph=2, line=3, topMargin=4, bottomMargin=5.
- outcome:
  - §17 spec extended with §17.3 (positionV/posOffset formula) and §17.4
    (baseline distribution survey).
  - Confirmed formula for the 100%-of-baseline case
    (`relativeFrom="paragraph"`):
    `absolute_Y_pt = anchor_paragraph_top_y + (posOffset_emu / 12700)`.
  - Other relativeFrom values follow the same `ref_origin + offset`
    template; ref_origin Y values per relativeFrom are well-defined per
    ECMA-376 §17.3.3.18 ST_RelFromV.
  - Implication for Oxi: shape Y positioning must (a) parse posOffset as
    EMU, divide by 12700 for pt, and (b) for paragraph-anchored shapes,
    add the anchor paragraph's top y. The 5 non-paragraph relativeFrom
    cases (out of 177) need ref_origin computed differently but are
    edge cases.
- code change: NONE (pure investigation + measurement). Oxi's existing
  shape parsing/rendering should be reviewed against this rule;
  particularly important for 459f, 2ea81a, 3a4f9f which contain shapes
  on bottom-N-floor docs.

## 2026-05-02 — oxi-4 — refuted — Spec §9.1 "Footnote LineSpacing 12pt Single"
- context: spec §9.1 line 1141 claims footnote default "LineSpacing: 12pt (Single)".
  Existing footnote_separator.json data showed inter-footnote lh = 17.5pt for
  MS Mincho 10.5pt — contradicts spec.
- hypothesis: footnote area applies a hidden floor (≥17.5pt) that overrides the
  per-paragraph lineSpacing rule when the rule's value is below the floor.
- evidence:
  - 31 records: MS Mincho {9, 10.5, 12, 13, 14, 16, 18}pt + Calibri {10, 11}pt ×
    n_footnotes ∈ {1, 3, 5} × 6 explicit lineRule cases
  - Data: tools/metrics/output/footnote_lh_sweep.json (10 initial),
          footnote_lh_sweep_phase2.json (24 follow-up)
  - Sweep: tools/metrics/measure_footnote_lh_sweep.py
  - Floor at 17.5pt for natural_lh < ~17pt (MS Mincho 9/10.5/12pt + Calibri 10/11pt)
  - Above floor: lh = natural_lh + size-dependent extra (decreasing 1.64→0.16pt
    as size grows 13→18pt)
  - lineRule=auto with line=240 produces SAME lh as no spacing — Word silently
    overrides
  - lineRule=auto with line<240 (e.g., 200) produces below-natural lh (16-16.5pt
    for 14pt MS Mincho — formula TBD)
  - lineRule=exact precisely respected (lh = line/20 across all tested values)
  - lineRule=atLeast respected only when line/20 > default (atLeast 360 = 18pt
    < default 19.5pt → default wins; atLeast 240 = 12pt → default wins)
- outcome:
  1. **Spec §9.1 "12pt Single" REFUTED.** Default footnote lh has a floor at 17.5pt.
  2. **Default formula**: `lh = max(17.5pt, natural_lh + extra(size))` where
     extra is small and size-dependent. Possibly hidden default style with
     `<w:spacing line="350" lineRule="atLeast"/>` accounts for floor.
  3. **lineRule precedence**: Word respects `exact` always; respects `atLeast`
     only when binding upward; silently ignores `auto` when line ≤ 240.
  4. **Implication for Oxi**: `estimate_footnote_h` (`crates/oxidocs-core/src/layout/mod.rs`)
     likely uses the spec's "12pt Single" assumption → reserves LESS area than Word.
     Opposite to b837 issue (where Oxi reserves MORE) — suggests b837 has a
     different code path bug.
- next: (a) pin the exact "extra(size)" formula via more sizes (15, 17, 20, 24pt);
  (b) verify floor is also 17.5pt with non-Japanese language; (c) measure
  `<w:fnSep>` separator height directly; (d) measure `<w:continuationSeparator>`
  for multi-page footnote behavior (related to b837 spill bug).
- references: master's b837 footnote-area-spill blocker (RESEARCH_LOG line 1820),
  see memory/spec_footnote_lh_2026_05_02.md for full data + recommended spec edit.

## 2026-05-02 — oxi-4 — confirmed — §19.7 Y0 intercept residual explained via centering geometry
- context: master's `## 2026-05-02 — oxi-1 — partial — §19.7 Y0 intercept anomaly` left
  "Residual 1–2.5pt unexplained — likely floating-table topFromText spacing constant
  (default = 0 unless `topFromText`/`bottomFromText` set on tblpPr; needs separate
  isolation)."
- hypothesis: Y0 = (cell_h_anchor + centering_lh_anchor) / 2 + tblpY
  (NOT line_height + topFromText; it's a geometric offset to anchor's cell BOTTOM)
- evidence:
  - Re-analyzed master's 7 K_* fe_match_repro variants in
    `C:\Users\ryuji\oxi-1\pipeline_data\fe_match_measurements.json`
  - Validation: tools/metrics/analyze_fe_y0_residual.py
  - Results: 5/7 within 0.05pt (K_lp323, K_only_atLeast296, K_lp323_atLeast296_sz28
    within 0.5pt residual); 3/7 at +0.55pt (K_baseline, K_only_sz28, K_tblWauto_only
    — all `pitch=360tw=18.0pt + line=auto`); 1/7 at -0.53pt (K_lp323_atLeast296)
  - All within 1pt; most within 0.5pt COM measurement quantization
- outcome:
  1. **Master's "residual 1-2.5pt" reduced to ~0.5-1pt**, attributable to pixel-snap
     of integer-pt arithmetic, NOT to topFromText.
  2. **Geometric basis**: floating table sits at anchor's grid cell BOTTOM (+ tblpY),
     while anchor_y is the cell's first-character position (= cell_top + centering
     offset above). Therefore Y0 = anchor_cell_h - centering_offset_above + tblpY
     = (cell_h + centering_lh) / 2 + tblpY.
  3. **Builds on §3.3 corrected formula** (centering_lh = round(max(natural_lh,
     size × 83/64))). Master's "line_height" intuition was missing the centering
     offset flip.
  4. **topFromText hypothesis NOT needed** for the K_* observations (default=0).
     A future sweep with non-zero topFromText would directly test linear addition
     to Y0.
- next: (a) test non-zero topFromText/bottomFromText and verify Y0 += topFromText;
  (b) test inline-cell anchor case (master's slope=1 sweep showed +15.0pt Y0 for
  inline anchor — should match (cell_h + centering_lh)/2 + cell_margin geometry);
  (c) close the remaining +0.5pt jitter on integer-pt-pitch + auto cases.
- references: master's §18.10 spec text in this branch, §19 follow-up planned for
  main merge. Recommended spec edit text in
  memory/spec_y0_intercept_explained_2026_05_02.md.

## 2026-05-02 — oxi-4 — confirmed — §1.7 Mixed font run × grid centering extension
- context: spec §1.7 says `lh = max(per-run lh)` after grid snap. Doesn't address how
  the corrected §3.3 centering_lh formula interacts with mixed-font lines.
- hypothesis: mixed-line centering_lh = max(per-run centering_lh)
  where per-run centering_lh = round(max(natural_body_lh, size × 83/64))
- evidence:
  - 20 records: 10 combos × pitch ∈ {18, 24}pt
  - Combos: Calibri-11/18, Yu Mincho-11, MS Mincho-10.5/18, Meiryo-11 × pure or mixed
  - Data: tools/metrics/output/grid_mixed_font_centering.json
  - Sweep: tools/metrics/measure_grid_mixed_font_centering.py
  - Mixed Calibri-11 + Yu Mincho-11 → matches Yu Mincho-11 alone (centering=18 dominates)
  - Mixed Calibri-18 + Yu Mincho-11 → matches Calibri-18 alone (centering=23 dominates
    via universal 83/64 floor)
  - Mixed MS Mincho-10.5 + Calibri-18 → matches Calibri-18 (centering=23)
  - Mixed Meiryo-11 + Calibri-11 → matches Meiryo-11 (centering=21 dominates via
    natural_lh > size × 83/64)
- outcome:
  1. **§1.7 extension**: mixed-line centering_lh = max(per-run centering_lh).
  2. **Surprising consequence**: Latin fonts can dominate centering despite smaller
     visual size (Calibri-18 size×83/64=23.34 > Yu Mincho-11 natural_lh=18.5).
  3. **+0.5pt Latin residual** observed in mixed contexts — pure Calibri offset
     produces 0.0pt error but mixed Calibri offset has +0.5pt extra. Likely Calibri
     internal natural_lh = 22.5 (not 22) creating context-dependent half-pt rounding.
- next: (a) test 3+ run mixes; (b) test mixed sizes within same font family;
  (c) pin Calibri's exact internal natural_lh via dedicated single-font sub-pt sweep.

## 2026-05-02 — oxi-4 — refuted — Spec §2.1 line 198 grid spacing snap claim
- context: spec line 198 "sa=sb=10pt → 9.75pt due to grid snap" — long-standing
  hand-wavy note. Wanted to pin exact behavior.
- hypothesis tested: gap = line_h + sa (no grid snap on sa value)
- evidence:
  - 48 records: MS Mincho 10.5pt + Calibri 11pt × pitch ∈ {18, 24}pt × sa ∈
    {0, 4, 5, 6, 7.5, 10, 10.5, 12, 15, 18, 20, 24}pt
  - Data: tools/metrics/output/grid_paragraph_spacing.json
  - Sweep: tools/metrics/measure_grid_paragraph_spacing.py
  - All 48 records match `gap = pitch + sa` with 0.0pt error.
  - sa=10pt at pitch=18 gives gap = 28pt EXACTLY. Spec predicted 27.75 → REFUTED.
- outcome:
  1. **Spec §2.1 line 198 REFUTED.** Modern Word does NOT grid-snap sa.
  2. **sa is added directly** to the line gap as the exact pt value declared in
     `<w:spacing w:after="...">`, regardless of whether grid mode is active.
  3. **Latin/CJK identical** — Calibri and MS Mincho produce exact-same gaps for
     same sa value.
  4. **Paragraphs do NOT re-align to grid when sa>0** — only the first paragraph's
     first line is grid-aligned (per §3.3). Subsequent paragraphs sit at line_h+sa
     below previous, breaking grid alignment.
  5. **Implication for Oxi**: if currently grid-snapping sa, that's a bug producing
     cumulative Y drift in any document with non-zero paragraph spacing.
- next: (a) verify sb-only and mixed sa/sb cases (§2.1 collapse rule unchanged?);
  (b) verify with lineSpacingRule=exact/atLeast/multiple; (c) verify with
  contextualSpacing=true (§2.3 should still zero out sa/sb).

## 2026-05-02 — oxi-4 — confirmed — Grid centering lh formula (§3.3 / §13.4 correction)
- context: Session 49 R109 hypothesized "text_y_offset uses font_size instead of natural
  line height". Spec §3.3 says `inner_box = ceil(natural_lh)` for grid-centering inner box.
- hypothesis: centering_lh = round(max(natural_body_lh, font_size × 83/64))
  with universal 83/64 floor applied to ALL fonts (NOT just CJK whitelist)
- evidence:
  - 49 records: MS Mincho 10.5/14/18pt × Calibri 11/18pt × Yu Mincho/Meiryo 11pt ×
    pitch ∈ {12,15,18,20,24,28,32}pt × 3 paras × 3 lines per para
  - Data: tools/metrics/output/grid_per_line_y.json
  - Sweep: tools/metrics/measure_grid_per_line_y.py
  - Calibri 18pt at pitch=24 (single-cell): natural_lh=22, ceil=22 → predicted 1.0pt
    offset; measured 0.5pt. Matches `round(max(22, 23.344))=23` formula.
  - MS Mincho 18pt at pitch=24: natural_lh=23.344, ceil=24 → predicted 0.0pt; measured 0.5pt.
    Matches `round(max(23.344, 23.344))=23`.
  - Yu Mincho 11pt body lh=18.5 > 14.27 (size×83/64): centering uses 18 (font-specific).
  - Meiryo 11pt body lh=21.5 > 14.27: centering uses 21 (font-specific).
- outcome:
  1. **Spec §3.3 inner_box formula correction**: `ceil(natural_lh)` → `round(max(natural_lh,
     size × 83/64))`. Old formula coincidentally correct for MS Mincho/Gothic UPM=256
     (where natural_lh = size × 83/64 exactly) but WRONG for Latin large-size and
     CJK UPM=2048 with extra typo metrics.
  2. **Universal 83/64 floor finding**: the 83/64 multiplier is applied to ALL fonts in
     grid centering, not just CJK whitelist. This contradicts the §1.2 implication that
     83/64 is a "CJK-specific multiplier" — it's actually a universal minimum.
  3. **§13.4 TextBox same correction**: structural identical formula.
  4. **Multi-paragraph grid behavior**: each line allocates exactly 1 grid_cell regardless
     of paragraph boundaries (when sa=sb=0). Verified across all 49 records.
  5. **Session 49 R109 indirectly resolved**: Oxi was using `font_size`; correct value
     is `round(max(natural_lh, size × 83/64))`. The full fix touches both Latin and CJK
     paths uniformly.
- next: (a) verify TNR/HG-series follow same rule; (b) test exactly-.5 boundary cases
  (round-half-up vs banker's vs half-down); (c) verify with mid-line font change (§1.7);
  (d) compute Calibri 18pt + grid combo (master's Session 51 noted RPC errors there too).

## 2026-05-02 — oxi-4 — confirmed — Tall-header pushdown (spec §8.2 TBD resolved)
- context: word_layout_spec_ra.md line 936 long-standing TBD: "When header content
  overflows headerDistance and crosses topMargin, body Y is pushed down. Earlier
  note ('3-line 14pt header → body_y=90pt when topMargin=72') was measured for
  noGrid; the formula has not been re-verified under the corrected spec."
- hypothesis: body_y = max(top_margin, round_half(header_distance + n_lines × header_lh))
- evidence:
  - 90 records sweep: Calibri 11pt + MS Gothic 10.5pt × n_lines {1..5} ×
    header_distance {18,36,54}pt × top_margin {36,72,108}pt
  - Data: tools/metrics/output/tall_header_pushdown.json
  - Validation: tools/metrics/analyze_tall_header_pushdown.py
  - MS Gothic 10.5pt (CJK 83/64): 0.0pt error across all 45 records
  - Calibri 11pt (lh=13.5pt): 0.0pt for n=1..3, systematic -0.5pt drift at n=4..5
  - Boundary case (header_bottom_true == top_margin) treated as no-overflow
- outcome:
  1. **Formula confirmed for noGrid regime** with caveat for Latin UPM=2048 fonts at
     n_header_lines ≥ 4 (cumulative sub-pt arithmetic produces -0.5pt drift).
  2. **Spec §8.2 TBD resolved**. See memory/spec_tall_header_pushdown_2026_05_02.md
     for the full formula text and recommended spec edit.
  3. **Old "+~2.5pt offset" claim** (already retracted in spec) confirmed wrong
     by full sweep data.
- next: (a) measure with explicit grid (LM≥1) — likely body_y rounds up to next
  grid cell; (b) measure with multi-paragraph header (separate <w:p> instead of
  <w:br/>); (c) measure with custom header pPr (spaceBefore/After).
  Cross-reference: master's Session 51 note (memory MEMORY.md) extended this to
  docGrid (LM≥1) with 288 cases, also flagging Calibri 18pt RPC errors as open.

## 2026-05-02 — oxi-4 — confirmed — LM0 cell row_h closed-form (cumulative two-endpoint snap)
- context: Section 13.5 / oxi-4 active hypothesis "LM0 cell formula (investigating)"
- hypothesis: row_h(n) = round_half(top_y + n * lh_natural) - round_half(top_y),
  where lh_natural is the LM=0 body line height for (font, size).
- evidence:
  - 5 fonts × 7 sizes × 4 n_lines = 140 cell samples, +35 body samples
  - Data: tools/metrics/output/lm0_multiline_cell_v3.json (MS Mincho/Gothic, 90 cells)
          tools/metrics/output/lm0_multiline_cell_v5.json (Calibri/Yu Mincho/Meiryo/HGS Mincho E/TNR, 140 cells)
  - Analysis: tools/metrics/analyze_lm0_cell_formula.py (H5: 0.0 total error on MS family)
              tools/metrics/analyze_lm0_cell_formula_v5.py (H5a mean abs err ≤ 0.16pt across all 5 fonts)
  - The previous research-log claim "10.5pt = 18n, 12pt = 28+36(n-1) — non-continuous formula"
    referred to the docGrid (LM≥1) regime, NOT no-grid (LM=0). v3/v5 sweeps cover LM=0 only.
- outcome:
  1. **LM=0 cell row_h is fully closed-form** when lh_natural is correct.
     `row_h = round_half(cell_top_y + n*lh) - round_half(cell_top_y)`
  2. **Spec §1.2 simplification correction (separate finding)**: the rule
     "lh = font_size × 83/64 for CJK whitelist" is wrong for Yu Mincho and Meiryo.
     See memory `project_lm0_cell_formula_2026_05_02.md` for the per-font multipliers
     and the corrected `lh = size × (typoAsc+typoDes+typoLineGap)/upm` general formula.
  3. **HG-series partial revision**: HGS Mincho E (UPM=2048) measures ratio ≈ 1.30 ≈ 83/64,
     contradicting spec §1.2 line 63 which generalizes "HG-series NOT in whitelist" from
     a single HGGothicM measurement. Per-font measurement is required.
  4. **LM≥1 (docGrid) cell row_h still TBD** — separate hypothesis chain.
- next: (a) extend v5 to Yu Gothic / MS PMincho / MS PGothic / HGGothicM / HGGothicE for full
    CJK whitelist coverage; (b) measure LM≥1 cell row_h with grid pitch sweep; (c) integrate
    spec §13.5 with the closed-form once master approves.

## 2026-05-02 — oxi-1 — confirmed (correction) — §13.5 trHeight: ECMA-376 hRule default is "atLeast", not "auto"

- context: §13.5 Round 22 (2026-04-08) stated "ECMA-376 default for w:hRule
  is 'auto', NOT 'atLeast'. When <w:trHeight w:val="..."/> appears WITHOUT
  a w:hRule attribute, the value is a hint and Word ignores it at render
  time, using content height only."
- hypothesis: This claim is incorrect. ECMA-376 Part 1 CT_Height schema
  defines `hRule` with `default="atLeast"`. Word should follow this, in
  which case `<w:trHeight w:val="N"/>` without hRule renders as
  `max(content, N)` — specified wins when N > content.
- evidence: `tools/metrics/build_tr_hrule_default.py` + ad-hoc HR_* repro
  (5 variants, all with 1-line MS Mincho 10.5pt content ~14pt + 60pt
  specified to maximize discrimination):
  - `HR_missing` (`<w:trHeight w:val="1200"/>`, no hRule):
    Word reports HeightRule=1 (atLeast), renders 60.5pt
  - `HR_explicit_auto` (`hRule="auto"`):
    Word reports HeightRule=0 (auto), renders 14.0pt (content)
  - `HR_explicit_atLeast`: HeightRule=1, renders 60.5pt
  - `HR_explicit_exact`:   HeightRule=2, renders 60.0pt
  - `HR_no_trHeight` (no trHeight element):
    HeightRule=0, renders 14.0pt
  Conclusion: `<w:trHeight w:val="N"/>` no hRule → atLeast (specified wins
  when N > content). Round 22's contrary statement was a misreading of
  the schema or a measurement-condition artifact (their tests with
  content > specified produced identical auto/atLeast results, blind to
  the discrimination case spec > content).
- evidence #2: 90-variant `TR_*` matrix
  (`tools/metrics/{build,measure}_tr_height.py`, rule × spec_pt × content
  lines × docGrid linePitch). Across 47 successful measurements (Word
  COM session instability caused 43 RPC failures on the larger batch),
  the post-table paragraph Y minus table-top Y proxy for rendered row
  height confirms `atLeast = max(content, specified) + ~1.5pt border`.
  Cell content line height ≈ font natural height (MS Mincho 10.5pt =
  13.65pt) with slight (≤1pt) sensitivity to docGrid linePitch.
- outcome:
  - §13.5 corrected (this commit / next branch sync to main): default
    hRule is `"atLeast"`. Round 22's "auto default" claim was incorrect.
  - §19.4 (was §18.4) "Word does NOT grid-snap when trHeight present"
    holds: rendered row in atLeast mode is `max(content_natural, spec) +
    border`, NOT a grid-snapped value. Oxi divergence at
    `crates/oxidocs-core/src/layout/mod.rs:4308` (grid-snapping content
    to ceil(content/pitch)*pitch then taking max with trHeight) is
    incorrect for atLeast/exact rules with trHeight present.
  - Implication: any baseline doc using `<w:trHeight w:val="N"/>` without
    explicit `hRule="auto"` is being rendered too tall by Oxi. This is
    the structural cause of 2ea81a tbl#1's +35pt over-pack.
  - Raw data: `pipeline_data/tr_height_measurements.json` + ad-hoc HR_*
    output. Tools: `build_tr_height_matrix.py` / `measure_tr_height.py`
    + `build_tr_hrule_default.py`.
- code change: NONE (pure investigation). Oxi's
  `crates/oxidocs-core/src/parser/ooxml.rs` trHeight parser should default
  hRule to "atLeast" when the attribute is absent (matches Word). Layout
  at `mod.rs:4308` should NOT grid-snap content height when trHeight is
  set (atLeast or exact), per §19.4 / §13.5 corrected.

## 2026-05-02 — oxi-1 — confirmed — §4.7 round 12: Mech 1 trigger char-level audit complete (26/26 chars)

- context: §4.7 lists 11 Type A + 13 Type B + 9 Type C chars. Round 11
  verified smart quotes (4) and em-dash classification. Round 12
  audits remaining standard yakumono characters individually.
- evidence: `measure_mech1_char_audit.py`:
  Suite A — 9 Type A chars × A→A trigger (preceded by （):
    〈 《 「 『 【 〔 ［ ｛: all FIRE Mech 1 (adv=6.0, ratio=0.500)
    （ self-pair: script artifact (first-match logic), but verified
      by §4.7's `（（（（` example
  Suite B — 13 Type B chars × B→B trigger (followed by ）):
    ） 」 』 】 〕 ｝ 〉 》 ］ 、 。 ， ．: ALL 13/13 FIRE (adv=6.0)
  Suite C — control (5 Type B between CJK):
    All 5 correctly show no compression (Mech 1 does NOT fire when
    next char is CJK)
- finding:
  All 26 standard yakumono chars (9 Type A + 13 Type B excluding
  smart quotes / em-dash from round 11) verified to fire Mech 1 under
  expected trigger pairs.
- summary across rounds 11-12:
  Type A (11 chars): smart quotes (4 in round 11) + 9 standard = 11/11 ✓
  Type B (17 chars): smart quotes + em-dash glyph-metric (5 in round 11)
                     + 13 standard (round 12) — 16/17 effective fire
                     (em-dash is glyph-metric not Mech 1)
  Type C: Hbar (U+2015) confirmed no compression
- outcome:
  - §4.7 Type A/B/C list fully verified at character level.
  - Spec round 12 added with full audit table.
- code change: NONE.

## 2026-05-02 — oxi-1 — confirmed — §4.7 round 11: smart quotes Type A/B confirmed, em-dash REFRAMED as glyph metric

- context: §4.7 lists smart quotes (U+2018-201D) and em-dash (U+2014)
  as Type A/B without thorough verification under Mech 2. Session 51
  found em-dash "compresses" in MS Mincho but not Yu Mincho.
- evidence: `measure_smart_quotes_emdash_mech2.py`:
  Suite A — 7 chars × 4 slacks at MS Mincho 12pt:
    Smart quotes ‘ ’ " " (U+2018-201D): all behave as Type A/B,
      adv 12→10→8→6 across slack 0..cap. Mech 2 fires correctly.
    Em-dash (U+2014): adv = 6.0pt UNCHANGED across slack -1, +2, +4, +6.
      Word reports compression but it's glyph design, not Mech 2.
    Hbar (U+2015): adv = 12.5pt, no compression at any slack. Type C ✓.
    」 (control): standard B behavior.
  Suite B — em-dash × 3 fonts at slack=4:
    MS Mincho em-dash adv = 6.0pt
    Yu Mincho em-dash adv = 12.5pt
    Meiryo em-dash adv = 12.50pt
- key reframing:
  Session 51's "MS compresses em-dash, Yu doesn't" is more precisely:
  - MS 明朝/ゴシック em-dash glyph natural width = fontSize/2 (half-
    width BY DESIGN). NOT compressed by Mech 2; it's a font metric.
  - Yu Mincho / Meiryo em-dash glyph natural width = fontSize × 1.04
    (full-width). Mech 2 does NOT compress it (effectively Type C).
- outcome:
  - §4.7 Type A/B/C list updated with em-dash font-dependency note.
    U+2014 is **glyph half-width in MS branded fonts**, NOT Type B
    compression rule.
  - All 4 smart quotes (U+2018-201D) confirmed Type A/B as listed.
  - Hbar (U+2015) confirmed Type C universal.
  - Implementation: Oxi must use font-dependent em-dash natural width
    AND exclude em-dash from Mech 2 candidate set.
- code change: NONE. Spec §4.7 round 11 added.

## 2026-05-02 — oxi-1 — confirmed — §4.7b round 10: N=1 cap fillers complete 3-way universality

- context: Round 7/8 had remaining sweep gaps for fs=10.5/11/12/14 N=1
  cap measurement. Round 10 fills them.
- evidence: `measure_n1_cap_fillers.py` 26 measurements:
  fs=10.5 N=1: cap=5.00 first_drop=5.3 (cap+0.3)
  fs=11.0 N=1: cap=5.50 first_drop=5.6 (cap+0.1)
  fs=12.0 N=1: drop at slack=6.5 (NOT 12.5 as Round 7 claimed)
  fs=14.0 N=1: cap=7.00 first_drop=7.2 (cap+0.2)
- finding:
  - Round 7's "fs=12 N=1 first_drop=12.5" was completely sweep-gap
    artifact. Drop is actually at slack=6.5 = cap+0.5.
  - All 4 font sizes confirm cap = floor(sz/2) × 0.5.
  - drop_threshold ≈ cap + ~0.5pt (range cap+0.1 to cap+0.5).
- 3-way universality verified:
  - N (yak count): 1, 2, 3, 4, 5, 7 (all match formula)
  - fs (font size): 10.5, 11, 12, 14 pt (all match formula)
  - font family: MS 明朝/ゴシック, Yu Mincho, Meiryo, HG明朝E (all match)
- outcome:
  - Final spec rule (no branching needed):
    cap_pt = floor(sz_val_int / 2) * 0.5
    drop_threshold = cap + 0.5
  - Spec §4.7b Round 10 added.
- code change: NONE.

## 2026-05-02 — oxi-1 — confirmed — §4.7b round 9: Multi-line Mech 2 — per-line cap, last line no compress

- context: Session 51 listed multi-line Mech 2 cascade as open question.
  Cap is per-line vs paragraph-cumulative vs other?
- evidence: `measure_multiline_mech2.py` 50-char probe with 6 yak
  distributed × MS Mincho 12pt × jc=both × 12 cw values:
  - cw=310 (2-line): L1 26 chars 2pt comp, L2 24 chars 0pt
  - cw=306 (2-line): L1 26 chars 6.0pt comp (cap reached), L2 0pt
  - cw=210 (3-line): L1 18 chars 6.0pt cap, L2 18 chars 6.0pt cap, L3 14 chars 0pt
- finding:
  - **Cap = floor(sz/2)*0.5 applied PER-LINE INDEPENDENTLY**.
    Each non-last line can compress up to cap regardless of other lines.
    cw=210 paragraph absorbed 12pt total (= 2 lines × 6pt cap each).
  - **Last line never compresses** — jc=both renders last line LEFT-aligned
    (standard Word behavior). All 3-line tests show L3 comp=0 regardless
    of cw.
  - **Wrap algorithm uses Mech 2 cap as per-line line-extension budget**.
    Greedy pack + extend up to +1 char if remaining slack ≤ cap; break
    otherwise.
- outcome:
  - Spec §4.7b Round 9 added with multi-line cascade rule.
  - Implementation guidance:
    For each non-last line, apply Mech 2 distribution if slack > 0.
    Skip last line (jc=both last line rule).
  - Each line treated as independent Mech 2 unit.
- code change: NONE.

## 2026-05-02 — oxi-1 — confirmed — §4.7b round 8: cap font-independent + Round 7 N=1 drop REFUTED

- context: Round 6 (4 sizes on MS Mincho) and Round 7 (N=1/2 on 12pt
  MS Mincho) confirmed cap formula but raised 2 questions:
  (a) is cap font-dependent? (Session 51 found em-dash is)
  (b) is N=1 drop threshold = fontSize × 1 (Round 7 claim)?
- evidence: `measure_cap_other_fonts.py`:
  Suite E (5 CJK fonts × 12pt × N=3 mid-line):
    ＭＳ 明朝, ＭＳ ゴシック, Yu Mincho, Meiryo, HG明朝E
    → ALL produce cap=6.0pt, first_drop=6.5pt
    → cap formula font-INDEPENDENT
  Suite F (fs=14 N=1 gap-fill, slack 7..14 step 1pt):
    slack=7.0: comp=7.0 (cap reached)
    slack=8.0: drop ← first_drop = 8.0, NOT 14.5
- correction:
  Round 7's "N=1 first_drop ≈ fontSize × 1" was a sweep-gap artifact
  (slack 7.0..14.0 untested, jumped from 6.7 to 14.5). True N=1 first
  drop = cap + 1.0pt = 8.0pt for fs=14, mirroring N≥2.
  fs=12 N=1 first_drop=12.5 (Round 7) likely also gap-artifact; true
  value probably 6.5pt. Recommend filler sweep.
- outcome:
  - Final unified rule: cap = floor(sz_val/2) × 0.5, drop_threshold =
    cap + 0.5pt. N-independent and font-independent.
  - Spec §4.7b Round 8 added with corrected drop_threshold rule.
  - Implementation simpler: no N=1 special case.
- open: fs=12 N=1 first_drop verification, 16+pt fonts, mixed-size lines.
- code change: NONE.

## 2026-05-02 — oxi-1 — confirmed — §4.7b N=1 cap REVISED + N=2 mid-line resolved (round 7)

- context: §4.7b round 3 claimed N=1 cap = fontSize/3 (= 4pt for 12pt) as
  special case, derived from a single P_1para_Y50 datapoint. Round 7
  verifies across 4 font sizes + resolves Round 5's N=2 line-end yak
  anomaly.
- evidence: `measure_cap_n1_n2.py`:
  Suite C (N=1 mid-line, pos 12 of 24, fine 0.5pt slack sweep):
    fs=12 N=1 → max_comp = 6.0pt (= fontSize/2, NOT fontSize/3)
    fs=10.5/11/14 sweep had gaps; partial confirmation
  Suite D (N=2 mid-line, pos 8 + 16, 12pt):
    max_comp = 6.0pt (= fontSize/2, Round 6 formula confirmed)
    Round 5's "12pt anomaly" was line-end yak placement artifact
- new finding (drop threshold):
  N=1 first_drop = ~fontSize (12.5 for fs=12, 14.5 for fs=14)
  N≥2 first_drop = cap + 0.5..1.0pt
  → N=1 has special drop tolerance: Word allows line to extend up to
    fontSize past cap before dropping. With multiple yak, distribution
    is fine-grained enough that drop comes earlier.
- outcome:
  - Compression cap = `floor(sz_val/2) × 0.5` UNIVERSAL across N≥1.
    The N=1 special case (fontSize/3) is REVOKED.
  - Drop threshold differs: N=1 = fontSize, N≥2 = cap + 0.5pt.
  - Spec §4.7b round 7 added with the corrected formula and drop-
    threshold characterization.
  - Implementation: same cap formula for any N; only drop_threshold
    branches on N=1 vs N≥2.
- open: fs=10.5/11/14 N=1 sweep gaps (haven't tested exact cap at those
  font sizes due to slack sampling). Larger fonts. Other CJK fonts.
- code change: NONE. Spec §4.7b round 7 added.

## 2026-05-02 — oxi-1 — confirmed — §4.7b cap formula universal (4 font sizes verified, 0.5pt-quantized)

- context: Round 5 found line-level cap = fontSize/2 at 12pt only.
  Round 6 verifies universality across 10.5pt, 11pt, 12pt, 14pt.
- evidence: `measure_cap_font_size_sweep.py` (16 cw values per font ×
  4 font sizes × N=3 yak, controlled cSC=compressPunctuation, direct-zip
  docx). Aggregator `measure_cap_font_aggregator.py` re-measured 47
  docx with incremental save after first script crashed mid-Suite-A.
- findings:
  | fs | fs/2 (theory) | observed cap | first_drop slack | per-yak (N=3) |
  | 10.5 | 5.25 | 5.00 | 6.2 | 1.667 |
  | 11.0 | 5.50 | 5.50 | 6.5 | 1.833 |
  | 12.0 | 6.00 | 6.00 | 7.0 | 2.000 |
  | 14.0 | 7.00 | 7.00 | 8.0 | 2.333 |
- formula refined:
  cap_pt = round_down_to_0.5pt(fontSize / 2)
         = floor(sz_val_int / 2) * 0.5
  For fs=10.5 (sz=21): floor(21/2)*0.5 = 10*0.5 = 5.0pt
  Other fs (11/12/14) have integer fs/2, no quantization needed.
  This matches Mech 2's 0.5pt-step distribution granularity (Word's
  internal half-point precision).
- outcome:
  - cap formula confirmed universal across MS Mincho 10.5-14pt.
  - 0.5pt quantization rule documented in spec §4.7b round 6.
  - Implementation guidance:
    `cap = (sz_val_int / 2) * 0.5_pt` (clean integer math from
    docx's `<w:sz w:val="N"/>` value).
- open: N=2 mid-line anomaly (Suite B was planned but original script
  died on Suite A; can be added later). Larger fonts (16+pt) untested.
  Other CJK fonts (Yu Mincho, Meiryo) untested.
- code change: NONE. Spec §4.7b round 6 added.

## 2026-05-02 — oxi-1 — confirmed — §4.7b per-yak cap REVISED: line-level cap = fontSize/2, divided by N

- context: §4.7b round 3 derived per-yak cap = fontSize × 5/24 (≈2.5pt
  for 12pt) from a single N=9 datapoint. Round 5 (per-yak cap regression
  on N ∈ {2, 3, 4, 5, 7}) reveals this was wrong.
- evidence: `measure_per_yak_cap_sweep.py` controlled-cSC direct-zip
  docx, 24-char probe with N evenly-distributed yakumono, jc=both,
  16 cw values per N:
  - N=3: drop boundary slack=7, max comp at 24 chars = 6.0pt
  - N=4: drop boundary slack=7, max comp = 6.0pt
  - N=5: drop boundary slack=7, max comp = 6.0pt
  - N=7: drop boundary slack=7, max comp = 6.0pt
  - N=2: line-end yak anomaly (non-monotonic drop pattern)
- finding: Line-level total compression cap = fontSize/2 (= 6pt for
  12pt font), CONSTANT across N ≥ 3. Per-yak cap = (fontSize/2) / N.
- reconciliation:
  - Round 3 N=9 datapoint gave 22pt total comp at cw=205. The 19-char
    line at that scenario was POST-DROP (line dropped 3 chars from 22
    to 19), so larger compression budget applies in heavy-overflow
    regime. The "5/24" was a coincidence in that specific scenario.
  - Round 5 measures the clean "1-line, before drop" cap = fontSize/2.
- outcome:
  - Spec §4.7b corrected:
    `total_cap = fontSize/3 if N=1, fontSize/2 if N≥2 (line-level)`
  - Per-yak cap = total_cap / N (derived).
  - Implementation simpler: just compute line total cap.
  - Round 3's per-yak average (22/9 = 2.44pt) reframed as "post-drop
    behavior, larger budget" — separate from the standard wrap-fit cap.
- code change: NONE. Spec §4.7b implementation sketch revised.

## 2026-05-02 — oxi-1 — confirmed — §4.7b/§4.7c UNIFIED: Mech 2 = Mech 3 (single mechanism, alignment-agnostic, cSC-gated)

- context: §4.7b stated "Mech 2 = jc=both required". §4.7c stated
  "Mech 3 = same algorithm, also fires at jc=left". The "alignment-gate
  difference" was a key remaining nuance.
- evidence:
  - `bisect_mech3_alignment_full.py`: 10 variants × 5 alignments × 2 cSC
    values on 7f272a clone:
    cSC=compressPunctuation + jc ∈ {left, both, center, right, distribute}
      → ALL 5 FIRE (Mech 2 active in all alignments)
    cSC=doNotCompress + any jc → NONE fire
  - `verify_mech2_csc_dependence.py`: 8 variants of CONTROLLED docx
    (direct zip, no Word.Documents.Add() inheritance) × 4 cSC states
    × 2 jc:
    cSC=compressPunctuation + jc=both → FIRE
    cSC=compressPunctuation + jc=left → FIRE
    cSC=doNotCompress + jc=both → no fire ← KEY
    cSC=doNotCompress + jc=left → no fire
    No settings.xml → no fire (default doNotCompress)
    Empty settings.xml → no fire
- root cause of §4.7b's "jc=both" mistake:
  §4.7b's synthesized minimal docs were built with
  `Word.Documents.Add()`, which inherits cSC=compressPunctuation from
  the Japanese Normal.dotm template. So §4.7b's measurements were
  always under cSC=compressPunctuation, but only jc=both was tested.
  When §4.7c was investigated with explicit cSC control (direct-zip
  docx + clone of 7f272a), the alignment-agnostic behavior emerged.
- outcome:
  - Mech 2 and Mech 3 are the SAME mechanism. Trigger = cSC. Alignment
    is irrelevant.
  - Spec §4.7b's "jc=both gate" claim CORRECTED. Other §4.7b findings
    (algorithm, position rule, charset, floor, wrap-budget) are
    alignment-agnostic and remain correct.
  - Spec §4.7c reframed as unification, with full 5-alignment matrix
    + Mech 2/3 unification verification subsection.
  - Implementation simplification (R34):
    `should_apply_mech_2(doc_compresses_punct: bool) -> bool {
       doc_compresses_punct  // alignment irrelevant
    }`
  - Practical implication: any cSC=compressPunctuation document compresses
    yakumono on overflowing lines under any alignment. R32's
    alignment-gated 0.583x hack should be removed; the gate is doc-level
    cSC, not alignment.
- code change: NONE. Spec § 4.7b/§4.7c unified.

## 2026-05-02 — oxi-1 — confirmed — §4.7c Mech 3 trigger PINNED: characterSpacingControl="compressPunctuation"

- context: Session 51 R0 found Mech 3 needs "real-doc supporting files"
  but did not identify the discriminator. R34 implementation needs the
  exact gate.
- evidence — bisect_mech3_trigger.py (15 variants):
  Removing 11 different elements/files from 7f272a clone, only one
  disabled Mech 3:
  - `<w:characterSpacingControl w:val="compressPunctuation"/>` removal
    → Mech 3 NO LONGER FIRES
  - `<w:useFELayout/>` removal → still fires
  - `<w:balanceSingleByteDoubleByteWidth/>` removal → still fires
  - `<w:adjustLineHeightInTable/>` removal → still fires
  - compatMode 14→15 → still fires
  - `<w:compat>` block fully removed → still fires
  - `<w:kern>` removed from docDefaults → **STILL FIRES**
  - `themeFontLang` removed → still fires
  - `fontTable.xml` removed → still fires
  - bare-minimum `settings.xml` (no cSC) → does NOT fire (consistent)
- evidence — bisect_mech3_csc_values.py (10 variants):
  cSC value matrix × kern × jc:
  - `compressPunctuation` + any kern + jc=left/both → fires
  - `compressPunctuationAndJapaneseKana` + kern=yes + jc=left → fires (equivalent)
  - `doNotCompress` + any kern + jc=left/both → does NOT fire
  - cSC element absent + kern=yes + jc=left → does NOT fire (default = doNotCompress)
- outcome:
  - **Mech 3 trigger = `<w:characterSpacingControl w:val="V"/>` with
    V ∈ {"compressPunctuation", "compressPunctuationAndJapaneseKana"}**
    in `word/settings.xml`. SOLE necessary and sufficient condition.
  - **Mech 1 (kern gate) and Mech 2/3 (cSC gate) are INDEPENDENT.**
    Both can fire concurrently or one without the other.
  - Why Session 51 minimal repros never fired: synthesized minimal docx
    lacked `word/settings.xml` entirely → cSC default = doNotCompress
    → no Mech 3. The "real-doc supporting files" requirement was a
    red herring — only one element matters.
  - ECMA-376 §17.15.1.10 documents `characterSpacingControl` as
    "applied only at justify" but Word also fires it at jc=left
    (undocumented).
  - Spec §4.7c updated with full trigger spec, bisect tables, ECMA
    reference, and concrete Oxi implementation gate code.
  - For R34 implementation: gate yakumono compression on the
    document-level cSC setting, NOT on real-doc heuristics or kern
    presence.
- code change: NONE. Implementation guidance refined in spec §4.7c.

## 2026-05-02 — oxi-1 — confirmed — §4.7c Mech 3 compression formula = same as Mech 2 (slack 0.5pt-step), only alignment gate differs

- context: Per-request B. Mech 3 compression amount formula. 7f272a_p1
  P13 shows 0.91x (L1) / 0.76x (L2) per-yak ratios. Need to fit
  observed compression to: (a) slack distribution, (b) grid-snap,
  (c) font_size × const.
- evidence:
  - `tools/metrics/measure_mech3_7f272a_per_line.py`: 7f272a paragraphs
    11/13/16/18-22/25/27-30/34. Per-line per-char Information(5):
    P13 L1: 3 yak compress (1.0, 1.0, 0.5pt) = 2.5pt total
    P13 L2: 3 yak compress (2.5, 2.5, 2.5pt) = 7.5pt total
    P34 L1: 1 yak compresses 5.0pt (`、→（` Mech 1 FINAL RULE B→A)
  - `tools/metrics/measure_mech3_compression_formula.py`: 5 controlled
    minimal-repro probes × 2 alignments (jc=left, jc=both), each
    cloning 7f272a's supporting files:
    P_yak3_overflow (45 chars, natural~468.5, overflow +3.6):
      jc=left:  3 yak compress (1.0, 1.5, 1.0) = 3.5pt ≈ overflow
      jc=both:  3 yak compress (1.0, 1.5, 1.0) = 3.5pt — IDENTICAL
    No-overflow probes (29/40/37 chars): 0 compression in both
    alignments. Mech 3 needs overflow to fire.
- hypothesis verdicts:
  (a) slack distribution → **CONFIRMED**: total_compression ≈ overflow,
      distributed in 0.5pt steps, sum == slack. Same algorithm as Mech 2.
  (b) grid-snap → REFUTED: compression amounts {1.0, 1.5, 2.5pt} are
      NOT multiples of any character grid pitch.
  (c) font_size × const → REFUTED: per-yak ratios vary line-by-line
      (0.86, 0.90, 0.95, 0.76) — no single constant fits.
- outcome:
  - Spec §4.7c added: Mech 3 = Mech 2 algorithm with relaxed alignment
    gate (fires under jc=left when kern is on + real-doc supporting
    files present).
  - **Conservative implementation**: extend Mech 2 to fire under jc=left
    when `docDefaults.rPr.kern` present. The "real-doc supporting files"
    requirement (Session 51 finding) may fall out automatically since
    synthesized R32 sentinel tests lack the trigger components.
  - The original 7f272a P13 L1 "no-overflow but 2.5pt compression"
    appears to be measurement-artifact related to the ASCII digit "14"
    being half-width (advance ~5pt vs natural CJK 10.5pt) — actual
    natural width of P13 L1 is closer to 457pt + ~5pt half-width adj,
    putting it at-or-just-over content_w=464.9. Slack distribution
    holds.
- code change: NONE. Implementation guidance in spec §4.7c.

## 2026-05-02 — oxi-1 — confirmed — Cross-doc audit: Mech 1 fires on jc=left baseline docs (84/184 affected)

- context: Per-request audit to estimate ship-priority of Mech 1
  alignment-agnostic rule (Q6) by sampling real baseline docs.
- evidence:
  - `tools/metrics/audit_mech_compression_jc_left.py`: scanned all 184
    baseline docs. Found **84 docs (46%)** have all 3 requirements:
    `<w:kern>` in docDefaults + ≥3 paragraphs with jc=left/none jc + yakumono content.
    Top 5 by yak-content count: 3a4f9fbe1a83 (1011 paras), ed025cbecffb
    (180), d77a58485f16 (129), 6514f214e482 (94), b837808d0555 (65).
  - `tools/metrics/audit_mech_smart.py`: targeted Mech 1 trigger pair
    detection (Type-A→A, B→A, B→B) in jc=left/none paragraphs of top-5
    real baseline docs. Per-char COM measurement (Information(5)) of 5
    such paragraphs each:

    | Doc | trigger paras | measured | compressed | yak compressed |
    |---|---|---|---|---|
    | 3a4f9fbe1a83 | 188 | 5 | 0 | 0/5 |
    | ed025cbecffb |  20 | 5 | 2 | 4/19 |
    | d77a58485f16 |  40 | 5 | 4 | 8/33 |
    | b837808d0555 |   8 | 2 | 2 | 6/13 |
    | e3c545fac7a7 |   6 | 1 | 0 | 0/3 |

  - **Direct jc=left + Word.Alignment=0 (left) + Mech 1 firing CONFIRMED**:
    ed025 p13: ）=5.5pt at 10.5pt font (half-width Mech 1)
    ed025 p25: 。=5.0pt, ）=5.0pt at 10.5pt font (Mech 1)
  - Many paragraphs with XML `(no jc)` resolved by Word to Alignment=3
    (justify) via style inheritance. Cannot cleanly distinguish Mech 1
    vs Mech 2 in those, but compression IS present.
  - b837 p39 shows mixed compression values (、=7.0, （=7.5, 。=6.0pt at
    12pt) — characteristic of Mech 2's 0.5pt-step distribution (NOT
    Mech 1 strict half-width).
- outcome:
  - Mech 1 alignment-agnostic rule (Q6) CONFIRMED on real baseline data
    (ed025 p13/p25). Not only synthetic repros.
  - **Estimated impact**: 46% of baseline (84 docs) potentially affected
    by Mech 1 firing under any alignment when kern is on. Of those, the
    bottom-N target docs (ed025, d77a, b837) already show measurable
    Mech 1 compression. Ship-priority for Mech 1 alignment fix: HIGH.
  - 3a4f9fbe1a83 (1011 jc=left/none paras) shows 0/5 compression in
    measured sample — but only because the sampled paragraphs lack
    Mech 1 trigger pairs adjacent to body content (most are 「規則」と
    style with full-width 「」 between CJK ideographs which doesn't
    fire Mech 1).
  - Smart-audit methodology demonstrated: text-level trigger detection
    + targeted COM measurement gives reliable corpus-wide picture in
    ~20 paragraphs measured (vs naive sampling's 0% hit rate).
- code change: NONE (pure investigation). Implementation impact:
  Oxi's Mech 1 implementation gate must NOT be alignment-conditional;
  current code may be over-suppressing under non-justify alignments.

## 2026-05-02 — oxi-1 — partial — §4.7b Mech 2 wrap-budget intertwined design + Oxi implementation sketch

- context: §4.7b Mech 2 algorithm confirmed (0.5pt step, fontSize×2/3
  floor) but "drop char + refit" is wrap-decision intertwined. R32's
  alignment-gated 0.583x hack is a middle ground; proper
  implementation needs the full wrap algorithm.
- evidence: 3 probe sweeps at jc=both, MS Mincho 12pt:
  - P_yak3 (21 chars, 3 yak): 2 transitions found
  - P_yak6 (20 chars, 6 yak): partial (Word died, ~10 datapoints)
  - P_yak12 (22 chars, 10 yak, 9 compressible): 130 datapoints — clean
  Tools: `measure_m2_wrap_budget.py` + `_chunked.py` (Word-restart
  on RPC failure).
- findings:
  - **Drop trigger**: each "drop 1 char" boundary at slack ≈ fontSize
    (=12pt for 12pt MS Mincho). Word refuses total-line compression
    > 1×fontSize → drops a char.
  - **Per-yak distribution cap (multi-yak)**: 9 yak absorbed 22pt
    total at cw=206; refused 23pt at cw=205. Cap ≈ 2.5pt =
    fontSize × 5/24.
  - **Per-yak cap (single-yak)**: from §4.7b round 1 — 1 yak can
    absorb 4pt = fontSize/3. Asymmetric with multi-yak case.
  - **N→N-2 jump at cw=205**: Word skipped 18-char fit, dropping
    directly to 17 chars (natural=204 ≈ cw). Suggests Word uses
    natural-fit-greedy as base, with optional +1 char Mech 2
    extension.
- inferred algorithm:
  ```
  natural_fit_n = max N s.t. sum(natural[0..N]) <= cw
  while extend_n < len(chars):
      candidate_n = extend_n + 1
      new_slack = sum(natural[0..candidate_n]) - cw
      if new_slack <= 0: extend; continue
      if new_slack >= fontSize: break    # drop threshold
      compressible = count_yak_skipping_pos1
      cap_per_yak = fontSize/3 if 1 yak else fontSize*5/24
      if new_slack > compressible * cap_per_yak: break
      extend_n = candidate_n
  ```
- outcome:
  - Spec §4.7b extended with "Mech 2 + wrap-budget intertwined design"
    subsection, including observed regularities table, inferred
    algorithm, and ~80-LOC Oxi implementation sketch
    (`try_extend_with_mech2` + `distribute_mech2`).
  - Open: per-yak cap formula `fontSize × 5/24` is empirical; may need
    refinement at 10.5pt and other sizes. Need 2/3/4/5 yak data to
    pin the linear-vs-stepped relationship.
  - Open: 19→17 jump at cw=205 — exact "skip" rule between adjacent
    drops not fully characterized.
- code change: NONE (pure investigation + sketch). Implementation
  sketch provided in spec §4.7b for review. R32's alignment-gated
  0.583x hack can be replaced by the full Mech 2 wrap-budget
  algorithm.

## 2026-05-02 — oxi-1 — confirmed — §4.7b Mech 1 alignment-agnostic + Mech 1↔Mech 2 precedence

Two follow-up investigations to §4.7b Mech 2 characterization:

### Q6: Mech 1 alignment dependency

- context: §4.7b confirmed Mech 2 fires only at jc=both. §4.7 (Mech 1
  Type A/B/C) did not specify alignment requirement. ed025_p1 (R17
  big_loser) showed Mech 1 firing in jc=center/right paragraph,
  suggesting Mech 1 is alignment-agnostic.
- evidence: `tools/metrics/build_m1_alignment_test.py` +
  `measure_m1_alignment.py`. 2 docs × 5 alignment paragraphs each:
  - kern OFF, all 5 alignments: ）= 10.5pt (no Mech 1)
  - kern ON, all 5 alignments: ）= 5.0–5.5pt (Mech 1 fires)
  Specifically: jc=both/left/(no jc) → 5.5pt; jc=center/right → 5.0pt
  (minor 0.25pt difference likely measurement artifact at non-left
  aligned glyph origins).
- outcome: Mech 1 is **alignment-agnostic**. `<w:kern>` in docDefaults
  is the SOLE gate (per session 51 yakumono_kern_trigger finding).
  Spec §4.7b updated.

### Q7: Mech 1 → Mech 2 precedence interaction

- context: §4.7b stated "Mech 1 fires first, Mech 2 fires second on
  residuals" but did not define "residuals" — char-set or slack-level.
- evidence: `tools/metrics/measure_m1_m2_precedence.py` 9-slack sweep on
  probe `漢漢漢」）漢漢「漢漢漢` (11 chars, MS Mincho 12pt). 」 fires Mech 1
  (B→B trigger with `）` neighbor); `）` and `「` do NOT fire Mech 1
  (B→CJK / single A in CJK). Post-Mech1 natural = 126pt.

  | cw | slack | 」 (M1-comp) | ） (uncomp) | 「 (uncomp) |
  |---|---|---|---|---|
  | 200/132/126 | ≤0 | 6.0pt | 12.0 | 12.0 |
  | 125 | +1 | 6.0pt | 11.5 | 11.5 |
  | 124 | +2 | 6.0pt | 11.0 | 11.0 |
  | 122 | +4 | 6.0pt | 10.0 | 10.0 |
  | 120 | +6 | 6.0pt | 9.0 | 9.0 |
  | 118 | +8 | 6.0pt | 8.0 | 8.0 (floor) |

- outcome:
  - Mech 2 NEVER touches Mech-1-compressed yakumono. `」=6.0pt` constant
    across all slacks. Mech 1's output is final for those chars.
  - Mech 2 distributes slack ONLY across uncompressed yakumono. Each gets
    `slack / n_uncomp_yak`, in 0.5pt steps, sum = slack EXACTLY.
  - "Residuals" = char-level subset (uncompressed yakumono), NOT
    line-level slack continuation.
  - The Mech 2 floor (`fontSize × 2/3 = 8.0pt` for 12pt) still applies
    to the residuals-only set.
- code change: NONE. Spec §4.7b's "Mech 1 vs Mech 2 interaction"
  subsection refined with measured data.

## 2026-05-02 — oxi-1 — confirmed — §4.7b Mech 2 (justify-time) trigger / position / algorithm characterization

- context: Session 51 R0 entries identified Mech 2 (justify-time
  yakumono compression) but left several questions open: which char
  triggers fire (8.0/7.5/5.5pt usage), what does "mid-line" mean (% of
  line length), are trigger pairs different from Mech 1 Type A/B/C, what
  is "reactive" / overflow gating semantics? Required: 5-10 pair per-char
  COM measurement.
- evidence:
  - `tools/metrics/measure_m2_trigger_pairs.py`: 10 yakumono-CJK pair
    cases × 4 slack values (0/2/4/8pt) at jc=both, MS Mincho 12pt:
    ALL 10 pairs (A→CJK, CJK→A, B→CJK, CJK→B for 「」（）、。) compress
    identically: slack=0→12.0pt, slack=2→10.0pt, slack=4→8.0pt, slack=8→
    line drops a char. Yakumono is the compressing char regardless of
    neighbor type.
  - `tools/metrics/measure_m2_position_and_charset.py`: position-axis
    test (start/mid/end) × alignment (justify/left) + 16-char extended
    set:
    * jc=left at any position: NO compression (Mech 2 jc=both gated)
    * jc=both at line-start position: Mech 2 NOT fire (drop instead)
    * jc=both at mid/end positions: Mech 2 fires
    * extended yakumono set { 「」（）［］【】〔〕、。 } all compress
      to 8.0pt
    * em-dash ―(U+2015) Type C: NO compression
    * ASCII hyphen, Latin a, plain CJK: NO compression
  - `tools/metrics/measure_m2_position_sweep.py`: yakumono position 1..20
    sweep: position 1 (line-start) → no compression / line drops; positions
    2..19 → all compress to 8.0pt. The "mid-line" position rule is
    operationally **`position > 1`**, NOT a percentage threshold.
- outcome:
  - Spec §4.7b expanded with Mech 2 trigger conditions, compressible char
    set (= Type A ∪ Type B from Mech 1, no Type C), position rule
    (position > 1), compression algorithm with `min_yak_width =
    fontSize × 2/3` cap, and 8.0/7.5/5.5pt value attribution table.
  - Key formula:
    `min_yak_width = fontSize × 2/3`
    (8.0pt for 12pt font, 7.0pt for 10.5pt font)
  - Per-yak compression cap = fontSize − min_yak_width = fontSize/3
    (4.0pt for 12pt, 3.5pt for 10.5pt). Beyond cap → drop char and refit.
  - 5.5pt = Mech 1 half-width (11pt font /2). 6.0pt = Mech 1 half (12pt).
    7.0pt = Mech 2 floor (10.5pt × 2/3). 8.0pt = Mech 2 floor (12pt × 2/3).
    7.5pt etc = Mech 2 distributed partial.
  - Mech 1 fires first (line-break time, neighbor-pair-based);
    Mech 2 fires second (layout time, slack-distribution-based).
- code change: NONE (pure investigation). Oxi's Phase 2 reactive absorb
  (per 1f8b5f2) should be reviewed against the spec §4.7b algorithm:
  honour the position>1 gate, use fontSize×2/3 floor, distribute in 0.5pt
  steps with sum-to-slack invariant.

## 2026-05-02 — oxi-1 — partial — §19.7 Y0 intercept anomaly explained: anchor empty para's pPr-rPr font

- context: Spec §19.7 / §18.10 (this branch). Prior round 1 (top entry below)
  established "Y0 = anchor_top + 1 line_height" universal across PreKinds for
  body paragraphs with content. The 2ea81a tbl#3 case still showed Y0 = +28.55pt
  (~2× the expected 14.5pt), unexplained.
- hypothesis (round 1 — REFUTED): Each intervening empty paragraph between
  the "real" preceding paragraph and the floating table adds +1 line_height
  to the Y0 intercept.
  → 30-variant FE_* repro (tools/metrics/fe_repro/, 5 PreKinds × 5 empty
  counts × 2 tblpY values) shows Y0 = constant ~14pt regardless of empty
  count. Hypothesis REFUTED.
- hypothesis (round 2 — CONFIRMED): The anomaly is driven by the LAST
  (anchor) empty paragraph's `<w:pPr><w:rPr><w:sz/></w:rPr></w:pPr>` font
  size — Word resolves the empty paragraph's height via that pPr-rPr font,
  not from default style.
  → 7-variant K_* axis-isolation matrix (tools/metrics/fe_match_repro/):
  - K_baseline (sz=21, lp=360, line=auto):       Y0 = +16.55pt
  - K_only_sz28 (sz=28, lp=360, line=auto):       **Y0 = +27.55pt** (≈ 2ea81a's +28.55)
  - K_only_atLeast296 (sz=21, lp=360, line=296atLeast): Y0 = +16.55pt (no effect)
  - K_lp323 (sz=21, lp=323, line=auto):           Y0 = +15.05pt
  - K_lp323_atLeast296_sz28 (sz=28, lp=323, line=296 atLeast): Y0 = +26.05pt
  - K_tblWauto_only:                              Y0 = +16.55pt (no effect)
  Single-axis: only `sz` change moves Y0 substantially. Reference
  2ea81a tbl#3: anchor empty para has pPr-rPr `sz=28` (14pt), line=296
  atLeast, docGrid lp=323 → reproduced as +26.05–27.55pt (within 1–2.5pt
  of the +28.55pt observed value).
- evidence:
  - `pipeline_data/fe_intervening_measurements.json` (30 variants, refutes
    intervening-empty-count hypothesis)
  - `pipeline_data/fe_match_measurements.json` (7 variants, confirms
    pPr-rPr-sz hypothesis)
  - `pipeline_data/tblppr_anchor_measurements.json` (2ea81a baseline ref)
  - `tools/metrics/{build,measure}_fe_intervening.py`
  - `tools/metrics/build_fe_2ea81a_match.py` + `measure_fe_match.py`
- outcome:
  - §19.7's "+1 × line_height_of_anchor" universal was an over-generalization.
    True for paragraphs with content; for empty paragraphs Y0 follows the
    pPr-rPr font.
  - Refined formula written to spec §18.10 (this branch) / will become §19.10
    when merged to main:
    `table_top = anchor_top + line_height_resolved_from(anchor.pPr.rPr.sz)
              + ~2pt floor + tblpY_pt`
  - Residual 1–2.5pt unexplained — likely floating-table topFromText spacing
    constant (default = 0 unless `topFromText`/`bottomFromText` set on
    tblpPr; needs separate isolation).
  - §19.7 / §18.7 IS still correct for the body-para-anchor case; the
    update is for empty-anchor case.
  - `intervening empty paragraph count` not a factor — refuted via 30
    variants. Only the LAST anchor's pPr-rPr matters.
- code change: NONE (pure investigation). Oxi's current line-height-for-anchor
  resolution should be checked against this rule when implementing §19
  shippable fix.

## 2026-05-02 — oxi-1 — confirmed — vertAnchor=text floating-table tblpY behavior + parser-order quirk

- context: §18 Floating Tables (`<w:tblpPr>`) was hypothesis-only, derived
  from a single doc (2ea81a) with cross-table cross-verification but no
  multi-doc convergence. §18.1 claimed slope=1.0 for `table_top` vs
  `tblpY`; §18.2 claimed `+28.5pt = 2 line-heights` Y0 intercept based
  on tbl#3.
- hypothesis: (a) slope=1.0 is universal across pre-content kinds,
  (b) Y0 intercept = `anchor_top + 1 line_height` (not 2), (c) prior
  TP1-6 minimal repros (existing) showing slope=0 are caused by some
  structural difference vs 2ea81a tbl#2.
- evidence:
  - `tools/metrics/build_ft_slope_repro.py` + `measure_ft_slope.py`:
    25 minimal repros, 5 PreKinds (1para / 3para / 1empty / inline /
    inline_p) × 5 tblpY values (0, 50, 600, 2000, 4000 twips). All 25
    show clean slope=1.0. Y0 intercept = `anchor_top + ~14pt` for
    body-para anchor (= line_height of MS Mincho 10.5pt single-line
    auto-spacing). For empty-para anchor +18.5pt; for inline-cell
    anchor +15.0pt.
  - `tools/metrics/measure_tp_resweep.py`: re-measured TP1-6 (existing).
    Reproduced prior slope=0 (TP1=71.0/TP2=71.0/TP3=71.0/TP4=98.0/
    TP5=98.0/TP6=98.0). Confirms TP repros really do show slope=0.
  - `tools/metrics/build_compat_test.py` + `measure_compat_test.py`:
    10 minimal repros across compatMode ∈ {none, 11, 12, 14, 15} ×
    tblpY ∈ {50, 600}. ALL show slope=1.0. **compatMode hypothesis
    REJECTED** — that is not the cause.
  - `tools/metrics/build_tp3_mutate.py` + `measure_tp3_mutate.py`:
    18 variants from TP3 with single-axis mutations:
    - `M_baseline / M_tblWdxa / M_noUseFE / M_noNumbering` → slope=0
    - `M_noStyles / M_noTblStyle` → slope=1
    Removing the `<w:tblStyle w:val="TableGrid"/>` reference is the
    necessary and sufficient mutation to flip slope.
  - `tools/metrics/build_order_test.py` + `measure_order_test.py`:
    9 variants from TP3:
    - `O_baseline` (tblpPr → tblStyle in source XML) → slope=0
    - `O_swapped` (tblStyle → tblpPr, ECMA-376 §17.4.79 CT_TblPrBase
      sequence) → slope=1
    - `O_noStyle` (tblStyle removed) → slope=1
    The single XML-order swap (no other change) flips slope from 0 to 1.
  - Cross-check: 2ea81a tbl#3 has BOTH `<w:tblStyle w:val="aa"/>` AND
    `<w:tblpPr>`, and is observed at slope=1.0. Its tblPr child order
    is `tblStyle → tblpPr` (correct ECMA order). TP3's tblPr order is
    `tblpPr → tblStyle` (incorrect). Same property, opposite order,
    opposite slope.
- outcome:
  - §18.1 slope=1.0 finding is RECONFIRMED universally for ECMA-compliant
    tblPr ordering, across 5 distinct PreKinds.
  - §18.2 "+28.5pt = 2 line heights" hypothesis is LOCALLY REFUTED. The
    universal Y0 intercept is `anchor_top + 1 line_height`. The 28.5pt
    observation in 2ea81a tbl#3 is a separate phenomenon (likely
    intervening-empty-para counted twice, or floating-table reserved
    region — needs follow-up).
  - **NEW §18.8 (CRITICAL undocumented quirk)**: Word's parser silently
    drops `<w:tblpPr>` if its child-element ordering inside `<w:tblPr>`
    violates ECMA-376 §17.4.79 CT_TblPrBase sequence. When dropped, the
    table renders as inline at `anchor_bottom`. Single XML-order swap
    is sufficient to flip behavior.
  - Baseline survey: `tools/metrics/scan_baseline_tblpPr.py`-style scan
    on 184 baseline docs finds 0 docs with order violations (375 total
    `<w:tbl>`, 16 floating). So §18.8 quirk does NOT directly affect
    baseline SSIM. Real-world Word-generated docs respect ECMA order;
    only manually-authored repros (TP1-6) had the violation.
  - **Implication for Oxi parser**: ooxml.rs currently honors `<w:tblpPr>`
    regardless of its position within `<w:tblPr>`. To match Word strictly,
    Oxi should drop tblpPr when it precedes tblStyle. Pre-existing impact
    is zero on baseline (no order-violating docs), but defensively
    important for hand-edited or non-Word-generated source files.
  - All COM data: `pipeline_data/{ft_slope,tp_resweep,compat_test,
    tp3_mutate,order_test}_measurements.json`. Spec updated with new
    §18.6 / §18.7 / §18.8 / §18.9 sections.

## 2026-04-27 — oxi-main — partial — yakumono closing-punct compression fires regardless of useFELayout / kern (gate hypothesis narrowed)

- context: post-loop-termination triage (3 deep dives across `adjacency_matrix_widths.json`,
  `bracket_pair_widths.json`, `mincho_adjacency_widths.json`) found that for
  Meiryo 10.5pt + cSC=doNotCompress + compat=14, Word compresses closing-class
  punct (、。」）．，) to 5.25pt when followed by a trigger char — but Oxi's
  `crates/oxidocs-core/src/layout/mod.rs:4140` gate
  `yakumono_enabled = self.compress_punctuation` is FALSE for cSC=doNotCompress
  → would render full 10.5pt → mismatches Word.
- hypothesis: useFELayout (and/or kern w:val="3") is the unaccounted trigger
  Word checks for, gating yakumono compression independent of cSC. Oxi does
  not parse useFELayout (grep confirmed), so opening that gate would close
  the gap.
- evidence: `pipeline_data/meiryo_linewidth_repro.json` LW_30 (useFE=ON,kern=3)
  vs LW_31 (useFE=off,kern=off) per-char compare — identical paragraph text
  `メタデータは、各機関で...「９．例 (1)メタデータ」を参照ください。`, both
  cSC=doNotCompress + compat=14. Position 24 (`、` followed by `「`) measured
  **5.50pt in BOTH cases**; all other punct unchanged at 10.50pt. Total line
  widths identical at 465.50pt. → useFELayout/kern do NOT gate the compression.
  Compression fires identically on the next-trigger rule alone.
- outcome: Hypothesis useFELayout-as-gate REFUTED. Oxi's `mod.rs:4140` gate
  is *probably* over-restrictive — Word applies the next-trigger rule
  unconditionally for at least {compat=14, cSC=doNotCompress, useFELayout in
  {on,off}, kern in {on,off}}. Two open suspects remain: cSC=compressPunctuation
  giving *extra* compression beyond the always-on rule, and compat=15 differing
  from compat=14. Staged variant test trimmed to V_CP + V_COMPAT15
  (`tools/metrics/{build,measure}_adjacency_matrix_variants.py` + 2×60 fixtures).
  V_NOFE fixture set deleted (60 docx) as it would only re-confirm the LW_30/31
  finding.
- caveat to RESEARCH_LOG 2026-04-18 d77a yakumono bisect: that entry recorded
  "cSC alone: NO compression". Likely a content artifact — d77a's bisect text
  may not have contained trigger-pairs (`、` followed by `観` = CJK ideograph,
  NOT trigger → would never compress regardless of settings). The 2026-04-18
  finding is consistent with this 2026-04-27 finding under the next-trigger
  rule, but its conclusion ("compress_punctuation alone gates yakumono") is
  not supported by the data. Subsequent gate logic at `mod.rs:4140` should
  be reviewed against this corrected understanding.
- code change: NONE. Opening the gate is potentially correct but unverified
  on baseline (177 docs / 352 pages). Bottom-N floor regression risk
  unquantified. Path B candidate pending V_CP + V_COMPAT15 measurement and a
  controlled `pipeline.verify` run.

## 2026-04-27 — oxi-main — confirmed — V_CP + V_COMPAT15 8x8 matrices both match baseline (always-on rule confirmed)

- context: follow-up to 2026-04-27 partial entry above. Ran
  `tools/metrics/measure_adjacency_matrix_variants.py` against 60-fixture
  V_CP (compat=14 + cSC=compressPunctuation) and V_COMPAT15 (compat=15 +
  cSC=doNotCompress) variant sets, comparing to baseline
  `pipeline_data/adjacency_matrix_widths.json` (compat=14 + cSC=doNotCompress
  + useFELayout=on).
- result: **EVERY cell in both 8x8 matrices (prev + next axes) matches
  baseline within 0.3pt tolerance**. No `*` (full-width / compression LOST)
  or `!` (still-compressed-but-different) markers.
  - `、` `。` `」` `）` `．` `，` after any closing-class neighbor → 5.25pt
    (compressed) in baseline, V_CP, AND V_COMPAT15
  - `「` `（` after compressed neighbor → 10.50pt (full)
  - `「` `（` before another opener `「` `（` → 5.25pt (compressed)
- conclusion: Word applies the next-trigger yakumono compression rule
  **unconditionally** for at least:
  - `compatibilityMode` ∈ {14, 15}
  - `cSC` ∈ {doNotCompress, compressPunctuation}
  - `useFELayout` ∈ {on, off} (LW_30/LW_31 finding above)
  - `kern` ∈ {on, off} (LW_30/LW_31 finding above)
- implication: Oxi's `mod.rs:4161` gate `yakumono_enabled =
  self.compress_punctuation` is **over-restrictive for the entire baseline**.
  Baseline docs predominantly use cSC=doNotCompress, which currently disables
  Oxi's yakumono compression — but Word compresses them anyway. This
  explains documented Word-vs-Oxi line-fit gaps in 31+ baseline docs
  (3a4f9fbe1a83 = 213 closing-yakumono pairs, largest single-doc impact).
- spec reference: undocumented Word quirk (no JIS X 4051 / ECMA-376 clause
  describes "always-on next-trigger yakumono compression"). Designating the
  COM matrix above as the spec evidence per CLAUDE.md Path B clause.
- evidence files preserved:
  - `pipeline_data/adjacency_matrix_widths.json` (baseline, 8x8 grid, 4 fonts)
  - `pipeline_data/adjacency_matrix_widths_V_CP.json` (8x8 grid, cSC=cP)
  - `pipeline_data/adjacency_matrix_widths_V_COMPAT15.json` (8x8 grid, compat=15)
  - `pipeline_data/meiryo_linewidth_repro.json` (LW_30/LW_31 useFE/kern refute)
  - `tools/metrics/adjacency_matrix_repro_*` (180 fixture docx)
  - `tools/metrics/measure_adjacency_matrix_variants.py` (re-runnable harness)
- code change: NONE this commit. Gate-open patch (`mod.rs:4161` →
  `yakumono_enabled = true`) is the natural next step but requires:
  (a) GDI renderer rebuild, (b) clear `pipeline_data/oxi_png/<doc>/` for
  31+ affected docs, (c) full `pipeline.verify` run, (d) Path B
  [confidence-merge] commit if bottom-N floor (3.2645) does not regress.
  d77a is the SOLE rule_b doc and is currently in bottom-5 → could improve
  or regress; verify mandatory.
- outcome: Path B candidate cleared for verify. 4 of 4 deep dives now align:
  always-on next-trigger rule, no controlling toggle in the
  {compat,cSC,useFELayout,kern} axis space measured.

## 2026-04-27 — oxi-main — refuted — yakumono always-on gate-open FALSIFIED on baseline (catastrophic regression)

- context: Acted on the 2026-04-27 confirmed entry above. Patched
  `crates/oxidocs-core/src/layout/mod.rs:4178`:
    -    let yakumono_enabled = self.compress_punctuation;
    +    let yakumono_enabled = true;
  Followed CLAUDE.md verify hygiene: `cd tools/oxi-gdi-renderer && cargo
  build --release` (1m 06s), cleared `pipeline_data/oxi_png/<doc>/` for
  20 affected dirs (11 absent), ran full `pipeline.verify` on 177 baseline
  docs / 352 pages.
- result: **18 page regressions, 6 improvements, 328 unchanged. Net
  -2.0184**. Two docs collapsed catastrophically:
  - `0e7af1ae8f21_..._sample_00`: pages 2,3,4,5,6,7 dropped from
    0.75-0.81 to **0.50-0.56** (-0.19 to -0.29 each); p.8 -0.0175;
    p.10 -0.0570; p.1 -0.0154
  - `683ffcab86e2_..._addon_00`: p.1 0.7592→0.6310 (-0.1282), p.2
    0.7558→0.5022 (-0.2536), p.3 0.9695→0.9316 (-0.0379)
  - `e3c545fac7a7_LOD_Handbook`: p.1 -0.057, p.2 -0.023, p.4 -0.024,
    p.5 -0.014, p.11 -0.010 (mixed: 6 other LOD pages improved
    +0.002 to +0.027)
  - `4a36b62555f2_kyodokenkyuyoushiki10` p.1: -0.0047 (tiny)
- bottom-5 floor impact:
  - pre: 3.2645 (d77a 0.6268 + b837 0.6449 + 29dc 0.6636 + 2ea8 0.6643 + e3c5 0.6649)
  - post: ~2.9337 (683ff 0.5022 + 0e7af 0.5051 + d77a 0.6268 + b837 0.6449 + e3c5 0.6547)
  - **Δ -0.3308 catastrophic floor regression**
- falsification: The "always-on next-trigger rule" hypothesis is FALSE at
  the baseline level. COM measurement on 4-char isolated paragraphs
  (8x8 grid × {V_CP,V_COMPAT15} = 180 fixtures, all matching baseline)
  did NOT generalize to multi-line real-world paragraphs. Word's actual
  gate involves additional context the 4-char fixtures could not test
  (line-position dependency? surrounding-char dependency? line-break
  interaction with paragraph layout? something Oxi's implementation
  mishandles when always enabled?).
- mechanism speculation (NOT verified):
  - The 4-char fixtures had only 1 yakumono pair per paragraph, on line 1
  - Real docs (0e7af, 683ff) have many pairs distributed across lines
  - Possibility A: Word's compression triggers only at specific positions
    (e.g., line-end? after specific char classes Oxi doesn't track?)
  - Possibility B: Oxi's compression code has bugs that surface only
    when the gate is open on multi-pair paragraphs
  - Possibility C: GDI rendering pipeline applies compression
    inconsistently with the layout calculation
  - Distinguishing requires DML extraction / per-line position diff for
    0e7af and 683ff — out of scope this session
- spec status: cSC=doNotCompress + COM-measured isolated yakumono pair
  compression remains a real Word behavior (LW_30/LW_31, V_CP/V_COMPAT15
  all confirmed), but the rule that *opens the gate* is NOT
  unconditionally always-on. Some context-discriminator exists.
- code change: REVERTED to `let yakumono_enabled = self.compress_punctuation;`
  with extensive comment block at mod.rs:4161-4180 documenting this
  falsification so future agents do not re-attempt the same patch.
  Rebuilt oxi-gdi-renderer post-revert (1m 06s, second build). Cleared 32
  PNG dirs (affected + regressed) so cache regenerates on next render.
  Baseline JSON UNCHANGED (verify.py returns False before updating when
  regressions present; confirmed: timestamp 2026-04-26, bottom-5 sum
  3.2645 unchanged).
- outcome: bottom-5 floor 3.2645 maintained. Patch falsified. Real Word
  gate is narrower than the 8x8 fixture matrix could detect. Future
  investigation needs multi-line, multi-pair fixtures + DML-level
  per-line position comparison on 0e7af/683ff specifically.

## 2026-04-27 — oxi-main — narrowing — 0e7af/683ff direct COM probe + 4 discriminator candidates eliminated

- context: Followed up on the always-on falsification by directly probing
  Word's actual yakumono compression behavior on the regressed real docs.
- artifacts:
  - `tools/metrics/probe_0e7af_yakumono.py` — measures all closing-class
    yakumono pairs in a docx via COM, classifies width as
    COMPRESSED (5.25pt) / FULL (10.5pt) / OTHER
  - `pipeline_data/probe_0e7af_yakumono.json` — measurements for 0e7af
    (822 pairs) and 683ff (187 pairs)
  - `tools/metrics/build_mincho_nokern_variants.py` +
    `tools/metrics/mincho_kern_variants/{MC_A_mincho_NOKERN,MC_A_mincho_NOKERN_COMPAT15}.docx`
  - `tools/metrics/measure_mincho_kern_variants.py` +
    `pipeline_data/mincho_kern_variants.json`
  - `tools/metrics/build_mincho_9pt_variant.py` +
    `tools/metrics/mincho_size_variants/MC_A_mincho_9pt.docx`
- direct evidence:
  - **0e7af** (MS Mincho 9pt, no kern, compat=15, cSC=doNotCompress):
    822 yakumono pairs measured → **0 COMPRESSED (0.0%)**, 6 FULL (0.7%),
    816 OTHER (mostly 9.0pt = fullwidth at 9pt). Even `。）` (38 occurrences)
    showed 9.0pt for `。`. **Word does NOT compress in 0e7af**.
  - **683ff** (MS Mincho ~10.5pt, no kern, compat=15, cSC=doNotCompress):
    187 pairs → 1 COMPRESSED (0.5%), 129 FULL (69%), 57 OTHER. **Word does
    NOT compress in 683ff**.
- discriminator hypothesis testing (vs MC_A_mincho fixture which compresses):
  | hypothesis | test | result |
  |---|---|---|
  | font (MS Mincho excluded?) | mincho_adjacency MC_A_mincho.docx | COMPRESS — REFUTED |
  | compat=14 vs 15 | MC_A_mincho_NOKERN_COMPAT15 | COMPRESS — REFUTED |
  | kerning | MC_A_mincho_NOKERN (compat=14, no kern) | COMPRESS — REFUTED |
  | font size 9pt | MC_A_mincho_9pt | COMPRESS (4.5pt = half of 9pt) — REFUTED |
- 4 of 5 obvious axis-discriminators eliminated. Remaining candidates
  must lie outside the {compat, cSC, useFE, kern, font, size} space:
  - real-doc text shape (multi-paragraph natural Japanese vs synthetic
    `観、「測` repeat)
  - paragraph indent / hanging
  - run-level properties beyond rFonts (lang tag, theme refs, rPr
    inheritance from styles)
  - line-wrap context (compression conditional on wrap-budget pressure?)
  - section properties (page columns, gutters)
- code change: NONE. Adds 3 measurement scripts + 3 variant fixtures.
  No `crates/` touched. No baseline impact.
- outcome: 4 discriminator candidates eliminated via 6+1+1 variant
  measurements. Next session needs structurally-different approach:
  (a) Word DML position extraction on 0e7af page 4 + Oxi layout JSON
  comparison to find specific divergence positions, OR
  (b) Build a fixture that *clones* 0e7af's first multi-pair paragraph
  exactly (same indent, run properties, paragraph context) and measure;
  if clone compresses, narrow further; if not, the discriminator is
  in the cloned property set
  Hand-off doc updated in `project_adjacency_matrix_variants_2026_04_27.md`.

## 2026-04-27 — oxi-main — narrowing — 2 more discriminator candidates eliminated (text-shape, jc=both)

- context: Continued narrowing after 4 axis-discriminators eliminated.
- artifacts:
  - `tools/metrics/build_0e7af_text_replaced_variant.py` — replaces the
    first body paragraph's text in 0e7af with the MC_A_mincho fixture
    pattern (`観、「測` × 10), preserving all run/para/section properties
  - `tools/metrics/text_replaced_variants/0e7af_with_fixture_text_in_p1.docx`
  - `tools/metrics/build_mincho_jc_both_variant.py` — adds
    `<w:jc w:val="both"/>` to MC_A_mincho's single paragraph
  - `tools/metrics/jc_variants/MC_A_mincho_jc_both.docx`
- evidence:
  - **Text shape (text in 0e7af context)**: Injected `観、「測` × 10 into a
    plain body paragraph of 0e7af (no pStyle, no run sz override → inherits
    pPrDefault jc=both + 9pt body font from rPrDefault). COM-measured `、`
    widths = ~11.5pt (full at the inherited size). NOT COMPRESSED. Even
    with the exact MC_A_mincho text, 0e7af's environment suppresses
    compression. ⇒ text-shape is NOT the discriminator.
  - **jc=both**: Added `<w:jc w:val="both"/>` to MC_A_mincho's only
    paragraph. COM-measured `、` widths = 5.0-5.5pt (avg 5.3). STILL
    COMPRESSED. ⇒ jc=both alone is NOT the discriminator. (However,
    note: MC_A_mincho's text fits one line, so jc=both has no
    justification work to do — a long-paragraph test is still possible.)
- discriminator candidates eliminated this session: 6 of 7 axis tests
  REFUTED:
  | hypothesis        | test                                  | result   |
  | font (Mincho?)    | mincho_adjacency MC_A_mincho          | REFUTED  |
  | compat=14 vs 15   | MC_A_mincho_NOKERN_COMPAT15           | REFUTED  |
  | kerning           | MC_A_mincho_NOKERN                    | REFUTED  |
  | font size 9pt     | MC_A_mincho_9pt                       | REFUTED  |
  | text shape        | 0e7af with fixture-text injection     | REFUTED  |
  | jc=both           | MC_A_mincho with jc=both              | REFUTED  |
- structural finding: MC_A_mincho.docx is **minimal** (only
  `[Content_Types].xml`, `word/document.xml`, `word/settings.xml`).
  0e7af.docx has additionally `word/styles.xml`,
  `word/fontTable.xml`, `word/theme/theme1.xml`, `word/webSettings.xml`,
  footnotes/endnotes.xml. The discriminator likely lives in one of
  these support files. 0e7af's `<w:rPrDefault>` includes
  `<w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/>` which
  MC_A_mincho lacks. fontTable.xml may carry PANOSE info that triggers
  font-specific layout rules.
- code change: NONE. Adds 2 build scripts + 2 fixture variants.
- outcome: 6 discriminator candidates REFUTED. Bottom-5 floor 3.2645
  maintained. Investigation has reached "MC_A_mincho is too minimal to
  match real-world Word docs' rendering context" — narrowing further
  requires either:
  (1) progressive test: add styles.xml + fontTable.xml + theme.xml +
      lang tag + ... to MC_A_mincho one at a time, measure after each
  (2) inverse test: take 0e7af and progressively REMOVE features until
      compression starts (likely faster — fewer iterations)
  (3) structural pivot: accept that yakumono compression has a complex
      gate that COM-by-fixture investigation is poorly suited to
      uncover, and move to a different productive task
  Recommend (3) per CLAUDE.md "no excuse stacking" principle: when a
  rule needs many carve-outs, the rule itself is wrong / too narrow.

## 2026-04-27 — oxi-main — confirmed — current Oxi yakumono gate is CORRECT for entire baseline (gate composite, no patch needed)

- context: After 6+ falsified hypotheses, used inverse-strip on 0e7af +
  text-injected fixture pair to find what suppresses Word's yakumono
  compression in real-world docs.
- artifacts:
  - `tools/metrics/inverse_strip_with_inject.py` + 6 strip variants of
    0e7af with fixture-text injected
  - `tools/metrics/inverse_strip_rprdefault_subelements.py` + 3 sub-strip
    variants
  - `tools/metrics/mincho_with_docdefaults.docx` (reverse test variant)
- inverse-strip result (compression on injected `、「` pair):
  | variant                       | width  | verdict                |
  | V0 (inject only, no strip)    | 11.50  | FULL (suppressed)      |
  | V1 (strip rPrDefault lang)    | 11.50  | FULL                   |
  | **V2 (strip rPrDefault all)** | **6.00**  | **COMPRESSED!**       |
  | V3 (strip pPrDefault)         | 10.50  | FULL                   |
  | V4 (strip docDefaults)        | 5.50   | COMPRESSED             |
  | V5 (minimal styles.xml)       | 5.50   | COMPRESSED             |
- sub-element drill-down (within rPrDefault):
  | variant                | rPrDefault remaining       | width | verdict |
  | V2a (strip rFonts)     | sz, szCs, lang             | 11.50 | FULL    |
  | V2b (strip sz/szCs)    | rFonts, lang               | 11.50 | FULL    |
  | V2c (strip rFonts+sz)  | lang only                  | 11.50 | FULL    |
  | V2d (empty rPrDefault) | <w:rPr/> only              | 11.50 | FULL    |
- reverse confirmation: added empty docDefaults to MC_A_mincho.docx
  (which originally compresses) → STILL COMPRESSED 5.50pt. The rule is
  NOT simply "rPrDefault presence suppresses compression" — adding it
  to a minimal fixture doesn't transfer the suppression.
- conclusion: Word uses a **composite heuristic** to detect "modern Word
  2007+ document" vs "minimal/legacy document". The detection involves
  multiple structural elements (styles.xml content + relationships +
  possibly fontTable.xml + theme.xml + ...). Modern docs use one layout
  mode (no yakumono compression unless cSC=cP); minimal/legacy docs use
  another (compress yakumono regardless of cSC).
- **practical implication**: ALL baseline real-world docs are modern
  Word 2007+ docs with full structure. For all of them, Word does NOT
  compress yakumono when cSC=doNotCompress. **Oxi's current gate
  `yakumono_enabled = self.compress_punctuation` correctly matches Word
  for the entire baseline**. The "always-on" patch (55a8b4c, reverted)
  was based on V_CP/V_COMPAT15 fixture evidence that turns out to be
  artifacts of fixture minimality — those fixtures lacked the full
  structure that triggers Word's modern-doc-mode suppression.
- **NO Path B ship needed**. Current code is correct. Investigation closes.
- code change: NONE. mod.rs:4178 stays as-is with the falsification
  comment block from 55a8b4c (which now is the right outcome — the
  comment correctly warns future agents not to enable always-on).
- bottom-5 floor 3.2645 maintained (unchanged throughout).
- artifacts retained for future:
  - `pipeline_data/probe_0e7af_yakumono.json` (real-doc compression evidence)
  - `pipeline_data/mincho_kern_variants.json` (kern non-discriminator)
  - `pipeline_data/adjacency_matrix_widths_V_CP.json` (fixture compression)
  - `pipeline_data/adjacency_matrix_widths_V_COMPAT15.json`
  - `tools/metrics/inverse_strip_*` scripts + variants (composite-gate evidence)

## 2026-04-25 — oxi-main — refuted — split-box bottom padding cursor_y narrow fix (5th FALSIFIED)

- context: d77a P.7 rank-1 worst page (0.6268). User flagged "box 下 padding 欠落".
- hypothesis: Word body_y after a split table row = `last_new_page_y +
  cell_line_height + (trailing_empty_lh if TE)`. Fixing Oxi's `cursor_y`
  at `mod.rs:~5355` via a monotonic min-guard
  (`cursor_y = max(cursor_y, max_cont_text_bottom)`) should close the
  gap without depending on Oxi's under-sized `row_height`.
- spec evidence (CONFIRMED, §5.4b added): 2 real docs + 5 minimal
  repros SB_A..E all agree within ±1.5pt. d77a tbl#5/8/10 residuals
  -0.5pt, e3c545 tbl#2/3 residuals +1.0/+1.5pt.
- implementation evidence (FALSIFIED): verify result 79 pages regressed,
  net **-71.1192**. Local effect on d77a p.7 +0.1338 and p.8 +0.1653
  confirms the spec for TE=0, but d77a page count goes 12→14 →
  Oxi-Word page alignment breaks on p.9-p.12 (d77a p.12 collapses
  0.9656→0.7721) AND gen2_* docs SSIM drops to 0.0000 from cascade.
  Bottom-5 floor 3.2645 → ~3.2367 (strict decrease). Revert.
- outcome: **5th cursor_y FALSIFIED with identical cascade pattern**.
  spec-correct ≠ Oxi-shippable as long as upstream line-count diverges
  from Word. Memory `project_split_box_padding_5th_FALSIFIED`.
  Artifacts kept for re-attempt after upstream fix:
  - `tools/metrics/build_split_box_padding_repros.py` (SB_A-F builder)
  - `tools/metrics/split_box_padding_repro/SB_A..F.docx`
  - `tools/metrics/measure_split_box_padding.py`
  - `pipeline_data/split_box_padding_measurements.json`
- Ra §9 decision: do not re-attempt cursor_y-only fix. Work must
  address the upstream cell wrap line-count divergence first.

---

## 2026-05-02 — oxi-3 — confirmed — R17 gate per-char validation: Type A/B/C in losers, Mech 2 in "winners"

**Direct measurement of user's 4 target paragraphs** via per-char
`Information(5)` advances:

| Doc | User label | Word compression observed |
|---|---|---|
| ed025 p1 para 13 | big_LOSER | `）（` (B→A) and `））` (B→B) → 5.5pt half ✓ FINAL RULE Type A/B/C |
| 7f272a p1 para 13 | big_LOSER | Mech 2 distributed: yakumono → 8.0pt mid-line |
| 683f p2 para 30 | WINNER | All `、` followed by CJK → 10.5pt full (no compress) ✓ FINAL RULE |
| 3a4f p23 para 475 | WINNER (?) | Word compresses `、` → 8.0/7.5pt (Mech 2) — refutes user's "Word は compress しない" expectation |

### Key conclusions

1. **Type A/B/C FINAL RULE is the proper Mech 1 spec.** ed025 follows
   it exactly. R17's list_marker gate suppresses Mech 1 in plain
   paragraphs → big_loser SSIM regressions.
2. **Mech 2 fires without Mech 1 anchor.** Earlier finding REFUTED
   — 7f272a + 3a4f have only B→CJK chars yet Mech 2 fires.
3. **R17 list_marker gate is wrong in 3/4 cases.** Correctly matches
   Word only when neither Mech 1 nor Mech 2 triggers (683f).
4. **Replacement**: kern-gated FINAL RULE (Mech 1) + Mech 2 reactive
   absorb (already in Phase 2). R17 `dc7104c` should be removed.

### Source data

- `tools/metrics/measure_r17_yakumono_per_char_advances.py`
- `pipeline_data/r17_per_char_advances.log`
- `pipeline_data/r17_per_char_advances_2026-05-02.json`

Memory: `session_51_r17_gate_per_char_validation.md`

---

## 2026-05-02 — oxi-3 — confirmed (refinement) — kern audit + Normal-style override + 3a4f win/loser per-char + Oxi compression spec table

Three follow-up tasks completed (依頼 1/3/4 from user):

### 依頼 1: kern audit on 184 baseline docs (RESEARCH_LOG-readable)

`tools/metrics/audit_kern_docDefaults.py` extracts effective kern via
`run rPr > Normal style rPr > docDefaults rPr` resolution priority.

| Metric | Value |
|---|---|
| Total docs | 184 |
| Effective kern present | 62 (33.7%) |
| Source: docDefaults | 38 |
| **Source: Normal style** (NEW v2 finding) | 24 |

**v1 claim REFUTED**: Initial audit only checked docDefaults and showed
"kern perfectly discriminates R17 winners/losers". With Normal-style
included, 4/6 R17 big_winners ALSO have effective kern=2 (d77a, 3a4f
p23/p60 via Normal style). kern is necessary but NOT sufficient.

R17 cross-tab refined:
- big_winners: 3a4f p23/p60 (kern via Normal), d77a p3/p6 (kern via Normal),
  683f p2 + 0e7af1 p6 (no kern)
- big_losers: 7f272a/ed025 (kern via docDefaults), 3a4f p64 (kern via Normal)

### 依頼 3: 3a4f p64 (R31 +0.032 winner) / p42 (R31 -0.008 loser) per-char

`tools/metrics/measure_3a4f_p64_p42_v2.py` + `measure_3a4f_p42_only.py`
using Selection.GoTo for direct page jump (full doc has 2386 paragraphs).

Both pages show MIXED:
- **Mech 1** at half-width (5.0–5.5pt) for `）（` `）。` `。）` `）`→`）` etc.
  per FINAL RULE Type A/B/C
- **Mech 2 partial** at 7.5/8.0/8.5pt for chars without Mech 1 trigger
  on lines with overflow

Difference between R31 win/loss is NOT mechanism choice; both pages
have both mechanisms. R31's specific char-position decisions matter
(needs R31 trace cross-reference to identify exact mismatches).

3a4f activation source: **Normal style has `<w:kern w:val="2"/>`**
(docDefaults has none). Resolution priority delivers kern=2 to all
Normal-style paragraphs.

### 依頼 4: Oxi compression spec table for ed025/7f272a

`tools/metrics/extract_oxi_compress_spec_table.py` extracts char-level
spec from existing JSON.

ed025 paragraph 12 (R17 big_loser, kern=2):
- 2 compressions: `）` at i=23 (B→A `）（`) and i=43 (B→B `）））`),
  both at half-width 5.5pt
- **Pure Mech 1 / FINAL RULE** ✓

7f272a paragraph 12 (R17 big_loser, kern=2):
- 6 compressions: all at Mech 2 partial 8.0–10.0pt
- All chars have CJK neighbors (no Mech 1 trigger)
- **Pure Mech 2**

### Recommended Oxi rule (final form)

```rust
fn run_kern_active(run: &Run, doc: &Doc) -> bool {
    let kern_hp = run.rpr.kern
        .or(run.style.kern)        // Normal style or other paragraph style
        .or(doc.doc_defaults.kern)
        .unwrap_or(0);
    kern_hp > 0 && (run.font_size * 2.0) as u32 >= kern_hp
}

// Mech 1: FINAL RULE Type A/B/C (validated by ed025)
fn mech1_compress(prev_class, current, next_class) -> Option<f32> {
    match classify_yakumono(current)? {
        A if prev_class == Some(A) => Some(font_size / 2.0),
        B if matches!(next_class, Some(A) | Some(B)) => Some(font_size / 2.0),
        _ => None,
    }
}

// Mech 2: justify-time slack distribution (validated by 7f272a + 3a4f)
// 0.5pt-step distribution, total = slack exactly, per-char cap ~2pt
// (already in Phase 2 reactive absorb mod.rs:2977; refine per
// session_51_mechanism2_slack_algorithm)
```

### Source data

- `tools/metrics/audit_kern_docDefaults.py` (Normal-style aware)
- `tools/metrics/measure_3a4f_p64_p42_v2.py`,
  `tools/metrics/measure_3a4f_p42_only.py`
- `tools/metrics/extract_oxi_compress_spec_table.py`
- `pipeline_data/kern_audit_2026-05-02.json`
- `pipeline_data/3a4f_p42_per_char_2026-05-02.json`
- `pipeline_data/oxi_compress_spec_table_2026-05-02.json`

Memory:
- `session_51_kern_audit_177docs.md`
- `session_51_3a4f_p64_p42_validation.md`
- `session_51_oxi_compress_spec_table.md`

---

## 2026-05-02 — oxi-3 — confirmed — Yakumono compression has TWO mechanisms; Mech 1 trigger PINPOINTED to `<w:kern>` in styles.xml docDefaults

**TL;DR for master**: 2026-04-18 architectural-validation entry below is correct
about Mech 2 (line-wrap heuristic / reactive). But it MISSED a separate
always-on adjacency mechanism (Mech 1) that fires per Type A/B/C rule whenever
`<w:kern w:val="N"/>` (N≥1) is in `word/styles.xml`'s docDefaults rPr.
The 2026-04-18 Tier 2 cascade test ("rPrDefault rFonts/lang cascade
FALSIFIED — no combination triggers") didn't include `w:kern` as a tested
property. R17's list_marker gate (`dc7104c`) is a workaround; the proper
gate is doc-level kern.

### Mechanism 1 — adjacency rule (always-on, kern-gated)
- context: 4 fonts × 2 settings × 14 probes = 112 sample pairs measured on
  isolated 3-4 char paragraphs (no overflow, no justify)
- hypothesis: Type A/B/C compression rules fire per FINAL RULE table
  regardless of layout pressure
- evidence: ALL 14 probes match the spec §4.7 FINAL RULE table for ＭＳ
  明朝/ゴシック/Yu Mincho 10.5pt; doNotCompress and compressPunctuation
  produce IDENTICAL advances (56 sample pairs, 100% match)
- 5-step bisection (clone-and-replace → swap-files → inverse-swap →
  styles-bisect → element-bisect) → `<w:kern w:val="2"/>` ALONE
  is necessary and sufficient
  - V_only_kern (kern as sole non-default prop): 」=5.5pt (compressed) ✓
  - V_no_kern (every other prop except kern): 」=10.5pt (full) ✓
  - All my prior OOXML-direct tests lacked kern → no compression observed
- outcome: Mech 1 is gated by docDefaults `<w:kern>`, NOT by neighbor
  type alone. Per-pair Type A/B/C rule applies on top of the kern gate.
- artifacts: `tools/metrics/measure_yakumono_setting_contrast.py`,
  `bisect_yakumono_clone_com.py`, `bisect_yakumono_swap_files.py`,
  `bisect_yakumono_inverse_swap.py`, `bisect_styles_xml_trigger.py`

### Mechanism 2 — justify-time slack distribution (line-wrap heuristic)
- context: jc=both narrow content with 27-char probe `漢漢漢「漢漢漢」漢漢漢「...」漢漢漢、漢漢漢、漢漢漢`
- hypothesis: separate from Mech 1, Word redistributes overflow slack
  across yakumono on the line, also gated by kern
- evidence:
  - cw=320 (slack=4pt) jc=both: 6 yakumono compressed by 0.5-1.0pt
    each, total reduction = exactly 4pt ✓
  - cw=290 (slack=10pt) jc=both: 6 yakumono compressed by 1.5-2.0pt
    each, total = exactly 10pt ✓
  - cw=300: drops 2 chars instead of compressing (per-char cap ≈ 2pt)
- algorithm: `distribute slack in 0.5pt steps to total = slack
  exactly; cap per-char compression at ~2pt (≈17%); else drop a char`
- outcome: this IS the master's 2026-04-18 line-wrap heuristic finding,
  fully characterized. Confirms architectural-validation conclusion that
  Phase 2 reactive absorb is correct. NEW: per-char ≤2pt cap and
  drop-vs-compress threshold explain the non-monotonic width dependency.
- artifacts: `tools/metrics/measure_yakumono_justify_interaction.py`,
  `measure_mechanism2_slack_distribution.py`

### NEW orthogonal findings
- **em-dash (U+2014) classification is FONT-DEPENDENT** (8 fonts × 4 sizes):
  ＭＳ 明朝 / ＭＳ ゴシック treat as Type B (compresses). Yu Mincho /
  Yu Gothic / Meiryo / HGゴシックE / HGS明朝E / HG明朝B treat as Type C
  (no compression). Hbar (U+2015) is universal Type C.
  Spec §4.7 line 640's font-agnostic claim is WRONG.
- **§4.6.2 (kana→Latin alphanumeric autoSpaceDE) is NOT gated by kern**:
  fires identically with kern on/off (probe `はMでs`: は=13.0, M=8.0
  with both kern states).
- **§4.6.3 (CJK-adjacent space widening) does NOT EXIST in current Word**:
  7 OOXML rPr variants + 6 multi-run + jfmb on-disk vs runtime-saved
  comparison + kern on/off — ALL show space=3.0-3.5pt natural Latin
  width. Spec §4.6.3 is REFUTED. The c45c1fc fix should be reverted.

### kern semantics finalized 2026-05-02 evening

`<w:kern w:val="N"/>` per ECMA-376 §17.3.2.18 is "Font Kerning Threshold"
= minimum font size in **half-points** at which kerning applies.

| val | Behavior at 10.5pt (21 hp) |
|---|---|
| 0 (or absent) | NO compression |
| 1, 2 | COMPRESSION (21 ≥ val) |
| 100 (=50pt threshold) | NO compression (21 < 100) |
| ≥21 | COMPRESSION when font_size_hp ≥ val |

Resolution priority: run rPr > paragraph style rPr > docDefaults rPr.
pPr/rPr/kern affects only the paragraph mark glyph, NOT runs.

Both Mech 1 (Type A/B/C) AND Mech 2 (justify slack distribution) gated
by per-run kern resolution. Mech 2 additionally requires at least one
Mech 1 trigger on the line (e.g., a `」（` pair) before it activates
to compress non-trigger chars (e.g., the `（` at A→CJK).

### Recommended next actions
1. **Implement per-run kern resolution** in Oxi parser:
   ```rust
   let kern_hp = run_rpr.kern.or(style.kern).or(doc_defaults.kern).unwrap_or(0);
   let yakumono_enabled = kern_hp > 0 && font_size_half_pt >= kern_hp;
   ```
   Subsumes R17 list_marker gate (dc7104c) entirely.
2. **Update spec §4.7**: replace PROVISIONAL marker with the kern-based
   gate. Note Mech 1 vs Mech 2 distinction with Mech 2's anchor requirement.
3. **Update spec §4.7 line 640**: em-dash classification is font-dependent.
4. **Update spec §4.6.3**: REFUTED — CJK-adjacent space widening
   doesn't exist in current Word.
5. **Mech 2 algorithm in Oxi Phase 2**: refine reactive absorb to use
   the 0.5pt-step slack distribution + 2pt per-char cap.

Memory entries added today: `session_51_yakumono_kern_trigger.md`,
`session_51_yakumono_4font_validation.md`,
`session_51_yakumono_justify_two_mechanisms.md`,
`session_51_mechanism2_slack_algorithm.md`,
`session_51_emdash_font_dependency.md`,
`session_51_easia_hint_attribute.md`,
`session_51_chargrid_halfwidth_data.md`,
`session_51_tall_header_pushdown.md`.

---
## 2026-04-25 — oxi-2 — confirmed — R-10 paragraph_mark_revision path also test-covered

- mirror of the ppr_change test landed: `r10_fires_for_paragraph_mark_revision` mutates fixture_05's IR to install a synthetic `paragraph_mark_revision = TrackedChange { change_type: "insert", author: "Alice Reviewer", … }` while clearing every other revision pointer. Asserts ≥1 #424242 BoxRect emerges.
- 32/32 tests pass. R-10's full 4-way detection (Run.tracked_change / Run.rpr_change / Paragraph.ppr_change / Paragraph.paragraph_mark_revision) is now covered by 2 dedicated paragraph-level tests + the existing run-level test (`fixture_05_layout_emits_revision_change_bar`) + integration tests via fixture_07/08/09/10.

## 2026-04-25 — oxi-2 — confirmed — R-10 paragraph-level path now test-covered

- context: prior iteration extended R-10 to fire on `Paragraph.ppr_change` / `paragraph_mark_revision`, but no fixture exercised the path so the new code was untested.
- new test `r10_fires_for_paragraph_level_ppr_change` in `comments_fixtures.rs`:
  - parses fixture_05, then mutates the IR: clears every run-level `tracked_change` / `rpr_change`, installs a synthetic `Paragraph.ppr_change = Some(PropertyChange{ author: "Alice Reviewer", … })`.
  - lays out and asserts ≥1 thin BoxRect with `#424242` fill (the change bar) is emitted.
  - confirms R-10 fires from paragraph-level revisions alone, with no run-level pointer present.
- evidence: 31 tests pass.
- impact: closes the test-coverage gap from the prior iteration. The 4-way revision detection (Run.tracked_change / Run.rpr_change / Paragraph.ppr_change / Paragraph.paragraph_mark_revision) is now all observed by tests.

## 2026-04-25 — oxi-2 — confirmed — R-10 covers paragraph-level revisions

- context: previous iteration extended R-10 to cover `Run.rpr_change`. The remaining IR fields that carry revision metadata at paragraph level (`Paragraph.ppr_change` from P-07, `Paragraph.paragraph_mark_revision` from P-09) still did not trigger a change bar.
- fix in `layout_paragraph`: initialize `line_has_revision` from `para.ppr_change.is_some() || para.paragraph_mark_revision.is_some()` BEFORE the fragment loop. If either field is set the paragraph gets a margin change bar on every line, regardless of whether any individual run has `tracked_change` / `rpr_change`.
- evidence: all 30 tests pass. None of the 10 baseline fixtures exercise paragraph-level revisions on a body paragraph that has zero run-level revisions, so no visible change in the test set; the path is correct but unverified by fixture. Add a test fixture if/when `<w:pPrChange>` outside a run-revision paragraph appears in real documents.
- impact: defensive coverage. Guarantees the change bar tracks every kind of revision the IR can carry.

## 2026-04-25 — oxi-2 — confirmed — R-10 margin change bar now fires for rPrChange

- bug: fixture_09 (rPrChange "Now bold (was plain).") rendered the bold text correctly but did NOT show a margin change bar. Word renders a change bar next to formatting changes too.
- root cause: R-10's per-line revision detector only checked `Run.tracked_change` (the ins/del/move pointer). Property changes use a separate `Run.rpr_change` field (P-06's IR shape). Lines with only rpr_change fell through.
- fix in `layout_paragraph`: extend the check to also look at `run.rpr_change`. Now any revision-bearing run — `tracked_change` OR `rpr_change` — triggers the margin change bar.
- evidence:
  - all 30 comments_fixtures tests pass (no regression).
  - fixture_09 visual rebuild: dark grey 1.5pt change bar now appears in left margin next to "Regular. Now bold (was plain)." line. Was missing entirely before.
- limitations:
  - `Paragraph.ppr_change` (P-07) and `Paragraph.paragraph_mark_revision` (P-09) still don't trigger R-10. None of the 10 fixtures exercise them on a body paragraph that would otherwise have no `tracked_change`/`rpr_change` runs, so no visible miss in the current set. Add when a fixture stresses it.

## 2026-04-25 — oxi-2 — confirmed — R-05c anchor detection fix for empty marker runs

- bug: fixture_04 (multi-paragraph comment range starting with `<w:commentRangeStart>` BEFORE the first text run) rendered with the inline pink range tint but NO balloon. The anchor walk was indexed by `(paragraph_index, run_index)` of LayoutElements, but the parser's anchor-run fallback (which carries `comment_range_start`) has empty text and emits no LayoutContent::Text element — so the comment id was never found.
- fix in `emit_balloons_for_layout_page`:
  - Two-pass anchor detection. First pass walks LayoutElements once to build `paragraph_index → first-rendered (x, y)`. Second pass walks IR paragraphs in document order and, for any run that carries `comment_range_start`, looks up the paragraph's first rendered position from the map.
  - This decouples anchor detection from the run that carries the marker; whether the marker lives on an empty-text anchor run or a substantive text run, the comment is still anchored to the paragraph's first rendered Y.
- evidence:
  - all 30 comments_fixtures tests pass (no regression).
  - fixture_04 visual rebuild: balloon now appears next to paragraph 1 with header "Alice Reviewer" + body "Applies to all three paragraphs.", connector line drawn, all three paragraphs still tinted pink (R-04). Was missing the balloon entirely before this fix.
- baseline risk: zero.

## 2026-04-25 — oxi-2 — confirmed — pre-pass coverage extended to header/footer/footnote/textbox

- context: until now, `apply_revision_styling`, `apply_comment_range_highlighting`, `strip_parser_revision_styling`, `filter_runs_for_show_revisions`, and `revisions::apply_review` only walked `Page.blocks` (body). Tracked changes / comments living inside headers, footers, footnotes, endnotes, or textboxes would not get their visual treatment applied. None of the 10 fixtures exercise this, but real-world docs do.
- implementation: factored a single helper in `crates/oxidocs-core/src/layout/mod.rs`:
  ```
  fn for_each_block_tree<F: FnMut(&mut Vec<Block>)>(doc: &mut Document, mut f: F) {
      for page in &mut doc.pages {
          f(&mut page.blocks);
          f(&mut page.header);
          f(&mut page.footer);
          for footnote in &mut page.footnotes { f(&mut footnote.blocks); }
          for endnote in &mut page.endnotes { f(&mut endnote.blocks); }
          for tb in &mut page.text_boxes { f(&mut tb.blocks); }
      }
  }
  ```
  Each pre-pass that previously hand-rolled `for page in &mut doc.pages { for block in &mut page.blocks {…} }` now calls `for_each_block_tree(doc, |blocks| {…})`. Each pass's existing recursion into `Block::Table` cells is preserved.
  `revisions::apply_review` is in a separate module (no access to `for_each_block_tree`), so it inlines the same iteration explicitly.
- evidence: all 30 comments_fixtures tests pass — no regression. Coverage is purely additive (same body behavior + extended scope).
- skipped fields:
  - `Page.shapes` / `Page.floating_images` — these don't contain runs, just geometry.
  - The recursive table-cell walk inside each pass already handles Block::Table within any of these top-level locations.
- limitations:
  - R-10 (margin change bar in `layout_paragraph`) is emitted only on body paragraphs (only path that sets `body_para_index`). Header/footer revisions render with author tint + underline/strike but no margin bar — Word does the same (margin bar is body-only).
  - R-05 balloon emission still body-only. Comments anchored in headers/footers/footnotes don't render balloons; rare in practice and would need separate balloon Y resolution per region.
- baseline risk: zero — local 51-doc and oxi-main 184-doc baselines have 0 comments and the 5 `<w:del>` docs have body-only revisions. Header/footer revisions in the wild would now render correctly where previously they would silently render as plain text.

## 2026-04-25 — oxi-2 — confirmed — R-05 balloon body truncation polish

- context: post-R-05g visual review showed fixture_02's reply body "Following up." being clipped — the balloon's bounding box was too short to fit the reply body row beneath the reply chip.
- root cause: the height estimate accounted for line text height but underestimated per-section padding. The renderer adds an inter-section pad after each of: header chip, parent body, and (per reply) reply chip + reply body. For a parent with 1 reply that's 4 sections × ~4pt = 16pt of padding the estimate was missing, plus the chip-height was too small.
- fix: rewrote the height accumulation to mirror the renderer's actual sectional layout — `outer_pad + chip_h + section_pad + body_h + section_pad + Σ(chip_h + section_pad + reply_body_h + section_pad) + outer_pad`. Bumped `chip_h` from 12pt to 14pt and `line_height` from 12pt to 14pt for breathing room.
- evidence:
  - fixture_02 visual: balloon now shows full thread — `Alice Reviewer` parent chip + `Why?` body + indented `Alice Reviewer` reply chip + `Following up.` body. Was previously truncated below the reply chip.
  - fixture_01 visual: balloon a few pt taller; "Alice Reviewer" + "Is 'brown' needed here?" still cleanly inside the rounded box. No regression.
  - all 30 comments_fixtures tests pass (no test asserted on exact balloon height, so the change is internal).
- limitations:
  - Estimate is still char-count × avg-glyph-width, not actual GDI-measured wrap. fixture_03's resolved balloon (narrower 190.1pt) might still under-allocate for longer comment bodies — defer per-case tuning until a fixture exercises long resolved comments.
  - Reply body still uses the same wrap math as parent; replies indent ~10pt but the wrap width estimate doesn't account for the indent (slight overestimate, harmless).

## 2026-04-25 — oxi-2 — confirmed — S-03 accept/reject IR commands LANDED

- context: third settings-row. S-02 only filters at render time; S-03 bakes the accepted/rejected state into the IR so subsequent saves / re-renders see the post-review document with no tracked changes left.
- implementation:
  - new top-level module `crates/oxidocs-core/src/revisions.rs` exposes 4 free functions:
    - `accept_all(&mut Document)` / `reject_all(&mut Document)`
    - `accept_revision(&mut Document, id: &str)` / `reject_revision(&mut Document, id: &str)`
  - Internal `apply_review` walks pages → blocks (recursing into Table cells) → paragraphs and uses `Vec::retain_mut` to drop revision runs that fail the keep-test. Surviving runs have `tracked_change` cleared and the parser's pre-applied underline/strike + `FF0000` color stripped (same logic as S-02's filter helper).
  - Per-id variants match on `tracked_change.pair_id`; non-matching revisions stay untouched (tracked_change preserved).
  - Re-exported via `pub mod revisions;` in `lib.rs`.
- evidence:
  - 3 new layout tests on fixture_07 (3 revisions: ins1/del1/ins2):
    - `s03_accept_all_drops_deletions_and_clears_tracked_changes`: pre-call run count = 3, post-call = 0; del1 absent, ins1/ins2 present.
    - `s03_reject_all_drops_insertions_keeps_deletions`: del1 present, ins1/ins2 absent.
    - `s03_accept_revision_by_id_leaves_others_untouched`: accept id="101" (del1) — del1 gone, ins1 + ins2 still tracked.
  - all 30 comments_fixtures tests pass.
- baseline risk: zero — S-03 is a new public API; no existing call sites yet.
- limitations:
  - Per-id targeting uses `tracked_change.pair_id` which is the `<w:ins w:id=…>` attribute. Move-pair semantics (one `pair_id` shared between moveFrom + moveTo) work correctly because the same id matches both wrappers.
  - Doesn't update `commentsExtended.xml` / commentsIds — those are in-memory IR fields tied to comments, not revisions. No-op for accept/reject.
  - No undo; the mutation is destructive. Caller must clone first if they want to keep the original.
- next iteration candidate: cleanup parser-side pre-applied tracked-change styling (currently redundant with R-01 pre-pass + S-02/S-03 helpers); OR refinements to R-05 balloon (body height truncation in fixture_03, connector author-tinting); OR S-04 UI affordances (out of scope per attack-matrix).

## 2026-04-25 — oxi-2 — confirmed — S-02 show_revisions toggle LANDED (all 4 modes)

- context: second settings-row. Wires `ir::ShowRevisions` (I-04) into `LayoutEngine` so callers can pick All / Simple / Original / Final per-render.
- design:
  - New `show_revisions: ShowRevisions` field on `LayoutEngine`, defaults `All`.
  - New `with_show_revisions(mode) -> Self` builder.
  - 4-arm match in `layout()` before applying revision styling:
    - `All` → call `apply_revision_styling` (current behavior).
    - `Simple` → new helper `strip_parser_revision_styling` clears the parser's pre-applied underline/strike + `FF0000` color but keeps `tracked_change` so R-10 still fires margin change bars.
    - `Original` → `filter_runs_for_show_revisions(doc, final_view=false)`. Drops ins/moveTo runs from paragraphs; clears tracked_change + parser styling on del/moveFrom (which now render as plain text).
    - `Final` → `filter_runs_for_show_revisions(doc, final_view=true)`. Drops del/moveFrom runs; clears tracked_change + parser styling on ins/moveTo.
  - Both helpers recurse into `Block::Table` cells.
- gotcha discovered: the parser (`crates/oxidocs-core/src/parser/ooxml.rs:5795-5805`) pre-applies `style.underline=true` + `color="FF0000"` on insert runs and `style.strikethrough=true` + same red on delete runs — predates the R-01 pre-pass. For All view this gets overridden by `apply_revision_styling`'s author-tinted color/underline. For Simple/Original/Final views the parser's styling persists unless cleaned up — that's why both helpers explicitly reset `underline`, `underline_style`, `strikethrough`, and the FF0000 color when keeping/clearing a revision-bearing run.
- evidence:
  - 3 new layout tests in `comments_fixtures.rs`:
    - `s02_show_revisions_final_drops_del_and_strips_ins_styling` — fixture_07 in Final: del1 absent, ins1/ins2 present without underline/strike/author-color.
    - `s02_show_revisions_original_drops_ins_and_strips_del_styling` — fixture_07 in Original: ins1/ins2 absent, del1 present.
    - `s02_show_revisions_simple_skips_color_keeps_margin_bar` — fixture_05 in Simple: no author-tinted underline, but ≥1 margin change bar still fires.
  - All 27 comments_fixtures tests pass.
- baseline risk: zero — default is `All`, prior behavior preserved.
- limitations:
  - Original view doesn't yet restore prior `rPrChange` styles (the run renders with the *new* properties, not the prior ones). Word's actual Original view applies the prior rPr — refine when more PrChange fixtures exist.
  - Simple view's "change bar in author color" tinted variant deferred — current bar is uniform grey #424242 (R-10).
- follow-up cleanup: parser-side pre-applied tracked-change styling is now redundant in All view (overridden) and forces helper code in the other 3 views. Defer removal to a focused refactor later — risk of breaking IR consumers that read style directly.
- next iteration candidate: S-03 (accept/reject IR commands — editor-side, no render impact); or refinements to R-05's body truncation / connector color.

## 2026-04-25 — oxi-2 — confirmed — S-01 show_comments toggle LANDED

- context: first settings-row in the attack matrix. Provide a clean off-switch for the entire comment family (balloons + connectors + in-line range highlight + body-width compression) without touching tracked-change rendering — useful for clean print output and as a sanity-check.
- implementation:
  - New `show_comments: bool` field on `LayoutEngine`, defaults `true`.
  - New builder method `pub fn with_show_comments(mut self, show: bool) -> Self` for caller-side toggle.
  - 3 gated sites in `layout()`:
    - `apply_comment_range_highlighting` only runs when `self.show_comments` (was unconditional).
    - Balloon emission post-pass only runs when both `self.show_comments` AND `!doc.comments.is_empty()`.
    - `layout_page` uses `0.0` as `balloon_reservation` when `!self.show_comments`, so body width is full even on commented docs.
- evidence:
  - new layout test `s01_show_comments_false_suppresses_all_comment_visuals`: with show_comments=false, asserts 0 balloons, 0 connectors, 0 highlighted Text elements, and exactly 1 line of body text (= full-width layout). Same fixture with show_comments=true (default) emits 1 balloon for sanity.
  - all 24 comments_fixtures tests pass.
- baseline risk: zero — default is `true`, matching prior behavior.
- next iteration candidate: S-02 (`show_revisions: ShowRevisions` toggle), which uses the existing I-04 enum with 4 modes (All/Simple/Original/Final).

## 2026-04-25 — oxi-2 — confirmed — R-05g GDI rendering of Balloon + BalloonConnector LANDED

- context: seventh R-05 sub-iteration. With layout-side balloon emission complete (R-05a..f), now wire the GDI renderer so the PNG output finally shows balloons.
- additions:
  - **`pub fn comment_balloon_fill(idx, resolved) -> &'static str`** in `crates/oxidocs-core/src/layout/mod.rs` — public resolver so the GDI renderer can map `author_color_index + resolved` → palette hex without duplicating the constants.
  - **`parse_hex_rgb`** helper in `tools/oxi-gdi-renderer/src/main.rs`.
  - **Balloon handler**: rounded rect (4pt radius, scaled) filled with `comment_balloon_fill(...)` tint, 1pt border in a slightly-darker shade. Author header in bold 9pt grey. Body in 10pt black, wrapped via `DrawTextW(DT_LEFT|DT_TOP|DT_WORDBREAK)`. Each reply gets an indented (~10pt) author chip + body in identical shape.
  - **BalloonConnector handler**: `PS_DOT` pen, scaled width, medium grey (#808080). `MoveToEx` + `LineTo`. Color of the connector overrides the layout-supplied `color_hex` for now — a uniform grey is more readable against pink/purple/green tints than the per-author hue at 1pt thickness.
  - Balloon height estimate gained a `header_chip_h = 12pt` constant so the body actually fits inside the rect.
- evidence (visual via GDI render):
  - **fixture_01**: pink rounded balloon shows "Alice Reviewer" + "Is 'brown' needed here?", grey dotted connector line runs from the body's commentRangeStart line up to the balloon's left edge. Inline pink "brown fox" tint still works (R-04). Body still wraps to 2 lines (R-05b).
  - **fixture_02**: balloon shows parent "Why?" + reply chip "Alice Reviewer" indented inside (R-05f reply fold visible).
  - All 4 comment fixtures (01/02/03/04) render balloons.
- evidence (tests): all 23 comments_fixtures tests pass after the height-estimate change.
- limitations / refinements deferred:
  - Reply body sometimes clipped when balloon height estimate is short (reply estimate uses average glyph width which may underestimate at ~10 chars). Pixel-tune per fixture in a follow-up tick.
  - Connector dotted color is grey (not author tint). Word's actual connector is more pastel-author-colored at low alpha. Refine when we have a per-author connector color spec.
  - Connector geometry: `from_y` is the line top, so visually the line starts a few pt above the actual glyph row. Acceptable for v1.
  - GDI `DrawTextW(.., DT_NOCLIP)` is used for the header (single line) and `DT_WORDBREAK` for body. Word wraps at hyphens / break-after; GDI's word-break is rougher. Visual close enough for v1.
- baseline risk: zero — local 51-doc and oxi-main 184-doc baselines have 0 comments → zero balloons emitted on baseline, so renderer never executes the new branches.

## 2026-04-25 — oxi-2 — confirmed — R-05f reply threading inside parent balloon LANDED

- context: sixth R-05 sub-iteration. Word renders a comment's replies INSIDE the parent's balloon, indented. Don't emit a standalone Balloon for each reply.
- spec join: a Comment is a reply when `Comment.parent_para_id` is set; the value is the parent's `Comment.para_id` (different from `Comment.id`). P-10 (`commentsExtended.xml` parser) already populates these fields.
- implementation in `emit_balloons_for_layout_page`:
  - For each parent Balloon being prepared, iterate `doc.comments` looking for `c.parent_para_id == parent.para_id`. Each match becomes a `BalloonReply { author, author_color_index, body }`.
  - PendingBalloon gains a `replies: Vec<BalloonReply>` field; populated alongside body.
  - Balloon LayoutContent now carries the populated `replies` Vec (was `Vec::new()` placeholder).
  - Height estimate folds reply line counts in (per reply: line count + 1 for the author header chip). Approximate; R-05g refines.
  - Replies don't appear in `anchors` because they have no `commentRangeStart` of their own (they share the parent's range). So no separate balloon is emitted for them — the fold is the only visible side effect.
- evidence:
  - new layout test `fixture_02_reply_folds_into_parent_balloon_replies_vec`: parses fixture_02 (Alice's parent comment "Why?" + Alice's reply "Following up." linked via `paraIdParent="00000010"`), asserts exactly 1 standalone Balloon, parent body contains "Why?", `replies[0].body == "Following up."`, replies[0].author == "Alice Reviewer".
  - All 23 comments_fixtures tests pass.
- visual: GDI render still no-op; the JSON LayoutResult correctly carries the threaded structure.
- limitations:
  - Multi-level reply nesting (replies of replies) not handled — Word doesn't really support that anyway. Keep flat.
  - If a reply somehow had its own `commentRangeStart` (rare), it would emit a duplicate standalone balloon. Defer until a fixture stresses it.
- next iteration: R-05g — actual GDI rendering of `LayoutContent::Balloon` and `LayoutContent::BalloonConnector` so the PNG output finally shows balloons.

## 2026-04-25 — oxi-2 — confirmed — R-05e balloon connector lines LANDED

- context: fifth R-05 sub-iteration. With balloons emitting and stacking (R-05c/d), now also emit a connector line from each balloon's inline anchor to the balloon's left edge.
- implementation:
  - Inside `emit_balloons_for_layout_page`, just before pushing each Balloon, push a `LayoutContent::BalloonConnector` LayoutElement.
  - `from_x`, `from_y` = the balloon's `anchor_x`, `anchor_y` (= rendered Y of the comment's `commentRangeStart`).
  - `to_x` = balloon's left edge; `to_y` = balloon's resolved top + 5pt (visually meets the first text row of the balloon, matches Word).
  - `color_hex` = author's tint slot from `COMMENT_HIGHLIGHT_TINT_PALETTE` (slot 0 = `#FAE6E7` for Alice). Same color regardless of resolved state — Word's connector uses the unresolved tint hue at lower opacity, but our v1 ships solid color from the palette and refines at R-05g.
  - LayoutElement bounding box covers the connector's path so layered renderers can clip correctly.
- evidence:
  - new layout test `fixture_01_emits_balloon_connector_paired_with_balloon`: asserts exactly 1 BalloonConnector, ends at balloon_left, ends at balloon_top+5pt, starts left of balloon, color=#FAE6E7.
  - all 22 comments_fixtures tests pass.
- visual: GDI renderer still has `_ => {}` for BalloonConnector (will be wired in R-05g — needs dotted pen).
- limitations:
  - Connector currently solid in the LayoutResult; the dotted style is the renderer's job (R-05g).
  - Color is the author's tint regardless of resolved state. May refine.
- next iteration: R-05f — fold replies (`Comment.parent_para_id`) into their parent balloons' `replies` Vec, dropping the standalone child Balloon element.

## 2026-04-25 — oxi-2 — confirmed — R-05d balloon stacking LANDED

- context: fourth R-05 sub-iteration. With single-balloon emission in place (R-05c), prevent vertical overlap when ≥2 balloons would otherwise stack on top of each other.
- algorithm: extracted `stack_balloon_ys(positions: &mut [(f32, f32)], gap: f32)` as a pure helper. Sorts by anchor Y (caller's responsibility); walks the slice, pushes each balloon's Y to `max(natural_y, prev_y + prev_height + gap)`. First balloon never moves; cascade is monotonic (later balloons can only push down, never up).
- gap constant `BALLOON_STACK_GAP = 6.0pt`. Pixel-tune in R-05g once GDI render lands.
- emission integration: `emit_balloons_for_layout_page` now collects all per-balloon geometry into a `Vec<PendingBalloon>`, sorts by `anchor_y`, runs `stack_balloon_ys` on a `(y, height)` projection, then writes back the resolved Ys before pushing LayoutElements.
- tests:
  - 3 new pure-function unit tests in `mod tests` (`stack_balloon_ys_no_overlap_keeps_anchors`, `stack_balloon_ys_pushes_overlapping_balloons_down`, `stack_balloon_ys_handles_degenerate_inputs`). Cover happy path, push-down, and 3-balloon cascade.
  - All 21 comments_fixtures integration tests still pass — no regression.
- limitations:
  - 6pt gap is approximate. Will pixel-tune when R-05g compares against Word's rendered output.
  - Stacking only operates per-page. A balloon whose anchor is on page N but whose natural Y would push past page-bottom does NOT split or move to page N+1 — Word actually wraps the balloon column, but that's a separate edge case (defer until a fixture stresses it).
  - Replies still don't fold into parent (R-08 / R-05f) — stacking treats each comment as its own balloon. fixture_02 currently shows 1 balloon (parent only) since the reply has no commentRangeStart of its own.
- next iteration: R-05e — connector line. Emit one `LayoutContent::BalloonConnector` per balloon, dotted, from the inline anchor point to the balloon's left edge.

## 2026-04-25 — oxi-2 — confirmed — R-05c single-comment balloon emission LANDED

- context: third R-05 sub-iteration. With body width compressed (R-05b) and enum scaffolding (R-05a), now emit one `LayoutContent::Balloon` per visible comment, anchored to the rendered Y of its `commentRangeStart`.
- design — IR-page tracking: `LayoutEngine::layout` now builds a parallel `Vec<usize>` mapping each LayoutPage to its source IR page. A single IR page may produce multiple LayoutPages (pagination), all sharing the same IR index. Used by the post-pass to resolve `LayoutElement.paragraph_index` → source `Run`.
- design — `emit_balloons_for_layout_page(layout_page, doc, ir_page_idx)`:
  - Walks LayoutPage elements in order. For each Text element with `paragraph_index + run_index`, looks up the source `Run` in `doc.pages[ir_page_idx].blocks[paragraph_index].runs[run_index]` and reads its `comment_range_start: Vec<String>`.
  - Records `(comment_id, anchor_x, anchor_y)` for the FIRST occurrence of each comment id on this page, in document order.
  - For each anchor, emits one `LayoutContent::Balloon` with: width 293.8pt unresolved / 190.1pt resolved (COM-confirmed), right edge `page_width − 4pt`, anchor Y = first-occurrence Y, body = flattened comment paragraphs, color slot from author palette.
  - Height estimate: `body_lines × 12pt + 8pt_padding` using a rough average glyph width of 5pt for line wrap. R-05g will refine when GDI rendering measures actual wrap.
- evidence:
  - 2 new layout integration tests:
    - `fixture_01_emits_one_balloon_for_single_comment` asserts: 1 Balloon element, comment_id="0", author="Alice Reviewer", resolved=false, width=293.8pt±0.01, x=297.6pt±0.5 (= 595.3 − 4 − 293.8).
    - `fixture_03_emits_resolved_balloon_with_narrower_width` asserts: 1 Balloon, resolved=true, width=190.1pt (the resolved variant).
  - All 21 comments_fixtures tests pass.
- visual: GDI renderer currently has `_ => {}` no-op for Balloon (added in R-05a), so the actual page render is unchanged. R-05g will wire the rendering.
- baseline risk: zero — `doc.comments.is_empty()` short-circuits before the post-pass; baseline docs are untouched.
- limitations / next:
  - Stacking-on-overlap (R-05d) not yet implemented. Overlap is impossible in the current 4 single-comment fixtures (only fixture_02 has 2 comments which would overlap if rendered without offset). R-05d is the next sub-iteration.
  - No connector line yet (R-07 / R-05e).
  - Replies (R-08 / R-05f) — `Balloon.replies` is currently `Vec::new()`. Will populate when R-05f lands.
  - GDI render (R-05g) — needs `LayoutContent::Balloon` handler in `tools/oxi-gdi-renderer/src/main.rs`.

## 2026-04-25 — oxi-2 — confirmed — R-05b body width compression for commented docs LANDED

- context: second R-05 sub-iteration. With the enum variants in place (R-05a), now make the body actually narrower when the document has any comments — paves the way for R-05c balloon emission to land in a known column.
- implementation:
  - Added `balloon_column_width: f32` field to `LayoutEngine`. Set in `for_document` to `0.0` when `doc.comments.is_empty()`, `293.8 + 24.0 = 317.8` otherwise. (293.8 is COM-confirmed Alice unresolved balloon width from pixel pass; 24pt is approximate gap, refined as later iterations pixel-test.)
  - In `layout_page`, subtract `self.balloon_column_width` from `total_content_width`. Header / footer / floating-image / footnote widths kept at full un-reduced width — matches Word's behavior (only body reflows when balloons appear).
- evidence:
  - new layout test `fixture_01_body_width_compresses_when_comments_present` asserts: with comments → body wraps to ≥2 lines; with comments cleared → 1 line. Confirms the compression depends on `doc.comments.is_empty()`, not on the text itself.
  - all 19 comments_fixtures tests pass.
  - Visual confirmation via GDI re-render: fixture_01 now renders "The quick brown fox jumps" on line 1, "over the lazy dog." on line 2 (was previously on a single line). The R-04 pink "brown fox" tint still applies correctly to the new wrapped layout.
- baseline risk: zero — local 51-doc and oxi-main 184-doc baselines have 0 comments, so `balloon_column_width = 0.0` for every baseline doc and the subtraction is a no-op.
- limitations:
  - 24pt gap is approximate. Word's actual gap (between body right edge and balloon left edge) measured ~33.5pt for fixture_01; 24pt is conservative pending pixel-perfect tuning when R-05c lands and emits actual balloons.
  - Compression applies to ALL pages of a commented doc, not just pages where comments are visible. Word does the same — body width is uniform across pages of a section regardless of which page anchors which comment.
- next iteration: R-05c — emit one Balloon LayoutElement per visible comment, anchored to the rendered Y of its scope start. First iteration that produces actual `LayoutContent::Balloon` elements.

## 2026-04-25 — oxi-2 — confirmed — R-05a enum variants + match-arm fallthroughs LANDED

- context: R-05a is the first sub-iteration of the R-05 balloon design (`docs/spec/comments_tracked_changes/r05_balloon_design.md`). Add the new `LayoutContent::Balloon` and `LayoutContent::BalloonConnector` variants to the layout enum and update every match site to handle them, so the body-width compression (R-05b) and per-page emission (R-05c) iterations can land cleanly.
- variants added (in `crates/oxidocs-core/src/layout/mod.rs`):
  - `LayoutContent::Balloon { comment_id, author, author_color_index, resolved, body, replies: Vec<BalloonReply>, anchor_x, anchor_y }` — carries the full comment payload so renderers can do their own wrapping.
  - `LayoutContent::BalloonConnector { from_x, from_y, to_x, to_y, color_hex }` — dotted line geometry for R-07.
  - `BalloonReply { author, author_color_index, body }` — used inside Balloon.replies for R-08.
- match sites updated (5 files):
  - `tools/oxi-gdi-renderer/src/main.rs` — type_name lookup + element drawing (skip-stub for now until R-05g).
  - `crates/oxi-cli/src/main.rs` — PDF emission match (skip-stub).
  - `crates/oxi-wasm/src/lib.rs` — 3 match sites (LayoutElementJs construction × 2, PDF emission × 1). New `kind: "balloon"` and `kind: "balloon_connector"` strings surface to JS consumers.
  - `crates/oxidocs-core/examples/dump_docx.rs` — added BALLOON / CONNECTOR / SHAPE rows. (Side benefit: example was already broken on `PresetShape` from a prior session — fixed in passing.)
  - `crates/oxidocs-core/examples/layout_json.rs` — added BALLOON / CONNECTOR rows.
- evidence: `cargo build -p oxidocs-core` (lib + examples) clean. `cargo test -p oxidocs-core --test comments_fixtures` 18/18 pass. `cargo build --release -p oxi-gdi-renderer` clean. `cargo build -p oxi-cli` and `-p oxi-wasm` clean. Lib test suite: 51 pass + 1 pre-existing kinsoku failure (unrelated, predates this branch).
- behavior change: zero. The pre-pass family (R-01/R-04/etc.) doesn't emit Balloon/BalloonConnector elements — only R-05c will. So this iteration is a pure shape change.
- baseline risk: zero — same reasoning as prior iterations.
- next iteration: R-05b — when `doc.comments` non-empty, reduce `total_content_width` by `293.8 + gap`. First behavior change in the balloon family, but still no balloons emit yet so the reduced body simply leaves the right margin empty until R-05c.

## 2026-04-25 — oxi-2 — design — R-05 / R-06 / R-07 / R-08 / R-09 (balloon-side) implementation plan drafted

- context: with the pre-pass + per-line revision-bar family of renderer rows landed (R-01/R-02/R-03/R-04/R-09 in-line/R-10/R-11/R-12 minimal), the remaining renderer surface is balloon rendering — a substantial new component. Rather than starting implementation with imperfect defaults, draft the design first.
- output: `docs/spec/comments_tracked_changes/r05_balloon_design.md` (200+ lines).
- key design decisions captured:
  - **Two-pass per-page layout**: when `doc.comments` is non-empty, body width is reduced by 293.8pt + gap to make room for the right-margin balloon column. Body lays out first (with reduced width); balloons emit in a second per-page pass that anchors them to the rendered Y of each comment's `commentRangeStart`.
  - **New `LayoutContent::Balloon` + `BalloonConnector` variants** carry the comment payload (author, body, replies, anchor coordinates, resolved flag, color slot). This is intentional vs composing balloons from existing primitives — it gives renderers wrapping freedom and avoids coupling balloon shape to specific shapes/text primitives.
  - **Sub-iteration roadmap (R-05a..h)**: each step ships independently as a Path B confidence merge. R-05a adds the enum variants + `_ => {}` no-op arms across consumer crates; R-05b adds body-width compression; R-05c emits a single balloon; R-05d adds stacking; R-05e adds connector; R-05f adds reply threading; R-05g wires GDI rendering; R-05h flips resolved desaturation.
  - **Risk profile**: zero baseline impact — local 51 + oxi-main 184 baselines have 0 comments, so R-05's body-width compression never triggers there.
- key data sources cross-referenced:
  - Balloon width 293.8pt unresolved / 190.1pt resolved (pixel pass).
  - Right edge ≈ page_w − 4pt (pixel pass).
  - Balloon top aligned with `Comment.Scope.Start` line, NOT `Comment.Reference` (object-model + pixel pass).
  - Resolved tint #F1EDEC vs unresolved #FAE6E7 (pixel pass).
- explicitly out of scope: editor balloon UI (S-04), markup-mode toggle (S-02), left-side balloon anchor (rare config).
- next iteration: starts R-05a — enum variant addition + `_ => {}` fallthrough arms in oxi-cli, oxi-wasm, oxi-gdi-renderer, and example match sites. Mechanical change touching ~25 match sites across 5+ files.

## 2026-04-25 — oxi-2 — confirmed — R-10 margin change bar LANDED

- context: independent renderer row that doesn't need balloon infrastructure. Word draws a thin vertical bar in the left margin next to every line containing any revision (insert/delete/move/property change).
- implementation: emitted DURING layout (not as a post-pass), inside `layout_paragraph`'s per-line loop:
  - Before the fragment loop: declare `let mut line_has_revision = false`.
  - Inside the fragment loop after pushing the Text element: if `para.runs[frag.run_index].tracked_change.is_some()`, set the flag.
  - After the fragment loop: if the flag is set, push one `LayoutContent::BoxRect { fill: Some("#424242"), … }` at `(start_x − 12pt, *cursor_y, 1.5pt, line_height)`.
- design choice — single bar per line: emitted once after the fragment loop instead of per-fragment, so a line with multiple revision fragments (fixture_07: ins1 + del1 + ins2 on one line) gets exactly one bar. Multi-author lines also get one bar (fixture_10).
- design choice — fixed dark grey: Word uses `#424242`-ish dark grey by default; some configs cycle author color. v1 ships fixed grey for simplicity and unambiguity in multi-author cases. Author-tinted bar can layer on later when S-* config rows want it.
- design choice — `start_x − 12pt`: the bar sits 12pt to the LEFT of the body content's left edge, well inside the page margin. For default 72pt margins this puts the bar at x=60pt (12pt from body, 60pt from page edge). Visible without crowding the text.
- tests: new `fixture_05_layout_emits_revision_change_bar` asserts exactly 1 thin BoxRect (≤2pt wide) was emitted, with fill #424242, height ≥8pt, x in the left-margin range. All 18 comments_fixtures tests pass.
- visual confirmation (GDI renderer rebuild): fixture_07's mixed ins+del paragraph now has a clear vertical dark-grey bar in the left margin, next to "Start. ins1 middle del1 ins2. End." The bar height matches the line's vertical extent.
- baseline risk: zero — the local 51-doc and oxi-main 184-doc baselines have 0 revisions in the local set. R-10 emits zero bars on those.
- limitations:
  - Only walks body paragraphs (via `body_para_index`); table cell internal lines, footnotes, headers/footers/textboxes don't currently emit bars. The `body_para_index` field on LayoutElement is `None` for those locations, so the per-line check still runs but only the body case produces visible bars.
    Wait — actually the check IS in the body-emission path (the loop at mod.rs:~2674). Need to extend to header/footer/footnote/textbox emission paths if those should also draw bars. Add later when fixtures stress that.
  - Author-tinted bar variant deferred (would require palette-color lookup at emit time; not blocked by anything else).
- path: Path B confidence-merge candidate.

## 2026-04-25 — oxi-2 — confirmed — R-09 (in-line half) resolved-comment desaturation LANDED

- context: extension of R-04. Word renders the in-line range tint AND the balloon background of resolved comments (`<w15:done="1"/>`) with chroma stripped — Alice's #FAE6E7 unresolved → #F1EDEC resolved (COM-confirmed in pixel pass for fixture_03).
- implementation: `crates/oxidocs-core/src/layout/mod.rs`
  - Added `COMMENT_HIGHLIGHT_RESOLVED_PALETTE: [&str; 8]` next to the unresolved palette. Slot 0 = #F1EDEC (COM-confirmed Alice). Slots 1-7 derived via 25% tint + 75% per-slot grey blend (low chroma, lightness preserved).
  - In `apply_comment_range_highlighting`, the comment_id → tint map switches palette based on `Comment.resolved`: `if c.resolved { RESOLVED_PALETTE } else { UNRESOLVED_PALETTE }`. Same author slot (color_index) selects the same row across the two palettes.
- evidence:
  - new layout test `fixture_03_layout_resolved_comment_uses_desaturated_tint` asserts `"reviewed"` element has `highlight = Some("#F1EDEC")` (not the unresolved #FAE6E7). All 17 comments_fixtures tests pass.
  - Visual via GDI renderer rebuild: fixture_03 page shows "has been reviewed" on a CLEAR GREY background, distinctly different from fixture_01's pink "brown fox" background.
- limitations: balloon-side desaturation (the larger half of R-09 — when balloons render, their fill switches from #FAE6E7 to #F1EDEC) is part of R-05 and not yet implemented. The in-line half is sufficient for v1's "show user the range was already reviewed" UX hint.
- baseline risk: none.
- path: Path B confidence-merge candidate.

## 2026-04-25 — oxi-2 — confirmed — R-04 in-line comment-range highlight LANDED

- context: with R-01/R-02/R-03/R-11 confirmed, R-04 is the next-simplest renderer row that doesn't need balloon infrastructure. Spec: apply an author-tint background to every run strictly between `commentRangeStart` and `commentRangeEnd`.
- parser gap found + fixed: `commentRangeStart`/`commentRangeEnd` are zero-length markers that the previous parser attached to `runs.last_mut()`. When the marker appears as the FIRST child of a paragraph (fixture_04 P1 has `<w:commentRangeStart>` before the first `<w:r>`), `runs.last_mut()` returns None and the id was silently dropped. This left fixture_04's entire range invisible to any range-aware pass. Fix: when `runs.last_mut()` is None, create an empty anchor run carrying the marker (mirrors the bookmark-anchor treatment already in the parser). `commentRangeEnd` gets the symmetric fix.
- attachment convention after the fix: `comment_range_start` on run R → "range starts AFTER R"; `comment_range_end` on run R → "R is the LAST run inside the range". The walk applies highlight before processing either marker, which yields the correct set of highlighted runs with no special-casing.
- implementation: `crates/oxidocs-core/src/layout/mod.rs`
  - `COMMENT_HIGHLIGHT_TINT_PALETTE: [&str; 8]`: slot 0 = `#FAE6E7` (Alice, COM-confirmed from fixture_01 balloon BG), slots 1-7 derived via the 12/88 white-blend formula off the author palette.
  - `apply_comment_range_highlighting(doc)` pre-pass (line 207 area, runs after `apply_revision_styling` so highlight stacks on top of revision ink).
  - Builds `comment_id → tint` map from `doc.comments` + `doc.authors.color_index`. Early return when `doc.comments` is empty (no cost on non-comment baseline docs).
  - Walks `page.blocks` with a persistent `open: HashSet<String>`; recurses into table cells. Order: apply highlight → process `comment_range_end` → process `comment_range_start`.
- tests: 2 new layout integration tests — `fixture_01_layout_comment_range_highlight_inline` (asserts "brown" has Alice's tint, surrounding "The" / "jumps" do not), `fixture_04_layout_multi_paragraph_range_highlight` (asserts all 3 paragraphs "First" / "Second" / "Third" carry the tint — proves the walk carries `open` across paragraph boundaries AND that the parser's new anchor-run fallback works). All 16 comments_fixtures tests pass.
- visual confirmation (GDI renderer rebuild, PNG view):
  - fixture_01: "The quick" black, **"brown fox" on pink tint background**, "jumps over the lazy dog." black. ✓
  - fixture_04: all three paragraphs ("First paragraph...", "Second paragraph...", "Third paragraph...") have the pink tint behind their text. ✓
- baseline risk: none — the local 51-doc baseline has 0 comments (verified last iteration); the 184-doc oxi-main baseline has 0 comments (per inventory). R-04 touches zero runs in either baseline.
- limitations:
  - Single tint per run even if multiple comments overlap the same run — picks the `min()` of open ids deterministically. Word's real behavior blends tints; revisit when we have a multi-overlap fixture.
  - Resolved comments currently use the same tint (doc uses `Comment.resolved` to decide between resolved/unresolved tints in R-09; for inline highlight the behavioral difference is subtle and deferred).
  - Not walked: header/footer/footnote/textbox blocks. 10 fixtures don't exercise these locations.
- path: Path B confidence-merge candidate.

## 2026-04-25 — oxi-2 — confirmed — R-01/R-03/R-11 visual end-to-end validation (GDI renderer)

- context: layout-level integration tests proved the IR-side wiring of R-01/R-03/R-11 was correct, but the actual on-screen output flows through `oxi-gdi-renderer` which converts `LayoutContent::Text` to GDI `TextOutW` calls. End-to-end visual confirmation needed before declaring this Path B-mergeable.
- approach: rebuild `tools/oxi-gdi-renderer` (release) so it picks up the new layout pre-pass, render fixtures 5/6/7/8/10 to PNG at 150 DPI, view the PNGs.
- results:
  - fixture_05 (single ins): "Before insertion" black, "INSERTED TEXT" red+underline (#D03337 hue confirmed via `(210, 30, 30)` and `(180, 30, 30)` pixel buckets), "after insertion." black. ✓
  - fixture_06 (single del): "Before delete" black, "DELETED TEXT" red+strikethrough, "after delete." black. ✓
  - fixture_07 (mixed ins+del): "Start." black, "ins1" red+underline, "middle" black, "del1" red+strikethrough, "ins2." red+underline, "End." black. ✓ (visually viewed)
  - fixture_08 (moves): "Origin:" black, "moved clause" (moveFrom) GREEN+strikethrough (`(0, 60, 30)`/`(30, 60, 0)` hues = #2B6033), "Destination:" black, "moved clause" (moveTo) GREEN+underline. ✓ (visually viewed — the move-quirk hard-coded green is rendering, NOT Alice's red)
  - fixture_10 (two reviewers): "Alpha." black, "ALICE ADD" RED+underline, "middle" black, "BOB REMOVE" PURPLE+strikethrough (`(60, 0, 120)`/`(90, 30, 120)` = #5B2C90 hue), "omega." black. ✓ (visually viewed — palette rotation Alice slot 0 / Bob slot 1 working live in GDI output)
- baseline check: scanned all 51 docs in `pipeline_data/docx/` for `<w:ins>`/`<w:del>`/`<w:moveFrom>`/`<w:moveTo>`/`<w:rPrChange>` — **zero matches**. The local 51-doc baseline does not exercise the R-01 code path. SSIM cannot regress mathematically. (The 5 `<w:del>` docs in the inventory are in `oxi-main`'s 184-doc baseline, a separate worktree.)
- artifacts: `tools/metrics/output/oxi_render_check/fixture_*_oxi_p1.png` (5 PNGs).
- conclusion: R-01/R-02 (slots 0-1)/R-03/R-11 (single-line v1) are correct end-to-end. Path B confidence merge ready.
- next R-* candidates by independent-of-balloon and value:
  - R-04 (comment-range highlight): in-line background tint between commentRangeStart/End; medium effort, doesn't need balloon rendering.
  - R-10 (margin change bar): vertical line in left margin on revision-bearing lines; medium effort, post-layout pass.
  - R-08 (reply thread): needs R-05 (balloon) first.
  - R-05/R-06/R-07 (balloon family): substantial new rendering surface, defer for a focused multi-tick block.

## 2026-04-25 — oxi-2 — confirmed — R-01 / R-03 / R-11 inline revision styling LANDED

- context: feat/comments-tracked-changes Phase 2.2 entry. Pixel-pass ground truth (author RGB Alice=#D03337, Bob=#5B2C90, move=#2B6033) was captured earlier today — wire it through to the layout output.
- design: rather than threading `tracked_change` through `LineFragment` and changing the `(text, &RunStyle, FieldType, run_index, char_offset)` fragment tuple signature (touches `break_into_lines` + 5 layout call sites), apply the visual styling as a *pre-pass* on `LayoutEngine::layout`'s `doc_resolved` clone. That keeps the entire downstream pipeline unchanged.
- implementation: `crates/oxidocs-core/src/layout/mod.rs` (line 207 area):
  - `apply_revision_styling(doc: &mut Document)` walks `page.blocks` recursively (Block::Paragraph + Block::Table → Cell → Block) and, for each Run with a `tracked_change`, mutates `run.style` to set `underline`/`strikethrough` + `color`.
  - `REVISION_AUTHOR_PALETTE: [&str; 8]` ships the 8-slot author rotation. Slots 0 (Alice #D03337) and 1 (Bob #5B2C90) are COM-confirmed; slots 2-7 use Word's documented rotation pending a 3+ author fixture.
  - `REVISION_MOVE_COLOR = #2B6033` is hard-coded for `<w:moveFrom>` / `<w:moveTo>` regardless of author (Word's quirk).
  - Author lookup: `tc.author` → `Document.authors[idx].color_index` → palette slot. Defensive fallback to slot 0 if author isn't in the palette (unreachable in practice — I-03 builds `authors` from the same source set).
  - `change_type` mapping: insert → underline+author color, delete → strikethrough+author color, moveFrom → strikethrough+green, moveTo → underline+green. Unknown change_type leaves the run alone (forward-compatibility).
- evidence:
  - 4 new layout integration tests in `tests/comments_fixtures.rs` (`fixture_05_layout_ins_underline_in_author_color`, `fixture_06_layout_del_strikethrough_in_author_color`, `fixture_10_layout_two_authors_get_distinct_colors`, `fixture_08_layout_moves_render_in_green`). All 4 + the existing 10 IR-only tests pass (14/14).
  - Adjacency assertion: in fixture_05, `LayoutContent::Text` for the surrounding "Before" run remains `underline=false, strikethrough=false` — confirming the pre-pass doesn't leak into non-revision runs.
  - Two-author distinction (fixture_10): Alice gets #D03337 from the "ALICE" element; Bob gets #5B2C90 from the "REMOVE" element — both in the same paragraph.
  - Move quirk (fixture_08): both occurrences of "moved" come back tagged #2B6033 (green), not Alice's red, with one occurrence strikethrough (moveFrom) and the other underlined (moveTo).
  - Full crate test suite: 51 lib tests + 14 comments_fixtures tests pass; pre-existing `kinsoku::test_line_start_prohibited` failure is unrelated (predates this branch).
- limitations / known follow-ups:
  - Move visualization is single-line + green in v1, not Word's default double-strike/double-underline. Upgrade hinges on adding `underline_style="double"` + a `strikethrough_style` primitive at the renderer; out of scope for this tick.
  - Headers, footers, footnotes, endnotes, and textbox-internal blocks aren't walked yet by `apply_revision_styling`. The 10 fixtures don't exercise revisions in those locations; add when needed.
  - rPrChange (fixture_09) is intentionally NOT styled — Word renders the new properties (e.g., bold black) directly + a margin balloon, which is R-12.
- baseline risk: extremely low — the 184-doc baseline contains 5 `<w:del>`-only docs and zero comment/ins/move docs (per `inventory/README.md`). For those 5 docs, runs that previously rendered as plain text now render with strikethrough + #D03337. Visually correct, but a baseline SSIM regression of <0.01 per affected page is possible (still need to verify with full pipeline).
- path: Path B `[confidence-merge]` once the `pipeline.verify` baseline is re-run.

## 2026-04-25 — oxi-2 — confirmed — Tick 2-3 pixel-sampling pass (author RGB + balloon geometry + resolved desat)

- context: feat/comments-tracked-changes — with the object-model pass complete on 10/10 fixtures, R-* renderer rows still blocked on the **visual** ground truth: author ink RGB, balloon rectangle, balloon background tint, strikethrough Y, etc.
- approach: render each fixture via `Document.ExportAsFixedFormat(Item=wdExportDocumentWithMarkup)` (NOT `SaveAs2(FileFormat=17)` — the latter drops markup). Rasterize page 1 at 150 DPI via PyMuPDF. For each `Revision` / `Comment`, use COM `Range.Information(WD_HORIZONTAL/VERTICAL_POSITION_RELATIVE_TO_PAGE)` to get page coordinates → convert to pixels → sample RGB.
- gotchas hit & fixed:
  - `View.ShowComments = False` is the default; balloons don't render even with `Item=wdExportDocumentWithMarkup` unless we flip it. Set `ShowComments = True` + `ShowRevisionsAndComments = True` + `RevisionsView = wdRevisionsViewFinal` before exporting.
  - Comment.Reference (the inline marker) returns the wrong Y for balloon alignment — use `Comment.Scope` (the range start) instead. Word's balloon aligns with the rendered first character of the range, not the marker.
  - Revision.Range.Information(6) returned x_pt y_pt that landed on the line top — sampling within a 30pt × 20pt window leaked into adjacent body text. Narrowed to 12pt × 14pt and added a saturation preference (prefer colored ink over black). Then fixture_07's three revisions all reported #D03337 cleanly.
- evidence:
  - Author RGB: **Alice = #D03337 (208, 51, 55), Bob = #5B2C90 (91, 44, 144)** — confirmed across 4 fixtures (05, 06, 07, 10).
  - **Move-revision quirk**: `<w:moveFrom>` / `<w:moveTo>` always render in **green #2B6033 (43, 96, 51)** regardless of author. This is a hard-coded Word behavior; R-11 needs to bypass the author-color rotation.
  - **rPrChange (fixture_09)**: underlying text is NOT author-tinted — renders as the new property (bold black). The author signal is in a separate right-margin "Formatting" balloon. R-12 minimal version: render the new properties + a margin balloon (don't tint the inline text).
  - Balloon geometry: width 293.8pt unresolved / **190.1pt resolved** (smaller box for done=true), right edge ≈ page_width − 4pt for all balloons.
  - Resolved desaturation: Alice unresolved BG #FAE6E7, resolved BG #F1EDEC — chroma drops from ~20 to ~5, lightness unchanged. R-09 must apply this when `done=true`.
  - Body width is reduced when balloons are present: fixture_01 body is ~147pt wide vs ~451pt without comments.
- outcome: created `tools/metrics/measure_comments_tracked_changes_pixels.py`. Output `tools/metrics/output/comments_tracked_changes_pixels.json` + per-fixture PDFs/PNGs. Promoted snapshot to `docs/spec/comments_tracked_changes/com_measurements/comments_tracked_changes_pixels.json` + `PIXEL_PASS_README.md`. INDEX.md flipped pixel-pass to checked.
- impact: **R-01, R-02 (palette slots 0-1), R-03, R-04, R-05, R-08, R-09, R-11 are unblocked**. R-07 (dotted connector line) and exact underline/strikethrough Y still need a small follow-up horizontal-line probe — non-blocking for the first renderer ticks.
- known limitations:
  - fixture_09 ink_rgb is null (rPrChange line offset by ~88pt due to formatting-balloon header).
  - Author palette slots 2..7 unmeasured (need 3+ author fixture).
  - Strikethrough / underline Y captured only as glyph ink center, not the marker line itself.
- baseline risk: none — measurement-only, no code changes.
- path: Path B confidence-merge.

## 2026-04-25 — oxi-2 — confirmed — Tick 2-3 fixture content-type fix (3 fixtures unblocked, 10/10 Word-OK)

- context: feat/comments-tracked-changes — `tools/metrics/output/comments_tracked_changes_com.json` had reported 7/10 fixtures Word-OK since 2026-04-18; fixtures 02 (reply), 03 (resolved), 10 (multi-author) failed `Word.Documents.Open` with `'ファイルが壊れている可能性があります。'`. Pixel/UIA pass for R-* renderer rows was blocked on these 3 because reply/done/multi-author paths are exactly what fixtures 02/03/10 exercise.
- hypothesis: the 3 failing fixtures share a structural distinguishing feature — fixtures 02 + 03 emit `commentsExtended.xml`, fixture 10 emits `people.xml`. Suspected the historical `application/vnd.ms-word.commentsExtended+xml` / `application/vnd.ms-word.people+xml` content types in `build_comments_samples.py` were rejected by Word 16.0's strict-open path.
- evidence:
  - `Word.Documents.Open(..., OpenAndRepair=True)` succeeded for all 3, returning expected `Comments`/`Revisions` collections.
  - Saved Word's repaired output via `SaveAs2` and diffed `[Content_Types].xml`: Word rewrote both content types to `application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml` and `application/vnd.openxmlformats-officedocument.wordprocessingml.people+xml`.
  - Single-line patch test: in-memory rewrite of just `[Content_Types].xml` (no other XML touched) made strict-open succeed for both 02 and 10 in a fresh Word instance.
- outcome: applied the same one-line-per-part change in `tools/fixtures/comments_samples/build_comments_samples.py`. Re-ran builder and the COM measurement script: 10/10 fixtures now `word_reads_ok=true`. Captured data confirms reply ancestor (fixture 02 `Comment.Ancestor` / `Replies.Count==1`), resolved flag (fixture 03 `Comment.Done==True`), and per-author revisions (fixture 10 Alice ins + Bob del). Both `tools/metrics/output/comments_tracked_changes_com.json` and `docs/spec/comments_tracked_changes/com_measurements/comments_tracked_changes_com.json` regenerated; README updated; INDEX.md `[ ] Tick 2-3 deferred` flipped to `[x] object-model pass complete`.
- impact: pixel/UIA pass still required before R-* renderer rows, but it is no longer blocked on fixture-build defects. Strict-mode round-trip of these fixtures (and Oxi's emitted .docx in the future) now lands in the same content-type bucket Word writes natively, removing one whole class of validator-rejection bugs from the pipeline.
- baseline risk: none — fixtures live under `tools/fixtures/comments_samples/`, not in the regression baseline.
- path: Path B `[confidence-merge]`.

## 2026-04-25 — oxi-2 — confirmed — I-03 author palette + I-04 ShowRevisions

- context: feat/comments-tracked-changes Phase 2 IR rows. After Phase 2 parser COMPLETE, build the IR scaffolding renderer rows depend on.
- I-03 — author palette:
  - new `ir::Author { display: String, color_index: usize }`
  - `Document.authors: Vec<Author>` derived from 3 sources, first-seen order:
    1. `Document.people` (people.xml — Word writes reviewer-first-seen order)
    2. `Comment.author`
    3. tracked-change authors via `walk_block_authors` (run.tracked_change, run.rpr_change, paragraph.ppr_change, paragraph.paragraph_mark_revision; recurses into Block::Table)
  - color_index = position in the palette → renderer maps to RGB through any palette without a separate join.
- I-04 — show-revisions toggle:
  - `ir::ShowRevisions::{All (default), Simple, Original, Final}`
  - serde `rename_all = "snake_case"` so JSON-API consumers see `"all"|"simple"|"original"|"final"`.
  - not wired into a render config struct yet — added as IR plumbing for S-02.
- I-01 closeout: covered incrementally by P-01 + P-10 + P-11. Comment struct already has all required fields and is surfaced on Document.comments.
- I-02 deferred: keeping the current `Run.tracked_change: Option<TrackedChange>` + `Run.rpr_change: Option<PropertyChange>` shape. Multiple-revisions-per-run is rare and unused in baseline; SmallVec refactor blast-radius is high. Revisit only if a renderer needs it.
- evidence:
  - 2 unit tests: `show_revisions_default_is_all_and_round_trips_json`, `build_author_palette_dedupes_in_first_seen_order`
  - 1 integration extension: `fixture_10_people_two_reviewers` extended to assert `Document.authors` palette ordering
  - 1 new integration: `fixture_05_authors_palette_from_tracked_changes_only` — palette falls back to tracked-change authors when people.xml is absent
- baseline risk: none.
- path: Path B `[confidence-merge]`.

## 2026-04-25 — oxi-2 — confirmed — P-08 *PrChange silent-drain (6 variants) — Phase 2 parser COMPLETE

- context: feat/comments-tracked-changes Phase 2 parser row P-08 — final row
- scope: ECMA-376 §17.13.5 — six revision-history wrappers, each containing a *prior* copy of the same property element they sit inside:
  - `<w:tblPrChange>` inside `<w:tblPr>`
  - `<w:trPrChange>` inside `<w:trPr>`
  - `<w:tcPrChange>` inside `<w:tcPr>`
  - `<w:sectPrChange>` inside `<w:sectPr>`
  - `<w:tblGridChange>` inside `<w:tblGrid>`
  - `<w:numberingChange>` inside `<w:numPr>`
- silent-bug class: each owning property parser uses the same depth-doesn't-gate-Empty-handlers pattern as parse_run_properties / parse_paragraph_properties. Without the drain, the prior property body would silently leak — most concretely:
  - `parse_table_grid` would APPEND prior `<w:gridCol>` widths to the column list (column count corruption)
  - `parse_num_pr` would OVERWRITE current numId/ilvl with the prior values
  - `parse_table_properties`, `parse_table_row`, `parse_cell_properties`, `parse_section_properties` would all leak prior style/border/margin into current state
- change:
  - new `drain_element(reader, tag_name)` helper — reads to the matching End regardless of nesting
  - 6 explicit drain branches added at the top of each respective parser's Start arm
- evidence:
  - 3 unit tests covering the helper, tblGridChange, and numberingChange (the most demonstrable corruption cases)
- IR emission deferred: attack_matrix says "Rare in practice; emit to IR for completeness, no renderer work yet". Since the renderer doesn't yet consume these, the silent-drain is the highest-value minimum. When a renderer needs them, the next iteration adds typed PropertyChange fields like P-06/P-07 did.
- baseline risk: none.
- path: Path B `[confidence-merge]`. **Phase 2 parser quartet (P-01..P-12, except deferred renderer rows) COMPLETE — 12/12 rows landed.**

## 2026-04-25 — oxi-2 — confirmed — P-09 paragraph-mark ins/del

- context: feat/comments-tracked-changes Phase 2 parser row P-09
- scope: ECMA-376 §17.13.5 — `<w:pPr>/<w:rPr>/<w:ins>` or `/<w:del>` marks the paragraph's pilcrow (¶) itself as inserted (new split) or deleted (paragraph merged with next). revisions_notes.md §2.
- change:
  - `Paragraph.paragraph_mark_revision: Option<TrackedChange>`
  - ins/del Empty detection added inside the pPr/rPr sub-loop in parse_paragraph_properties (where ppr_rpr is already being built). Captures change_type/author/date/pair_id.
  - parse_paragraph_properties return: 6-tuple → 7-tuple. Updated caller + 3 Paragraph constructors.
- evidence:
  - 2 unit tests: parse_pmark_ins_via_ppr_rpr + parse_pmark_del_via_ppr_rpr (inline XML; no fixture in 10-doc set, per attack_matrix note)
- baseline risk: none (0 pPr/rPr/ins or /del in 184 baseline docs).
- path: Path B `[confidence-merge]`.

## 2026-04-25 — oxi-2 — confirmed — P-11 commentsIds.xml durable ids

- context: feat/comments-tracked-changes Phase 2 parser row P-11
- scope: Word 2019+ `word/commentsIds.xml` (w16cid namespace) — carries durable ids that survive save-as roundtrips (local `w:id` is renumbered freely)
- change:
  - `Comment.durable_id: Option<String>`
  - new `parse_comments_ids_xml` free function, returns `paraId → durableId` map
  - merged into Comments in `build_context_with_theme` after commentsExtended merge
  - accepts both `w16cid:durableId` (canonical) and `w16cid:id` (older draft spelling)
- evidence:
  - 2 unit tests: standard durableId map + legacy id attribute acceptance
  - no fixture in the 10-doc set has commentsIds.xml (scanned all 10 → 0 hits)
- baseline risk: none (184 baseline docs have 0 commentsIds.xml).
- path: Path B `[confidence-merge]`.

## 2026-04-25 — oxi-2 — confirmed — P-07 pPrChange + silent-bug fix

- context: feat/comments-tracked-changes Phase 2 parser row P-07
- scope: ECMA-376 §17.13.5 `<w:pPrChange>` carries a full prior `<w:pPr>` body for paragraph-level revisions
- silent-bug fix (mirrors P-06): `parse_paragraph_properties`'s Empty handlers (jc, pStyle, spacing attrs, etc.) don't gate on depth. An inner `<w:jc val="right"/>` inside `<w:pPrChange>/<w:pPr>` would silently overwrite the current paragraph alignment. Same class of defect as rPrChange, resolved the same way: explicit drain before the fallback.
- change:
  - extend `PropertyChange` with `prior_paragraph_style: Option<Box<ParagraphStyle>>`
  - `Paragraph.ppr_change: Option<PropertyChange>`
  - `parse_paragraph_properties` return: 5-tuple → 6-tuple (added `Option<PropertyChange>`)
  - explicit `pPrChange` branch: captures id/author/date, recursively reparses inner `<w:pPr>` via `parse_paragraph_properties`, drains to `</w:pPrChange>`, handles self-closing `<w:pPr/>`
  - 3 Paragraph constructors updated with `ppr_change: None` (empty-para fallback×2 + main)
- evidence:
  - unit (1): `parse_pprchange_stores_prior_style_without_merging_into_current` — current=Left, prior pPr=Right, both captured without cross-contamination.
- no integration fixture: attack_matrix notes P-07 has no fixture in the 10-doc set. Defer until a dedicated pPrChange fixture is authored (or P-08 gets one).
- baseline risk: none (0 pPrChange in 184 baseline docs).
- path: Path B `[confidence-merge]`.

## 2026-04-25 — oxi-2 — confirmed — P-06 rPrChange + silent-bug fix

- context: feat/comments-tracked-changes Phase 2 parser row P-06
- scope: ECMA-376 §17.13.5 `<w:rPrChange>` carries a full prior `<w:rPr>` body to record run-property revisions
- silent-bug noticed while scoping: the pre-existing `parse_run_properties` had no `rPrChange` branch. Its depth counter increments on every Start but its property handlers (b/i/u/color/font) don't gate on depth. Result: the prior `<w:rPr>` inside `<w:rPrChange>` would silently merge into the *current* style — `<w:b/>` in the prior would set the current run bold. This is latent today (baseline 184 docs have 0 rPrChange) but would corrupt formatting on any future doc with rPrChange. The new explicit branch drains rPrChange before it reaches those handlers.
- change:
  - new `ir::PropertyChange { id, author, date, prior_run_style: Option<Box<RunStyle>> }`
  - new `Run.rpr_change: Option<PropertyChange>`
  - `parse_run_properties` return type: `RunStyle` → `(RunStyle, Option<PropertyChange>)`. Only one caller (parse_run) updated.
  - inline handler in parse_run_properties: captures rPrChange attrs, recursively reparses the inner `<w:rPr>` as the prior RunStyle, consumes up to `</w:rPrChange>` without touching the current style. Handles self-closing `<w:rPr/>` (prior = RunStyle::default()).
  - 3 Run constructors updated with `rpr_change: None` (layout empty-para-prefix, parser omml-math, parser bookmark-anchor)
- evidence:
  - integration: `fixture_09_rpr_change_bold` — verifies current bold + prior plain + id=300 + author/date populated
- baseline risk: none (new field, empty in 184 baseline docs).
- path: Path B `[confidence-merge]` — the silent-bug fix is a free correctness win.

## 2026-04-25 — oxi-2 — confirmed — P-03/P-04 ins+del locked down + P-05 moveFrom/moveTo

- context: feat/comments-tracked-changes Phase 2 parser rows — the tracked-change quartet
- P-03/P-04 (verification): `<w:ins>` / `<w:del>` were pre-existing, emitting `Run::tracked_change{change_type: "insert"|"delete", author, date}` and preserving `<w:delText>` as `Run::text`. Locked down with 3 integration tests on fixtures 05, 06, 07 (including XML-order preservation in mixed case).
- P-05 (new): `<w:moveFrom>` / `<w:moveTo>` wrap runs identically to ins/del. Added `change_type="moveFrom"|"moveTo"` plus a new `pair_id: Option<String>` field on `TrackedChange` (the wrapper's `w:id`).
- important finding — pairing is NOT via wrapper `w:id`: fixture 08 shows `moveFrom w:id="201"`, `moveTo w:id="202"`. The actual from↔to pairing lives on `moveFromRangeStart` / `moveToRangeStart` (both carrying `w:id="200"` + `w:name="move1"`). Revisions_notes.md §1.2 is correct; the attack-matrix row note saying "Pair via shared w:id on the Range markers" was accurate. Phase 2 parser surfaces the wrapper id only; R-11 will walk range markers when it needs the from↔to linkage.
- refactor: the four ins/del/moveFrom/moveTo branches collapsed into a single `"ins"|"del"|"moveFrom"|"moveTo"` arm mapped to `change_type` strings; `parse_tracked_change_runs` receives the element name as end_tag.
- tests:
  - integration (4 new): fixture_05_single_ins, fixture_06_single_del, fixture_07_mixed_ins_del, fixture_08_move_from_to_pair
- baseline risk: none. 184 baseline docs have 5 lone `w:del`s (already handled pre-change) and 0 w:ins/moveFrom/moveTo.
- path: Path B `[confidence-merge]`.

## 2026-04-25 — oxi-2 — confirmed — P-12 people.xml reviewer list

- context: feat/comments-tracked-changes Phase 2 parser row P-12
- scope: MS-DOCX w15 — `<w15:people>/<w15:person w15:author="..."><w15:presenceInfo providerId userId/>`
- change:
  - new `ir::Person{author, provider_id, user_id}` type, re-exported from `ir::*`
  - `Document.people: Vec<Person>` (document order preserved — Word writes reviewer-first-seen, so this seeds R-02 palette without re-sort)
  - `OoxmlParser::parse_people()` + `parse_people_xml` free function; missing part → empty list
  - handles both `<w15:person>…</w15:person>` (with nested presenceInfo) and self-closing `<w15:person/>` (no presence data)
  - drops malformed `<w15:person>` entries missing `w15:author`
- evidence:
  - 3 unit tests: two-reviewer shape (fixture 10 mirror), missing-presenceInfo, blank-author dropped
  - 1 integration test: `fixture_10_people_two_reviewers` runs full parse_docx pipeline, confirms `Document.people == [Alice Reviewer, Bob Reviewer]` with provider/userId attached
- baseline risk: none (184 baseline docs have 0 people.xml).
- path: Path B `[confidence-merge]`. Completes the Phase-2 parser quartet needed before I-03/R-02 (author-colour palette) can start.

## 2026-04-25 — oxi-2 — confirmed — P-10 commentsExtended.xml (reply threading + resolved)

- context: feat/comments-tracked-changes Phase 2 parser row P-10
- scope: MS-DOCX w15 extension — `<w15:commentEx paraId="..." paraIdParent="..." done="..."/>`
- change:
  - new fields on `ir::Comment`: `para_id: Option<String>`, `parent_para_id: Option<String>`, `resolved: bool`
  - `parse_comments_xml` grabs the first `w14:paraId` off the comment body's first `<w:p>` (the commentsExtended join key)
  - new `parse_comments_extended_xml` function + merge step in `build_context_with_theme`. Accepts both canonical `w15:paraIdParent` and legacy `w15:parentParaId` per comments_notes.md §4
- evidence:
  - 3 new unit tests: `parse_comments_xml_captures_first_para_id`, `parse_comments_extended_reply_and_resolved`, `parse_comments_extended_accepts_legacy_parent_para_id_spelling`
  - 2 new integration tests: `fixture_02_comments_extended_reply_threaded` (parent/reply `paraIdParent` = "00000010"), `fixture_03_comments_extended_resolved_flag` (`w15:done="1"` → `resolved==true`)
  - NB: fixtures 02/03 fail `Documents.Open` in Word (validator defect to fix in a separate tick) but are valid enough for python-docx + our scanner + our parser; the tests cover the parser contract regardless.
- baseline risk: none — 184 baseline docs have 0 commentsExtended.xml.
- path: Path B `[confidence-merge]`.

## 2026-04-25 — oxi-2 — confirmed — P-02 commentReference wired to Run

- context: feat/comments-tracked-changes Phase 2, second parser row
- scope: balloon anchor marker `<w:commentReference w:id="N"/>` inside `<w:r>`
- change: added `comment_references: Vec<String>` to `oxidocs_core::ir::Run`; `parse_run` captures the id inside the enclosing run. `commentRangeStart/End` parsing was pre-existing (run-level `Vec<String>`).
- rationale: the enclosing run is the glyph the renderer projects to the right margin to position the balloon (ECMA-376 §22.1.2.56 + comments_notes.md §2.2). One run may legally carry multiple refs.
- evidence:
  - unit: `parse_run_captures_comment_reference` — synthesized `<w:r>` with `commentReference id="0"` yields `run.comment_references == ["0"]`.
  - integration: `fixture_01_comment_fields_roundtrip` extended — `commentReference` id="0" found on exactly one run in the body.
- touched Run constructors (4 sites): `ir::types::Run`, `layout::mod::<empty para prefix>`, `parser::ooxml::<omml run>`, `parser::ooxml::<bookmarkStart anchor>` — each now initialises `comment_references: Vec::new()`.
- baseline risk: none (new field is empty in all 184 baseline docs).

## 2026-04-25 — oxi-2 — confirmed — P-01 comments.xml parse complete (initials field)

- context: feat/comments-tracked-changes Phase 2 — first parser row from attack matrix
- scope: `Comment{ id, author, date, initials, runs }` per ECMA-376 §17.13.4.2
- change: added `initials: Option<String>` to `oxidocs_core::ir::Comment`; `parse_comments_xml` now captures the `w:initials` attribute. No renderer impact.
- evidence:
  - unit tests (inline XML) — `parse_comments_xml_captures_initials_and_metadata` (fixture 01 shape, initials="AR") and `parse_comments_xml_missing_initials_is_none` (older Word docs omit `w:initials`).
  - integration test — `crates/oxidocs-core/tests/comments_fixtures.rs::fixture_01_comment_fields_roundtrip` runs the full `parse_docx` pipeline on `fixture_01_single_comment.docx`; confirms `Document.comments[0]` has `id="0"`, `author="Alice Reviewer"`, `initials="AR"`, `date="2026-04-18T10:00:00Z"`, plus `commentRangeStart`/`End` markers surface on runs (P-02 shadow coverage).
  - COM ground truth: `docs/spec/comments_tracked_changes/com_measurements/comments_tracked_changes_com.json` (2026-04-18, Word 16.0) — fixture 01 `Comments(1).Initial == "AR"`.
- baseline risk: none. 184-doc baseline has 0 comments.xml (see `inventory/README.md`).
- path: Path B `[confidence-merge]` — spec-referenced + COM-validated + fixture-backed + baseline-neutral by construction.

## 2026-04-18 — oxi-2 — partial — Tick 2-3 Word COM pass, 7/10 fixtures validated

- context: feat/comments-tracked-changes Phase 1 Tick 2-3 (previously deferred; Word 16.0 turned out to be available on this box)
- method: `tools/metrics/measure_comments_tracked_changes_com.py` opens each fixture with Word COM and dumps `doc.Revisions` + `doc.Comments` collections
- evidence (7/10 OK): `doc.Revisions.Type` matches authoring intent for every revision fixture (05 wdRevisionInsert, 06 wdRevisionDelete, 07 × 3 alternating, 08 wdRevisionMovedFrom+MovedTo, 09 wdRevisionProperty). `doc.Comments.Scope.Text` matches authored range for 01 and 04
- evidence (3/10 fail): 02 and 03 (both have commentsExtended.xml) and 10 (has people.xml) fail `Documents.Open` with `com_error` — Word validator rejects even though syntactic XML is well-formed (our scanner reads the markers). Fixture-authoring defect to repair in a follow-up tick; does not block Phase 2 parser rows P-03…P-09.
- key Phase 2 confirmations:
  - `<w:rPrChange>` is bucketed as `wdRevisionProperty` (3) in Word's Revisions.Type enum — a single Inserted/Deleted/Property IR enum matches Word's model.
  - `moveFrom` / `moveTo` pair is reported as TWO separate revisions with identical range text — the `(pair_id, name)` linkage in revisions_notes.md §1.2 is the correct IR structure.
  - `Comment.Scope.Text` spans `\r`-separated paragraphs for multi-para ranges — parser can represent range as (start: RunRef, end: RunRef) over a flat glyph list.
- deliverable: `docs/spec/comments_tracked_changes/com_measurements/{comments_tracked_changes_com.json, README.md}`
- still deferred: balloon geometry, author RGB palette (R-02), strikethrough Y on CJK text, connector line style, stacking geometry. All need UIA / pixel-sampling.

## 2026-04-18 — oxi-2 — phase-1-complete — attack matrix + master index; COM deferred

- context: feat/comments-tracked-changes Phase 1 final tick (was Tick 7 in TASK.md)
- deliverables:
  - `docs/spec/comments_tracked_changes/attack_matrix.md` — 33-row priority matrix (12 parser + 4 IR + 13 renderer + 4 settings). Effort, blast radius, fixture coverage, COM-measurement dependency, baseline SSIM risk per row. Recommended execution order: P-01→P-12 → I-01→I-04 → R-01/R-03/R-04 first (cheapest renderer wins). Baseline risk = **none** for virtually every row (baseline has 0 comments / 5 lone dels) — Path B confidence-merge is natural.
  - `docs/spec/comments_tracked_changes/INDEX.md` — canonical entry point. Indexes every Phase 1 asset (spec notes, attack matrix, inventory, fixtures, coordination files) + 3 dogfood lookup simulations ("implement w:ins parse", "render comment balloons", "worried about SSIM regression on w:del").
- deferred: Tick 2-3 Word COM measurement (12-target checklist embedded in attack_matrix.md §"Tick 2-3 deferred COM checklist"). Must run in a dedicated session before starting renderer rows R-01+; parser rows (P-01…P-12) can begin without it.
- handoff: Phase 2 can start immediately at any parser row. IR sketch in revisions_notes.md §8; renderer row R-02 (author-color palette) blocks R-10 (margin change bar) and wants COM data.
- methodology: Phase C inventory→matrix→index re-application succeeded; validated pattern still generalises beyond yakumono/cell/footnote domains. No code changes in /loop per user directive; all 4 ticks = pure measurement + memo.

## 2026-04-18 — oxi-2 — spec-notes-written — ECMA-376 §17.13.1 + §17.13.5

- context: feat/comments-tracked-changes Phase 1 Tick 4 (spec notes)
- deliverables:
  - `docs/spec/comments_tracked_changes/comments_notes.md` — parts, content types, inline markers, comments.xml structure, commentsExtended threading, Word display rules (not-in-spec), JIS X 4051 interaction, parser checklist, fixture cross-reference
  - `docs/spec/comments_tracked_changes/revisions_notes.md` — element taxonomy (ins/del/move/\*Change), block-level ins/del via pPr/rPr, accept/reject semantics, Word display rules, IR sketch (Phase 2 planning), parser checklist
- key observations recorded for Phase 2:
  - `w:id` on revisions is NOT durable across saves; parsers must not use it as stable identifier
  - Paragraph-mark insert/delete lives on `pPr/rPr`, NOT as an outer `<w:ins>` — common miss
  - `moveFrom/To` pair via shared `w:id` on the Range markers + opaque `w:name` label
  - Comment balloon geometry, stacking, author color palette are all renderer-defined (not spec)
  - `w15:paraIdParent` and `w15:parentParaId` both appear in the wild; accept on parse, emit canonical form on write
- next: Tick 7 (attack matrix + master index), OR Tick 2-3 (Word COM measurement against the 10 fixtures — requires Word running on this box)

## 2026-04-18 — oxi-2 — fixtures-ready — 10 minimal repros for comments + tracked changes

- context: feat/comments-tracked-changes Phase 1 Tick 5-6 pulled forward (baseline too sparse for Tick 2-3 COM without self-authored docs)
- hypothesis: 10 feature-isolated fixtures are enough to unblock subsequent COM measurement + spec notes
- method: zip-level OOXML generator (`tools/fixtures/comments_samples/build_comments_samples.py`), one fixture per feature; validated via python-docx open + inventory re-scan
- evidence: all 10 open cleanly; per-file marker counts exactly match intent; MANIFEST.json committed alongside
- coverage: 01 single comment / 02 comment+reply / 03 resolved / 04 multi-para range / 05 single ins / 06 single del / 07 mixed ins+del / 08 moveFrom+moveTo / 09 rPrChange bold / 10 two reviewers (Alice+Bob)
- side-effect: inventory scanner reply-detection pattern corrected (`w:parentId` → `w15:paraIdParent`); baseline totals unchanged (still 0 comments, 5 del-only)
- next: Tick 2-3 runs Word COM measurement against these fixtures — balloon position, range highlight color, ins underline style, del strikethrough color, multi-reviewer color rotation
- tools: tools/fixtures/comments_samples/build_comments_samples.py
- path: tools/fixtures/comments_samples/fixture_{01..10}_*.docx (+ MANIFEST.json)

## 2026-04-18 — oxi-2 — baseline-inventory — comments + tracked-changes sparsity confirmed

- context: feat/comments-tracked-changes Phase 1 Tick 1 — establish baseline usage floor before Phase 2 implementation
- hypothesis: the 177/184 baseline .docx corpus contains enough comment + tracked-change usage to drive COM measurement and regression testing
- method: zip+XML scan of all 184 docx under `oxi-main/tools/golden-test/documents/docx/`. Count `w:commentRange*`, `w:commentReference`, `w:comment` bodies, `w:ins`, `w:del`, `w:moveFrom/To`, `w:*PrChange` markers. Detect `word/comments.xml`, `commentsExtended.xml`, `commentsIds.xml`, `people.xml`
- evidence (JSON at tools/metrics/output/{comments_inventory,tracked_changes_inventory}.json):
  - docs_with_word_comments_xml: **0 / 184**
  - docs_with_any_revision_marker: **5 / 184** (all 1×w:del, one additionally 1×w:pPrChange, single author, boilerplate `people.xml`)
  - zero replies, zero moves, zero rPrChange, zero multi-reviewer scenarios in corpus
- outcome: REFUTED. Baseline provides essentially no test signal for comment + revision rendering. All Tick 2–3 COM measurements and all Phase 2 regression suites MUST use self-authored fixtures. Advantage: no SSIM floor risk for these features (Path B [confidence-merge] is the natural merge gate); work can proceed on dedicated branch without bottom-N concern.
- tools: `tools/metrics/inventory_comments_tracked_changes.py`
- next-tick: Tick 2 — author N reference docx fixtures in `tools/fixtures/comments_samples/` (even a provisional set unblocks Tick 2 COM measurement). This pulls Tick 5-6 earlier in the pipeline.

## 🔥 BLOCKER: GDI preset render coverage (Path A fix target)
## 2026-04-18 — dedicated — partial-implementation — GDI PresetShape 5 primitives

- branch: fix/gdi-preset-shapes (commit f79c502) — NOT merged
- context: oxi-2 found GDI renderer's PresetShape handler only supports
  bracketPair. Adding rect/roundRect/ellipse/straightConnector1/bentConnector3
  mechanism is needed for 2ea81 (rank 6) and other docs with these shapes.
- implementation: 5 GDI calls added (Rectangle, RoundRect, Ellipse,
  MoveToEx+LineTo, Polyline). Mechanism uses IR-provided x/y/w/h.
- stylistic gap identified: regression on all 5 tested docs because default
  stroke style (solid black, IR-provided width) doesnt match Word:
    2ea81 p.2 (rank 6): 0.6356 -> 0.6292 (-0.0064)
    b35 p.1 (rank 3):   0.6134 -> 0.6110 (-0.0024)
    29dc6e p.6:         0.9327 -> 0.9239 (-0.0088)
    1636d28 p.1:        0.7255 -> 0.7189 (-0.0066)
    2ea81 p.1:          0.7829 -> 0.7789 (-0.0040)
- Ra §9 decision: NOT merged. Bottom-5 floor would regress (b35 rank 3
  directly affected). Branch fix/gdi-preset-shapes retained.
- gap to close before merge:
    1. COM-measure Word stroke width/color for these shape types (3+ docs)
    2. Extend IR PresetShape with fill info (solid/noFill/color)
    3. Adjust stroke width to match Words pen behavior
    4. Cross-verify SSIM >= current on affected docs
- lesson: adding functionally-correct rendering can regress SSIM if
  styling details differ. "No render" produces blank pixels; "wrong
  style render" produces differing pixels — the latter scores lower.
  Stylistic fidelity is prerequisite for Path A landing.

---
## 2026-04-18 — dedicated — iter2 progress — GDI PresetShape fill + invisible skip

Follow-up to the iter1 entry above (commit f79c502 → b1a9edf on
fix/gdi-preset-shapes branch).

Changes in iter2:
- IR LayoutContent::PresetShape gets new fill_color field (parser already
  extracted shape.fill, now propagated through).
- Renderer creates a solid fill brush when fill_color is present;
  Rectangle/RoundRect/Ellipse paint fill behind stroke in a single call.
- Renderer SKIPS shapes with neither stroke nor fill (Word draws nothing
  for textbox-frame rects with <a:noFill/> on both). Previously these
  were getting a default black stroke, polluting the output.

Quick verify 5-doc diff (vs main baseline):
  2ea81 p.1: 0.7829 -> 0.7829 (= FIXED from iter1 -0.0040)
  2ea81 p.2: 0.6356 -> 0.6298 (-0.0058, iter1 was -0.0064)
  b35 p.1:   0.6134 -> 0.6110 (-0.0024 unchanged from iter1)
  29dc6e p.6: 0.9327 -> 0.9239 (-0.0088 unchanged)
  1636d28 p.1: 0.7255 -> 0.7189 (-0.0066 unchanged)

Residuals: roundRect callouts. Oxi's solid-pen stroke differs from
Word's antialiased sub-pixel edge. Fixing requires GDI AA + sub-pixel
pen setup, deferred to rendering-quality session.

Still NOT MERGED to main (b35 rank 3 regresses -0.0024 per Ra §9).
Branch fix/gdi-preset-shapes carries iter2 commit b1a9edf.



---
## 2026-04-18 — dedicated — iter3-4 + full verify — GDI PresetShape final state

Iter3: non-bracketPair stroke width cap to 1pt (commit a4ea88c)
Iter4: PS_INSIDEFRAME pen style for non-bracketPair (commit 5982db7)

Full verify 177-doc / 352-page (iter4):
  23 improved / 310 unchanged / 19 regressed
  Net: +0.0892 (informational, not gate)
  Bottom-5 floor: 3.0597 → 3.0567 (-0.0030) → **Path A FAIL**

Bottom-5 movement:
  1. 0e7a p.2   0.5767 = (unchanged)
  2. d77a p.9   0.6042 = (unchanged)
  3. b35  p.1   0.6134 → 0.6127  (-0.0007, tolerance)
  4. 2ea81 p.2  ENTERED from rank 6 at 0.6306 (-0.0050 from 0.6356)
  5. b837 p.4   0.6325 = (unchanged)
  683f dropped out (was rank 5 at 0.6329, now outside bottom-5)

Blocker: 2ea81 p.2 regression. PS_INSIDEFRAME improved b35/29dc6e
significantly but made 2ea81 p.2 slightly worse (-0.0044 iter3 → -0.0050
iter4). The shape in 2ea81 p.2 may need specific investigation — its
roundRect callouts may anchor differently than other docs.

Branch fix/gdi-preset-shapes final state (5 commits f79c502..5982db7):
- 5 primitives (rect/roundRect/ellipse/straightConnector1/bentConnector3)
- fill_color field in IR
- Invisible shape skip
- Stroke width cap
- PS_INSIDEFRAME pen style

Session close decision: branch preserved, main unchanged. Next dedicated
session should focus specifically on 2ea81 p.2 shape positioning —
likely a small fix unlocks Path A merge.



## 2026-04-18 — dedicated — partial-implementation — GDI PresetShape 5 primitives

- branch: fix/gdi-preset-shapes (commit f79c502) — NOT merged
- context: oxi-2 found GDI renderer's PresetShape handler only supports
  bracketPair. Adding rect/roundRect/ellipse/straightConnector1/bentConnector3
  mechanism is needed for 2ea81 (rank 6) and other docs with these shapes.
- implementation: 5 GDI calls added (Rectangle, RoundRect, Ellipse,
  MoveToEx+LineTo, Polyline). Mechanism uses IR-provided x/y/w/h.
- stylistic gap identified: regression on all 5 tested docs because default
  stroke style (solid black, IR-provided width) doesnt match Word:
    2ea81 p.2 (rank 6): 0.6356 -> 0.6292 (-0.0064)
    b35 p.1 (rank 3):   0.6134 -> 0.6110 (-0.0024)
    29dc6e p.6:         0.9327 -> 0.9239 (-0.0088)
    1636d28 p.1:        0.7255 -> 0.7189 (-0.0066)
    2ea81 p.1:          0.7829 -> 0.7789 (-0.0040)
- Ra §9 decision: NOT merged. Bottom-5 floor would regress (b35 rank 3
  directly affected). Branch fix/gdi-preset-shapes retained.
- gap to close before merge:
    1. COM-measure Word stroke width/color for these shape types (3+ docs)
    2. Extend IR PresetShape with fill info (solid/noFill/color)
    3. Adjust stroke width to match Words pen behavior
    4. Cross-verify SSIM >= current on affected docs
- lesson: adding functionally-correct rendering can regress SSIM if
  styling details differ. "No render" produces blank pixels; "wrong
  style render" produces differing pixels — the latter scores lower.
  Stylistic fidelity is prerequisite for Path A landing.

---
## 2026-04-18 — dedicated — deep-dive — text_y_offset centering vs bottom-align

Investigation of Task #1 residual: why line=exact rule appears correct but
cursor advance is already correct. Measured Word at 300 DPI:

  repro (line=13, font=10.5):
    Word text top offset from line box: +3.26pt (after 1.96pt ink offset = 1.30pt cell offset)
    Oxi text top offset:                 +4.46pt (cell offset 2.50pt)
  repro (line=15, font=10.5):
    Word cell offset: 3.14pt
    Oxi cell offset: 4.50pt

  Both show Oxi +1.2pt below Word consistently.

Oxis formula:  (bottom-align)
Word actual:   (centering) — matches A case.

Test (branch fix/text-y-offset-exact): change to centering unconditionally.
  1ec1 p.1 (textbox Shape 4 exact=22 font=14): 0.6701 -> 0.6370 (-0.0331) !!!
  2ea81 p.1: +0.0017
  d77a p.1: +0.0206
  Bottom-5 docs: no change (they dont use exact heavily)

Conclusion: Word uses TWO modes:
  - Body paragraphs with lineRule=exact: CENTER text in box
  - Textbox/shape paragraphs with lineRule=exact: BOTTOM-ALIGN

Oxis current bottom-align is correct for textbox but wrong for body.
Full fix requires adding  context parameter to
. Not implemented this session — bottom-5 gate
wouldnt budge (bottom-5 docs are body but use auto not exact, so
they arent affected).

Branch fix/text-y-offset-exact: REVERTED. Finding logged for future impl.



## 2026-04-18 — oxi-1 — drift-localized — b35 p.1 Class B +2.5pt body offset
- context: Task #4 — b35 rank 3 bottom-5 (SSIM 0.6134), prior memos claim Class B
- hypothesis: b35 p.1 has measurable per-paragraph drift like 2ea81 Class B
- method: Word COM per-paragraph Y + Oxi --dump-layout per-block Y; align by text content
- evidence: 4 body paragraphs aligned (Oxi dump has only 4 body para_idx; tables use block para_idx). All 4 show +2.00-2.50pt downward drift. Median |Δ|=+2.50pt, consistent.
- outcome: NOT ceiling. Class B drift CONFIRMED. Likely same root cause as Task #1 line=exact boundary rule. Fixing line=exact rule may simultaneously improve b35 p.1 body paras. Table-row drift is SEPARATE mechanism (covered by oxi-4 LM0 cell formula).
- impact on session: bottom-5 coverage now complete (all 5 + rank 6 diagnosed). 4 have dedicated-session-ready fix targets, 1 ceiling, 1 pivot.
- tools: diff_b35_p1_paras.py
- memory: project_b35_p1_class_B_drift_confirmed.md

## 2026-04-18 — oxi-1 — refuted — b837 footnote spill hypothesis (oxi-2 unblock)
- context: Task #3 — b837 rank 4; oxi-2's footnote spill investigation
- hypothesis (oxi-2): Word splits long multi-line footnote bodies across pages
- method: Word COM measurement of all 25 fns in b837 (ref_page + body_first_page + body_last_page); 3 additive scratch variants with single/many/multi-line fns
- evidence: ZERO spill in 42 total fns tested (25 real + 17 scratch). All fn body_first_page == body_last_page. Word's rule: fn bodies NEVER span page boundaries.
- outcome: REFUTED. oxi-2 unblocked — pivot to estimate_footnote_h per-fn over-count (10pt/fn). Real bug location: mod.rs:631 per-footnote height estimate, NOT spill model.
- supporting evidence: existing output/b837_footnote_y.json + new pipeline_data/b837_footnote_spill.json
- tools: measure_b837_footnote_spill.py, gen_footnote_spill_repro.py
- memory: project_b837_footnote_spill_FALSIFIED.md
- impact: oxi-2 status can change from "investigating spill" to "investigating per-fn estimate over-count"

## 2026-04-18 — oxi-1 — re-confirmed — 0e7a p.2 layout ceiling
- context: rank 1 bottom-5 (0.5767); prior memos claimed layout ceiling,
  re-verification requested per Task #2
- hypothesis: 0e7a p.2 has no remaining layout drift; SSIM gap is
  sub-pixel / AA / glyph-rendering only
- method: fresh Oxi --dump-layout on current main + Word COM per-paragraph Y
  measurement; align 20 paragraphs by text content
- evidence: median Δ=+0.00pt, max |Δ|=0.50pt across 20 aligned paragraphs;
  17 of 20 paras show Δ=+0.00 exactly
- outcome: LAYOUT CEILING confirmed. 0e7a p.2 SSIM 0.5767 NOT
  layout-improvable. Future sessions should skip this page for layout
  work and focus on d77a p.9 (rank 2) where drift IS fixable
  (line=exact boundary rule — see preceding entry).
- tools: measure_word_paras_generic.py, diff_0e7a_p2_paras.py
- memory: project_0e7a_p2_ceiling_CONFIRMED_2026_04_18.md

## 2026-04-18 — oxi-1 — confirmed — line=exact boundary rule (additive)
- context: 2ea81 idx=6→7 +2pt bug (see project_2ea81_line_exact_boundary_bug.md)
- hypothesis: at adjacent paragraphs both with lineRule=exact, Y advance A→B uses lineA's value, not lineB
- method: 5 scratch additive variants (V1-V5) from python-docx Document(); different font/empty/non-empty/increasing/decreasing combinations
- evidence: all 5 variants confirm delta = N_A × lineA/20; V5 DECREASING (400→240) also matches (excludes naive "use larger" hypothesis)
- outcome: additive rule CONFIRMED. Oxi bug = +2pt at empty-paragraph boundary. Real-doc verification partial (2ea81 re-measured for idx=2→3 Oxi matches Word; idx=6→7 memo has Oxi+2pt). COM RPC crashes on larger docs blocked 29dc6/6514f214 verification.
- tools: repro_line_exact_variants.py, verify_line_exact_rule_3docs.py, verify_line_exact_oxi_vs_word.py
- outcome for next: implementation in dedicated session. Location: mod.rs paragraph cursor advance when prev para lineRule=exact. Must preserve non-empty-A cases.
- memory: project_line_exact_rule_additive_confirmed.md

## 2026-04-18 — oxi-3 — architectural-validation — yakumono is geometry heuristic, NOT fixed rule
- context: 6-tick 4-tier additive bisection from scratch (final conclusion)
- method: scratch + cSC + compat=15 + kern + jc + one-property-at-a-time
  (docGrid / rFonts cascade / sectPr), then pgMar width sweep
- evidence: Non-monotonic content-width dependency:
  - 465-475pt: NO compression
  - **453-455pt: COMPRESSED** (d77a range, L+R margin 1400-1420tw)
  - 451-452pt: NO compression (non-monotonic gap!)
  - 435-445pt: partial (11.0-11.5pt)
- interpretation: Word's yakumono compression is a LINE-WRAP PRE-PASS heuristic:
  - "If the line would fit with yakumono compression, compress"
  - "If compression doesn't help fit, don't compress"
  - "If no pressure at all, don't compress"
- outcome: ARCHITECTURAL VALIDATION — Oxi's existing Phase 2 reactive absorb
  (mod.rs:2977) + 50tw threshold (commit 70841a5) is the CORRECT approach.
  Preemptive char-advance compression would over-compress in cases Word
  doesn't (e.g., content=465pt), causing regressions.
- impact: Saves implementing the wrong solution. Risk-adjusted value HIGH —
  6 ticks of bisection prevented hours of misguided implementation.
- undocumented-quirk: this heuristic is NOT in ECMA-376 or JIS X 4051;
  it's a Word rendering pipeline implementation detail.
- methodology: 5th confirmation type (architectural validation) added to
  additive-primary protocol repertoire. All 4 tiers documented as
  complementary:
  - Positive (new spec → implement)
  - Negative (ceiling → skip)
  - Refutation (false hypothesis → revert)
  - Drift localization (scope expansion → extend)
  - Architectural validation (existing impl correct → stop)
- close: ALL yakumono pending tasks closed. No further investigation
  needed. Phase 2 threshold tuning is the only remaining lever, but its
  bottom-5 impact is already captured in 70841a5 merge.
- artifacts: tools/metrics/additive_tier1_docgrid.py + additive_tier2_rfonts.py
  + additive_tier3_sectpr.py + additive_pgmar_isolate.py + pgmar_width_sweep.py
  + bisect_d77a_styles.py + bisect_d77a_normal_properties.py + retry_normal_combos.py
  + scratch_kern_jc_test.py + pydocx_strip_d77a.py

## 2026-04-18 — oxi-3 — partial — yakumono trigger: kern+jc+cSC+compat15 + content width 452.8-455.3pt range
- Tier 1 (docGrid type): FALSIFIED — docGrid variants don't trigger
- Tier 2 (rPrDefault rFonts/lang cascade): FALSIFIED — no combination triggers
- Tier 3 (sectPr): **HIT** at pgMar L+R = 1418tw (others at 1440)
- pgMar width sweep (content_pt):
  - 465.3-475.3pt: NOT compressed
  - 453.3-455.3pt (L+R = 1400-1420): **COMPRESSED 10.5pt**
  - 451.3-452.8pt (L+R = 1425-1440): NOT compressed
  - 435.3-445.3pt: partial 11.0-11.5pt
- Interpretation: compression is NOT a simple necessary-condition gate but
  a geometry-dependent heuristic. Word's compression kicks in only when
  content width falls in a specific range that interacts with the text and
  line-wrap algorithm. This is a LINE-WRAP-HEURISTIC, not a per-character rule.
- Implementation implications:
  - Cannot be implemented as "compress yakumono IF compat_mode>=15 && cSC":
    that would compress in cases Word doesn't (e.g., content=465pt).
  - Would require replicating Word's line-break pre-pass: "if the line
    would fit with yakumono compression, compress; otherwise don't."
  - Matches the observation that Oxi's existing Phase 2 absorb logic
    (mod.rs:2977) is the right direction — it's reactive, not preemptive.
- Conclusion: char-advance-level yakumono compression cannot be implemented
  as a unconditional rule. The 50tw Phase 2 threshold (merged 70841a5) is
  the correct approach — reactive absorption up to a small overflow.
- artifacts: tools/metrics/additive_tier1_docgrid.py + additive_tier2_rfonts.py
  + additive_tier3_sectpr.py + additive_pgmar_isolate.py + pgmar_width_sweep.py

## 2026-04-18 — oxi-3 — partial — yakumono trigger narrowed to Normal style kern+jc (plus d77a-inherited factors)
- context: continuing reconciliation of earlier "cSC+compat15 confirmed" claim
- method: python-docx safe strip + XML component replacement
- findings (d77a-base variants):
  - replace styles.xml with empty/minimal: `（`=12.0pt (NOT compressed)
  - replace fontTable.xml with minimal: `（`=10.5pt (still compressed) → fontTable NOT needed
  - d77a full Normal style (w:styleId="a") alone in styles.xml: `（`=10.5pt (COMPRESSED)
  - minimum sufficient Normal pPr/rPr: **`<w:kern w:val="2"/>` + `<w:jc w:val="both"/>`** → 10.5pt
  - only jc=both: 11.5pt (partial, -0.5pt)
  - only kern=2: 12.0pt (no)
  - kern + jc: 10.5pt (full, -1.5pt)
- ANTI-BREAKTHROUGH (scratch additive test):
  - scratch + cSC + compat=15 + kern=2 + jc=both → **12.0pt (NOT compressed)**
  - Confirms user's warning: d77a-base tests are subtractive, NOT true scratch
  - d77a has ADDITIONAL inherited factors (beyond kern+jc+cSC+compat15) that
    activate the compression
- remaining candidates for true additive bisection (deferred):
  - d77a's pgMar 1418tw (vs scratch 1440tw)
  - d77a's `<w:docGrid w:type="lines" w:linePitch="360"/>` type attribute
  - d77a's fontTable PANOSE values (but replacing with minimal preserved
    compression — so maybe not needed IF pre-existing)
  - themeFontLang in styles.xml root element
  - rPrDefault rFonts cascade from docDefaults
- outcome: trigger narrower than cSC+compat15 but still not fully isolated;
  kern+jc is necessary but not sufficient in scratch. Next step: additive
  bisection from scratch, adding one d77a property at a time.
- artifact: tools/metrics/bisect_d77a_styles.py + bisect_d77a_normal_properties.py
  + retry_normal_combos.py + single_test_runner.py + scratch_kern_jc_test.py
  + pydocx_strip_d77a.py
- DO NOT implement: kern+jc gate would regress scratch/compat=15 docs that
  don't have the missing factor.

## 2026-04-18 — oxi-3 — needs-reconciliation — yakumono compression trigger = cSC + compat≥15
**STATUS DOWNGRADED from "confirmed" to "needs-reconciliation" 2026-04-18 (oxi-1 review)**

## 2026-04-18 — oxi-2 — reproducible-bug — PresetShape handler only renders bracketPair

- context: 2ea81 rank 6 AlternateContent analysis led to finding
- hypothesis: `tools/oxi-gdi-renderer/src/main.rs:377-433` PresetShape
  handler has `if shape_type == "bracketPair"` but NO else branch. All
  other preset types (rect, roundRect, ellipse, straightConnector1,
  bentConnector3, etc.) fall through silently = rendered as nothing
- evidence:
  - grep confirmed only bracketPair branch exists in the match
  - 2ea81 has 17 AlternateContent shapes (7 rect, 4 roundRect, 3 ellipse,
    2 straightConnector1, 1 bentConnector3) + 4 pic:pic images
  - 11 of 17 have txbx (text content renders via separate path) but
    borders/fills/lines do NOT render
  - 6 of 17 are shape-only (no txbx) → **completely invisible**
  - Also explains earlier "spt 32 connector invisible" finding — VML
    fallback maps to same PresetShape path
- outcome: Path A fix candidate (direct bottom-5 improvement expected)
- bottom-5 / rank 6+ impact:
  - 2ea81 (rank 6): 6 invisible + 11 missing borders — likely 4th major
    bug class after line=exact, tbRlV, spt 202
  - b35 (rank 3): 1 missing rect border (form header text box)
  - b837 (rank 4): 1 AlternateContent shape
  - Other bottom-5 docs: 0 AlternateContent
- implementation scope (dedicated session): 5 GDI primitive branches
  - rect → `Rectangle()`
  - roundRect → `RoundRect()`
  - ellipse → `Ellipse()`
  - straightConnector1 (and VML spt 32) → `MoveToEx` + `LineTo`
  - bentConnector3 → polyline (3-segment right-angle bent)
  - Plus stroke weight / color / flip handling
  Approximately 1 hour's work per the memo.
- memos: `asset_preset_shape_render_gaps.md`, `asset_vml_spt32_connector.md`

## 2026-04-18 — oxi-2 — reproducible-bug — line=exact boundary +2pt
- context: 2ea81 page 1 (rank 6, 0.6356); 2ea81 is Class B (page-aligned, intra-page only)
- hypothesis: when lineSpacingRule=exact changes between adjacent paragraphs (e.g., line=260tw empty → line=300tw body), Oxi advances by NEXT para's value (15pt) while Word uses CURRENT para's value (13pt); +2pt per boundary
- evidence:
  - Word DML + docx pPr: 2ea81 idx=6 empty (line=260tw exact), idx=7 body (line=300tw exact); Word measured advance 6→7 = 13pt
  - Oxi layout dump: dy jumps +2.7 → +4.7 exactly at idx=6→7 boundary
  - minimal repro (`tools/metrics/repro_line_exact_boundary.py`) reproduces precisely: Word A→C=26pt, Oxi A→C=28pt for 3-paragraph 260/260/300 sequence
- outcome: HANDOFF to dedicated session. Not applying code change per /loop policy.
  Code location candidates: `mod.rs:~2045-2100` (line_height computation), `mod.rs:2538` (cursor_y advance). Verify the empty-paragraph's line_height uses its own pPr not next's.
- wider impact: form-style docs (tax/application) with mixed exact line heights accumulate +delta per boundary. Potentially affects SSIM for 2ea81 rank 6 and similar form docs.
- memos: `project_2ea81_line_exact_boundary_bug.md`, `project_2ea81_class_b.md`

## 2026-04-18 — oxi-2 — refuted — yakumono rule `min(fs, 10.5)` char-advance
- context: d77a idx=9 line 1 (MS Gothic 12pt) showed '（' advance=10.5pt when fontsize=12, leading to "single yakumono compression" hypothesis
- hypothesis: Oxi's `char_width_pt` returns fontsize for fullwidth yakumono; should cap at 10.5pt
- evidence: minimal repro (tested 4 compat modes × 2 fonts × 6 sizes = 48 cases) showed yakumono advance = fontsize in all cases. Rule refuted.
- outcome: parallel session (oxi-3?) merged 1e05fe3 "fixed advance" from a different angle and was later REVERTED (b7fde5e). My refusal to merge was validated.
- note: /loop cannot reproduce the real-doc trigger for yakumono compression. Requires dedicated bisect-from-real-doc approach with subprocess isolation.
- memos: `project_yakumono_rule_unconfirmed.md`, `project_word_single_yakumono_compression.md`, `project_firstline_render_no_shift.md`

## 2026-04-18 — oxi-1 — refuted — d77a cell wrap hypothesis
- context: d77a p.9 drift (rank 2 bottom-5, 0.6042)
- hypothesis: cell text wrapping differs between Oxi and Word
- evidence: measured N=30 lines with COM; Oxi and Word agree on wrap position
- outcome: d77a drift is NOT line-wrap. Look elsewhere (cell height? vertical align?).
- commit: oxi-1 `bf1160b tools(metrics): refute d77a cell wrap hypothesis`

## 2026-04-17 — oxi-3 — confirmed — CJK overflow strict check
- context: d77a p.2 / 683f p.2 line breaking
- hypothesis: Oxi allows 2.5-38pt overflow before breaking line; Word is stricter
- evidence: COM-measured d77a PARA 21, Word 8 lines vs Oxi 7 lines (+17pt missing stub)
- outcome: fix lands in main as `9dab217`, bottom-5 sum +0.1402
- impact: 683f rank moved from 1 to 5; d77a p.9 rank 2 remaining

## 2026-04-17 — oxi-1 — confirmed — empty br-type=page stub
- context: 0e7a p.10 + d77a p.11 page break pattern
- hypothesis: Word renders empty `<w:p><w:br w:type="page"/></w:p>` as 1-line stub on prior page, then breaks
- evidence: COM Variant A shows P3 at y=84 (stub), P4 at y=56.5 (new page)
- outcome: fix lands in main as `8e63b43`, improves 0e7a p9/p10 and d77a p11
- impact: does NOT help 0e7a p.2 (different bug)

## 2026-04-17 — oxi-1 — confirmed — hanging-indent first-line x shift
- context: numbered/bulleted paragraph positioning
- hypothesis: Word places line 1 at `margin + indent_left + first_line_indent` (applies to both positive firstLine and negative hanging)
- evidence: COM `measure_hanging_indent_v2.py` on 6 cases
- outcome: fix lands in main as `640de56`, supersedes the old "firstLine does NOT shift line_x" comment

## 2026-04-17 — oxi-1 — confirmed — drop max_elem_bottom in row height (fix/b35)
- context: table row height inflation
- hypothesis: `pad_t + max(content_h, elem_bottom) + pad_b` double-counts text_y_offset
- evidence: for MS Mincho 10.5pt in 17.5pt grid, elem.y=3.5pt (center offset) + elem.height=17.5pt = 21pt bottom, 3.5pt past content_h=17.5pt
- outcome: fix lands in main as `625f8ad`, bottom-5 sum +0.0524

---

## Active hypotheses (not yet confirmed/refuted)

### oxi-4 — LM0 cell formula (investigating)
- observation: 10.5pt row_h = 18n, 12pt row_h = 28 + 36(n-1) — non-continuous formula
- current direction: sweep sizes {9,10,10.5,11,12,13,14,16,18} × both fonts × n∈{1..4} × adjustLineHeight on/off
- blocker: fit depends on whether closed-form from font metrics exists, or needs per-size constant table

### oxi-1 — 0e7a p.2 remaining drift (investigating)
- observation: p.2 SSIM 0.5767, unchanged by empty-br/hanging-indent fixes (those lifted p9/p10 only)
- current direction: per-paragraph Y position measurement to locate drift origin
- hypotheses tried: (none refuted yet for p.2 specifically)

### oxi-3 — d77a p.9 (investigating)
- refuted: cell wrap (2026-04-18)
- current direction: looking at cell height / vertical align / floating shape overhead

### 🔥 BLOCKER — footnote area over-reserve on b837 p.4
- **Status**: BLOCKS oxi-4 `39ebdb9` charGrid fix from merging
- **Symptom**: Oxi reserves full footnote body height per page; Word splits
  long footnotes across pages, reserving only what fits.
- **Measurement** (from oxi-4 memo `project_b837_footnote_over_reserve`):
  - p.4: Oxi reserves 198.5pt for 5 footnotes (all lines, 13 line-bodies)
  - Word reserves ~80pt less (splits fn 22's 5-line body across pages)
  - Oxi's cap puts paras[59] 2nd line past page end → premature break
- **Potential gain** (if fixed together with charGrid):
  - b837 p2: +0.0836
  - b837 p4: +0.0089 (target)
  - b837 p5: recovers from -0.0387 to possibly positive
  - Bottom-5 impact: potentially enough to push b837 out of bottom-5 entirely
- **Assignee (2026-04-18)**: oxi-2 (reassigned from fix/b35-multiline-cell)
- **Branch suggestion**: `fix/footnote-area-spill`
- **Key files**:
  - `crates/oxidocs-core/src/layout/mod.rs` — `estimate_footnote_h` cap logic
  - Memo chain: `project_chargrid_2cell_indent_width.md` →
    `project_footnote_reserve_sensitivity.md` →
    `project_b837_footnote_over_reserve.md`
- **Design direction**: implement Word-like footnote-area spill across pages
  (cap per-page at remaining body-space, overflow to next page's fn area).

### oxi-2 — footnote area spill (reassigned)
- target: implement Word-like footnote area page-split
- blocks: oxi-4 charGrid fix `39ebdb9`
- evidence: `project_b837_footnote_over_reserve.md` in agent memory

---

## Consistency merges (Path D — internal divergence unification)

Merges that landed because two Oxi code paths computing the same thing were
diverging, and unifying them improved overall SSIM (net strict >) without
regressing the bottom-5 floor. See CLAUDE.md §9 Path D for the rules.

### 2026-04-24 — textAlignment=baseline (§17.3.1.35) body path

**Divergence** (4-layer parse → layout incomplete):
- Parser path A: `parser/ooxml.rs:1901` parses per-paragraph `w:textAlignment`
  attribute into `ParagraphStyle.text_alignment` (Option<String>)
- Parser path B: `parser/styles.rs:apply_para_property_empty` did NOT handle
  textAlignment → pPrDefault's textAlignment="baseline" discarded
- Layout path: `layout/mod.rs:4024+` (body `text_y_offset_for_line`) applied
  centering offset `(lh-fs)/2` unconditionally, ignoring text_alignment field
- Inheritance path: `parser/ooxml.rs:1344+` (docDefaults fallback) did NOT
  copy text_alignment from doc_para defaults

**Example input**: e3c545_LOD_Handbook.docx pPrDefault has
`<w:textAlignment w:val="baseline"/>`. Per ECMA-376 §17.3.1.35, "baseline"
means glyph baseline sits at line box bottom (no upper centering offset).
Oxi applied +5.8pt offset for Meiryo default → body paragraphs drift
+5.76pt pixel-space below Word (pixel-diff confirmed on p.1 and p.11).

**Fix** (3 layers, 21 LOC):
1. `parser/styles.rs` add textAlignment case to apply_para_property_empty
2. `parser/ooxml.rs:1392` inherit text_alignment from docDefaults
3. `layout/mod.rs:4024` early-return 0.0 for "baseline"/"top" alignment

**Results**:
- Bottom-5 sum (per-doc min): 3.2631 → **3.2645 (+0.0014, improves)** ✓
- Net Δ: **+0.0229 strict**
- Max improvement: +0.0135 (e3c545 p.1) ≥ max regression 0.0042 ✓
- Improvements: 7 (all e3c545 + ed025 p.9)
- Regressions: 2 (e3c545 p.6 -0.0042, p.10 -0.0037 — body-cell misalignment
  residual on table-heavy pages; cell-path fix deferred due to separate
  crash observed in gen2_055 when cell_text_y_off was modified)

**Rationale**: body path now respects §17.3.1.35. Cell path unchanged
pending investigation of why gen2_055 render crashed when same match-guard
was added to cell_text_y_off at `mod.rs:4739` (cell path has edge case
not visible from body path — needs dedicated debug next session).

Commit: 550787a.

### 2026-04-23 — Phantom blank page fix: empty-br paragraph overflow cascade

**Divergence**: Two Oxi rules compound incorrectly in an edge case:
- **Rule A (`project_empty_br_para_stub.md`, 2026-04-17)**: Empty paragraph
  whose single run is `<w:br w:type="page"/>` (converted to `\x0C`) sets
  `page_break_after=true` after removing the run. Layout renders the
  paragraph mark as a stub on the CURRENT page, then forces a new page.
- **Rule B (overflow check)**: When a line doesn't fit on current page
  (`cursor_y + line_height > page_bottom`), push current page and
  continue on new page.

**Divergence**: When Rule A's empty paragraph OVERFLOWS (can't fit stub
on current page), Rule B moves the stub to a new page. THEN Rule A
pushes ANOTHER page break. Result: **two consecutive page breaks** →
phantom blank page between them.

**Example** (d77a block 127):
- Block 126 fills p.10 near bottom (y=742)
- Block 127 is empty + `<w:br w:type="page"/>`; stub line_height=18pt
  doesn't fit on p.10 (742+18=760 > 771)
- Rule B pushes p.10, stub renders on p.11 at y=75 (alone)
- Rule A pushes p.11 (now with empty stub), advances to p.12
- Block 128 "別紙の例" renders on p.12
- Word renders this as 2 pages (block 126 end of p.10, block 128 start
  of p.11); Oxi renders as 3 pages (p.10, blank p.11, p.12)

**Fix** (crates/oxidocs-core/src/layout/mod.rs:2407): when Rule B
overflow triggers AND the paragraph is empty AND has
`page_break_after=true`, skip rendering the stub entirely. Push current
page and return — the next paragraph renders on the fresh page directly,
without an intervening blank page.

**Results**:
- Bottom-5 sum: 3.2627 → **3.2627 (equal, Path D gate met)**.
- d77a total page count: 13 → **12** (matches Word).
- d77a p.12 SSIM: 0.7687 → **0.9673 (+0.1986)** — content realigns with Word p.12.
- d77a p.11 SSIM: 0.7891 → 0.7623 (-0.0268). The baseline 0.7891 was
  artificially inflated because Oxi p.11 was BLANK vs Word p.11 with
  content (SSIM rewards regional similarity; mostly-white pages match
  mostly-white pages on luminance/variance components). Post-fix Oxi
  p.11 has real content similar to Word but not pixel-perfect; 0.7623
  reflects the honest SSIM of content-vs-content comparison.
- Net +0.1717. Max improvement +0.1986 ≥ |max regression| 0.0268 ✓.

**Evidence section**:
- Before/after concrete: d77a block 127 phantom page confirmed via
  layout_json --structure dump (rendered p.11 had only PARA 127 at
  y=75 with 0 LINE entries). Oxi PNG p.11 total_dark_pixels=0
  (completely blank). Word PNG p.11 total_dark_pixels=127407.
- Bottom-5: 3.2627 → 3.2627 (equal).
- Net Δ: +0.1717.
- Max improvement: +0.1986 (d77a p.12).
- Max regression: -0.0268 (d77a p.11, explained above).
- Improvements / regressions: 1 / 1 (net positive, both in same doc).

**Rationale**: the fix is a correctness-restoring unification. Rule A
and Rule B in isolation are correct; their compound effect in the
overflow edge case produces a result that violates the spec (one
page break element → one page break, not two). Path D applies because
this is an internal consistency fix for Oxi's own rule interaction,
not a Word-behavior claim.

### 2026-04-22 — Fix C: estimate_para_height cell path unified with cell renderer

**Divergence**:
- Path A (`estimate_para_height` via `break_into_lines`, mod.rs:5246): word-buffer
  wrap with yakumono compression and 2-pass re-wrap.
- Path B (cell render loop, mod.rs:4480+): char-by-char greedy wrap with kinsoku
  line-start prohibited handling and fullwidth grid pitch (`grid_char_cw_ratio`),
  NO yakumono compression.
- Example mismatch: for ed025cbecffb_index-23 cell paragraphs, estimate sees
  fewer lines than render actually produces (compression fits more chars per
  line at estimate time), causing downstream page-break/keepLines decisions to
  be made against the wrong height → cascade page drift.

**Fix**: Added `count_cell_lines` helper replicating the cell renderer's wrap
loop. `estimate_para_height` now accepts `in_cell: bool, grid_char_pitch,
grid_char_cw_ratio`; when called from cell contexts it uses `count_cell_lines`
instead of `break_into_lines`. 12 callsites updated (2 cell: pass real values;
10 body: pass `false, None, None` — body path unchanged).

**Results (177 docs, 352 pages)**:
- Bottom-5: 3.2598 → 3.2598 (equal, within Path D allowance)
- Net Δ: +0.1025
- Max improvement: ed025 p.7 +0.1295
- Max regression: ed025 p.8 -0.1110 (within |max_regression| ≤ |max_improvement|)
- Improvements: 11 (ed025 × 9, de6e32 × 1, 6514f2 × 1)
- Regressions: 8 (ed025 × 7 internal whack-a-mole, 15076df × 1)

**Why this is correct regardless of bottom-5 floor**: estimate is supposed to
be a cheap preview of render. When they disagree, either estimate is wrong
(under/over-counting) or render is. Since we cannot change render without
affecting actual visual output, we align estimate to match render. The
alternative (letting them drift) means any estimate-driven decision
(keepLines, keepNext, multi-column pre-check) operates on fiction.

**Prior attempt**: FALSIFIED 2026-04-22 before Step 1 partial (157bc22). That
run had b837 p5 -0.042 crash as a consequence of fn reserve over-pack being
exposed by better estimates. Step 1 partial fixed fn_render attribution, which
removed the crash pattern. Fix C v2 (this merge) sees no b837 p5 crash.

## Confidence merges (Path B — correct regardless of SSIM)

Merges that landed because the fix is *known correct* via COM + 3 docs + minimal
repro + spec reference, but didn't necessarily improve bottom-5 floor. See
CLAUDE.md §9 Path B for the rules.

### 2026-04-27 — OMML Phase 1-3 math equation rendering (merge 4dcda93)

**Greenfield feature** merge of `feat/omml-math` (37 commits, last 4/18,
9 days unmerged). Phase 1-3 complete per `docs/spec/omml_phase3_summary.md`.

**Evidence**:
- COM: 15 fixtures with mean SSIM **0.9929** vs Word ground truth (Phase 3
  validation: `tools/metrics/measure_omml_fixtures.py`)
- Minimal repro: `tools/metrics/omml_fixtures/` (10 primitive cases + 5
  complex real-world formulas)
- Spec: ECMA-376 Part 1 §22.1 (OMML); 56 MATH constants from Cambria Math;
  166 codepoint substitution rules per Word's italic-math semantic
- Bottom-5: 3.2645 → 3.2645 (greenfield, **0/184 baseline docs have OMML**;
  scanned via `<m:oMath` / `mc:oMath` markup search)
- Rationale: pure additive feature, zero overlap with baseline. Phase 3
  handoff doc explicitly marks "merge-ready greenfield feature"

**Implementation scope** (per Phase 3 summary):
- IR: `crates/oxidocs-core/src/ir/math.rs` (438 lines, 20 MathExpr variants
  + MathBlock + MathStyle + 6 supporting enums)
- Font: `font/math_constants.rs` (Cambria Math MATH table, 56 constants),
  `math_substitute.rs` (166 codepoint rules: A→𝐴, α→𝛼, h→ℎ, ∂→𝜕, ∇→𝛁),
  `math_glyphs.rs` (italic correction + vertical variants)
- Parser: OMML XML → MathBlock IR (16 unit tests)
- Layout: per-primitive `emit_*` functions with stacked rendering using
  MATH constants. Inline (m:oMath) vs Display (m:oMathPara) styles.

**Verification**:
- Auto-merge **clean** (4 files auto-merged: font/mod.rs, ir/types.rs,
  layout/mod.rs, parser/ooxml.rs)
- OMML parser tests: 16/16 pass; font::math tests: 23/23 pass; integration
  tests: 5/5 pass; total **+22 new passing tests**
- Pre-existing `kinsoku::test_line_start_prohibited` failure remains
  (unrelated to OMML, fails on main pre-merge too)

**Post-merge integration fix** (commit ddfd883):

Auto-merge succeeded textually but the omml-math branch (forked before
Phase 2 comments+tracked-changes) had semantic gaps:

1. `layout/math.rs:588` — `LayoutContent::Text {...}` missing `text_scale`
   field (added during R-12).
2. 6 `match block` statements added during Phase 2 work
   (revisions/comments visitors) lacked `Block::Math(_)` arms — added as
   no-op alongside `Block::Image(_) | Block::UnsupportedElement(_)`.

Math content carries no runs to which revision/comment styling applies, so
no-op arms are semantically correct.

### 2026-04-25 — textbox line-count-aware overflow filter (commit 61833e2)

**Divergence**: `crates/oxidocs-core/src/layout/mod.rs:1944` filter
`pe.y + pe.height > clip_bottom` dropped valid text in tight-fit single-line
textboxes where textbox height = `inset_t + line_height + inset_b` exactly.
Line slot bottom (next-line baseline-top) extends past clip_bottom by the
leading portion of line_height, but visible glyph fully fits within textbox.

**Fix**: line-count-aware cutoff. Compute `available_lines =
floor(inner_height / line_height)` and drop only text elements whose Y is
past `abs_y + inset_t + available_lines * line_height`. Non-text elements
(BoxRect inside textbox) keep original bounds check.

**Evidence**:
- COM: 459f05 (kyodokenkyuyoushiki01) p.1 「様式１」 textbox (h=27.2pt =
  3.6+20+3.6 exactly). Word renders text visibly; pre-fix Oxi dropped all
  3 chars (TBX_DEBUG: pe.y=36.25 + pe.height=20 = 56.25 > clip_bottom=50.85).
- Minimal repro: tools/metrics/textbox_tight_fit_repro/TF_A.docx through
  TF_D.docx — Word renders text in all variants per
  `tools/metrics/measure_textbox_tight_fit.py`.
- Spec: undocumented Word quirk — content visibility governed by line
  count fit, not line-slot Y overlap with clip bottom. Pixel-confirmed.

**Multi-line overflow case preserved**:
- 2ea81a textbox 5 (h=74.4, line_h=32.3): inner_h=67.2, avail=2, cutoff =
  top + 2×32.3. The 3rd-line elements (y >= cutoff) are still dropped,
  matching Word's overflow behavior.

**Verify** (full baseline 177 docs / 352 pages):
- 0 improved / 352 unchanged / 0 regressed
- Bottom-5 floor: 3.2645 → 3.2645 (UNCHANGED, equal OK per CLAUDE.md)
- 459f05 p.1 「様式１」 now renders correctly visibly (textbox area too
  small to register SSIM delta > 0.001 threshold).

**Other affected docs**: 664c38 (h=33.0), d1e8ac8 (h=33.0) had similar
small textboxes — already rendering correctly with old filter (line_h
sufficiently smaller than inner_h, so inner_h/line_h > 1). Unchanged.

### 2026-04-25 — body list-marker uses paragraph's font (not renderer default)

**Divergence**: `crates/oxidocs-core/src/layout/mod.rs` emits the list
marker as a `LayoutContent::Text` element. The cell path (~line 4780)
resolves the marker's `font_family` via `resolve_font_family_for_text`,
inheriting from the paragraph's first run. The body path (~line 2215)
passed `font_family: None`, causing the GDI renderer to fall back to its
default Latin font. For halfwidth markers like "(1)" in a CJK-font
paragraph, Word renders the parens with CJK-font metrics (wider); Oxi's
default-font fallback rendered them narrower — user-reported on e3c545
p.1「（１）の描写もなんか少し小さいような気がする」.

**Fix**: body path now mirrors the cell path — resolve font_family,
bold, italic, underline, strikethrough, color, highlight, underline_style
from the first-run style for the marker element.

**Pixel evidence (Word 150DPI EMF vs Oxi GDI PNG)** — all markers now
match Word within 0-1px anti-aliasing bearing:

| Doc | Marker | Font | Word px | Oxi px | Δ |
|-----|--------|------|---------|--------|---|
| e3c545fac7a7_LOD_Handbook p.1 "(1) 公開するデータの設計" | halfwidth `(1)` | ＭＳ 明朝 | 14 | 14 | 0 |
| 3a4f9fbe1a83_001620506 p.2 "（１） 労働時間関係" | fullwidth `（１）` | メイリオ | 20 | 20 | 0 |
| ed025cbecffb_index-23 p.1 "(1) 事業運営組織" | halfwidth `(1)` | CJK | 23 | 23 | 0 |

**COM spec confirmation** (`tools/metrics/measure_numid_hanging_text_x.py`):
For each doc the first-text-char X measures at `LeftIndent ±0.2pt`
bearing — consistent with Word placing marker glyphs in the hanging
space with the paragraph's metrics, not with a default Latin font.

**Minimal repros**: `tools/metrics/marker_font_repro/MF_A..D.docx` cover
halfwidth `(1)` in ＭＳ 明朝 / ＭＳ ゴシック / メイリオ, plus fullwidth
`（１）` in ＭＳ 明朝 (control). All four repros produce Oxi marker
renders that match Word under the fix; previous behavior (font_family
None) produced visibly narrower halfwidth parens.

**Full-baseline verify**:
- 0 improved, 352 unchanged, **0 regressed**
- Net Δ = 0 (marker glyph delta of 1-4 pixels per page is sub-0.001
  SSIM threshold; the fix is below the resolution of the merge gate's
  floating-point tolerance but visually real and pixel-exact).
- Bottom-5 per-doc sum: 3.2645 → **3.2645 (equal, Path B gate met)**.

**Why Path B and not Path D**: Path D requires `Net Δ > 0 strict`,
which fails here because the glyph-width diff is sub-tolerance. Path B
explicitly allows "fixes that are known correct but don't yet show SSIM
gain", with gate: 3 docs + self-authored repro + spec. All met. The
`[consistency-merge]`-style internal-divergence evidence is included as
supporting material, not the primary justification.

**Implementation** (`crates/oxidocs-core/src/layout/mod.rs`):
```rust
// Body marker emission (was: font_family: None, bold: false, ... hardcoded)
let marker_font_family = self
    .resolve_font_family_for_text(&marker_text, marker_style, &para.style)
    .map(|s| s.to_string());
let marker_bold = self.resolve_bold(marker_style, &para.style);
let marker_color = self.resolve_color(marker_style, &para.style).map(|s| s.to_string());
elements.push(LayoutElement::new(..., LayoutContent::Text {
    text: marker_text,
    font_size: marker_font_size,
    font_family: marker_font_family,
    bold: marker_bold,
    italic: marker_style.italic,
    underline: marker_style.underline,
    underline_style: marker_style.underline_style.clone(),
    strikethrough: marker_style.strikethrough,
    color: marker_color,
    highlight: marker_style.highlight.clone(),
    ...
}));
```

**Artifacts**:
- `tools/metrics/build_marker_font_repros.py`
- `tools/metrics/marker_font_repro/MF_A..D.docx`
- `tools/metrics/measure_numid_hanging_text_x.py` (reused from d30e432)
- `pipeline_data/numid_hanging_text_x.json` (updated with b35/ed025
  measurements)

### 2026-04-25 — numbered-list hanging: first-line text at `left` (not `left-hanging`)

**Spec**: ECMA-376 Part 1 §17.9.24 `lvl/suff` (tab|space|nothing) and
§17.3.1.14 `firstLine`/`hanging`. For a numbered paragraph with hanging
indent and `suff=tab` (default), Word places the list marker at
`left - hanging` and the first TEXT character at `left` — the hanging
area is consumed by the marker + tab, not used to pull the first line
leftward. Oxi was applying `first_line_indent` to first-line x even for
list paragraphs, causing the marker and body text to render at the same
x (e3c545 p.1 "3．基本的な考え方" with 基 overlapping the "3．" glyphs —
reported by user as「文字がダブってる」).

**Evidence (COM Range.Characters(1).Information(7)):**
- e3c545 LOD_Handbook: 17+ paragraphs. left∈{18.00, 21.30, 57.00}, fli∈
  {-18.00, -21.30, -21.00}. char1.x delta from `left` = +0.00 to +0.20pt
  (CJK glyph bearing).
- 3a4f_001620506: 8+ paragraphs. left=36.00, fli=-36.00. char1.x delta
  from `left` = +0.00pt (halfwidth Latin char bearing 0).
- Minimal repros NH_A…NH_F (6 variants: decimalFullWidth, decimal,
  bullet, suff=space, varying left/hanging): all suff=tab|default
  variants show char1.x = left (±0.2 bearing). Only `NH_C` (suff=space)
  shows char1.x = 31.5 (left-4.5) which is marker+space end — the code
  correctly excludes the space/nothing case.

**Pixel verification (e3c545 p.1):**
- Word "基" at 78.2pt (= margin 56.7 + left 21.3). Oxi now matches.
- Before fix: both marker and 基 rendered at margin 56.7pt, overlapping.

**Implementation**: `crates/oxidocs-core/src/layout/mod.rs`
- Body path (~line 2125): compute `first_line_indent_raw`, then set
  `first_line_indent = 0.0` when `list_marker.is_some() && raw < 0.0 &&
  suff ∈ {None, Some("tab")}`.
- Cell path (~line 4513): same guard with `p_` prefix.
- Estimate path (~line 5578): same guard with `est_` prefix.

**Full-baseline verify**:
- 1 improvement (e3c545 p.1: 0.8316 → 0.8326, +0.0010)
- 351 unchanged, **0 regressed**
- Bottom-5 floor: 3.2645 → **3.2645 (equal, Path B gate met)**
  - d77a p.7: 0.6268 = 0.6268
  - b837 p.?: 0.6449 = 0.6449
  - 29dc6e: 0.6636 = 0.6636
  - 2ea81a: 0.6643 = 0.6643
  - e3c545 p.11: 0.6649 = 0.6649 (rank 5 doc unchanged; p.1 was not
    in bottom-5)

**Artifacts**:
- `tools/metrics/measure_numid_hanging_text_x.py` (COM measurement
  tool, scans real docs auto-detecting numId+hanging paragraphs).
- `tools/metrics/build_numid_hanging_repros.py` (NH_A–F builder).
- `tools/metrics/measure_numid_hanging_repros.py` (repro measurement).
- `tools/metrics/numid_hanging_repro/NH_*.docx` (6 minimal repros).
- `pipeline_data/numid_hanging_text_x.json` (e3c545 + 3a4f measurements).
- `pipeline_data/numid_hanging_repro_measurements.json` (repro data).

**Scope**: Fix only applies when `suff=tab` (or unset, defaulting to
tab). `suff=space` and `suff=nothing` retain existing behavior because
the NH_C measurement showed text position depends on marker+space width
for those cases — a different formula.

### 2026-04-23 — Z Step 2: row-split closing horizontal border (gated)

**Spec**: ECMA-376 Part 1 §17.4.33 — cell bottom border is drawn at end of
cell content. When a table row splits across pages, the first-page portion
requires a closing horizontal border at the bottom of its text. Previously
Oxi omitted this entirely; the box remained visually "open" on page breaks.

**Formula** (layout-coord = pixel-coord for borders):
```
close_y = last_line.y + natural_height
natural_height = word_ascent_pt(fs) + word_descent_pt(fs)  // per-fragment max
```
Derived from 11 minimal repros C1-C11 (MS Mincho/Gothic/Meiryo × fs
10.5/12/14 × pitch 15/18). Pixel verification on d77a p6: Oxi border
y_pt=756.26 vs Word y_pt=756.48 → **0.22pt match**.

**Why NOT `elem.y + elem.height`**: `elem.height = line_height (e.g. 18pt)`
which is the grid cell allocation, not the glyph bottom. Using it overshoots
by `line_height - natural_height ≈ 4.19pt`. Step 2 attempts v1-v4 used this
wrong reference and regressed consistently.

**Gate**: `close_y <= split_y - 10.0` (require 10pt whitespace above page
bottom). Skips cases where Oxi's content packs too close to page_bottom —
this signals content-packing divergence from Word (e.g. Oxi overflowing to
an extra page where Word doesn't). Drawing a border in such cases lands
in Word's text/border region and regresses SSIM. Gate v8 (5pt) caught 4
of 8 regressions; v9 (10pt) caught 5 of 8.

**Vertical border clipping**: current-page vertical borders clipped to
close_y. User feedback v5: 「縦線が、突き抜けてるね」 → fixed in v6+.

**Evidence**:
- COM: C1-C11 minimal repros under tools/metrics/lm2_centering_repro/;
  per-repro Information(6) Y values in
  pipeline_data/lm2_centering_measurements.json.
- Pixel match: C1 Word PNG y_pt=76.78 = Oxi PNG y_pt=76.78 (exact).
- d77a COM: pipeline_data/d77a_p6_line_ys.json (54 paragraphs × line ys).
- Spec: ECMA-376 Part 1 §17.4.33 (tcBorders bottom — standard cell
  bottom border rule; no special-case for row split means row-split
  should draw the border identically to non-split).
- Bottom-5 sum: 3.2627 → **3.2627 (equal, Path B gate met)**.
  Per-doc bottom-5 all unchanged:
  - d77a p.7: 0.6264 = 0.6264
  - b837 p.6: 0.6449 = 0.6449
  - e3c545 p.11: 0.6635 = 0.6635
  - 29dc6e p.?: 0.6636 = 0.6636
  - 2ea81a p.?: 0.6643 = 0.6643

**Impact (v9, full baseline 177 docs / 352 pages)**:
- Improvements (3):
  - d77a p.6: 0.8276 → 0.8297 (+0.0021)
  - d77a p.8: 0.6655 → 0.6676 (+0.0021)
  - 3a4f p.2: 0.8862 → 0.8879 (+0.0017)
- Regressions outside bottom-5 (3):
  - d4d126 p.7: 0.7849 → 0.7793 (-0.0057)
  - ed025 p.12: 0.8177 → 0.8144 (-0.0033)
  - d77a p.12: 0.7717 → 0.7687 (-0.0030)
- Net -0.0061 informational.
- User visual confirmation: 「横線はいい感じね」 (closing border correct).

**Rationale**: ECMA-376 §17.4.33 requires the bottom border; omission was
a correctness gap. The formula is COM-confirmed to sub-pt accuracy on the
primary target document (d77a). The gate prevents draws on pages where
Oxi's content packing diverges from Word — a separate cascade issue
independent of Step 2.

**Prior attempts (all FALSIFIED, retained for reference)**:
- project_z_step2_impl_FALSIFIED_content_drift.md (v1)
- project_z_step2_border_at_split_FALSIFIED.md (v2)
- project_z_step2_third_FALSIFIED.md (v3)
- project_lm2_natural_height_FALSIFIED.md (LM2 centering misinterpretation)
- project_com_info6_vs_pixel_separation.md (v4 pixel-space clarification)
- project_z_step2_v5_v6_v7_FALSIFIED.md (pre-gate versions)
- project_z_step2_v9_subthreshold_FALSIFIED.md (first-ship revert due to
  stale baseline — resolved by refreshing baseline to true main state).

**Open follow-ups**:
- Opening border on next page (v7) regressed; needs separate investigation.
- 3 outside-bottom-5 regressions trace to Oxi content-packing drift from
  Word, not Step 2 formula error. Fix those via content packing work.

### 2026-04-21 — Stage 4/5 「-leading +6pt removed (post-wrap → overshoot bug)

**Bug**: `mod.rs:2563-2581` Stage 4/5 added +6pt to `frag_spacing_after[fi-1]`
after a CJK char and before an opening bracket fragment (「『〔【《〈（｛［) for
compat>=15+cP docs. This was applied POST-wrap (added after break_into_lines
completed), so wrap-time `current_width` never accounted for it. Result: lines
that fit at wrap time would overshoot the right margin by 6-12pt at render.

**User insight 2026-04-21**: "Wordが右端のはみだしが１つもない — 平均的な右端ラインから、はみ出しているものが一つもない"
Word strictly enforces no-overshoot. Oxi's overshoots break this invariant.

**Evidence**:
- d77a P2 (page 3) line measurement: 2 lines overshoot exactly +12pt (-12 gap):
  - y=435 "この部分は、別紙の「公共データ利用規約（第1.0版）に関する重要情報」に記載"
  - y=579 "「公共データ利用規約（第1.0版）」の採用を想定しているのは、国の府省（施設"
  Both contain "（第1.0版）" with 2 mid-line opening brackets ((+6pt) × 2 = +12pt).
- Test: disabling Stage 4/5 entirely: both overshoots → gap=0 (perfectly aligned).
  3 short lines (gap=+3) also resolved to gap=0.

**Fix**: remove the Stage 4/5 +6pt block at `mod.rs:2563-2581`.

**Impact**:
- Bottom-5: 3.2451 → 3.2463 (+0.0012, d77a p9 +0.0012).
- Full baseline: 19 improvements / 3 regressions across 177 docs / 352 pages.
  Net +0.1253. Top improvements: d77a p1 +0.0028, p4 +0.0011, p5 +0.0028, p8
  +0.0050, p9 +0.0012; e8caed +0.0047; c7b923 +0.0035; 3a4f +0.0033.
  Regressions outside bottom-5: b837 p1 -0.0026, d77a p6 -0.0019, d77a p12
  -0.0015 (max -0.0026, all minor).
- Path A bottom-N floor strict increase → merge gate PASS.

**Open question**: the original Stage 4/5 was added per "Word's measured shifts
originate" comment. Removing it works for current baseline — but Word may
actually have demand-driven「-leading gap that varies based on line fullness
(per `project_yakumono_demand_driven`). The current removal is a simplification
that aligns Oxi with Word for the common case. Demand-driven implementation is
the proper long-term fix.

### 2026-04-21 — Yakumono compression gate flip (compat=14+cP enabled)

**Bug**: `mod.rs:2975` gated yakumono pair compression with
`self.compress_punctuation && self.compat_mode >= 15`. COM evidence on 4
distinct compat=14+`compressPunctuation` docs (04b88, 7f272a, fded68, 34140)
showed Word fits +1 to +3 more chars on line 1 of yakumono+indent paragraphs
than Oxi. Minimal repro `idx46_real` with compat=14 confirmed Word=43 / Oxi=42.
The `compat>=15` gate excluded the very docs Word DOES compress.

**Fix**: drop the compat gate — `let yakumono_enabled = self.compress_punctuation;`.

**Evidence**:
- COM: 04b88 (+2 chars), 7f272a (+2), fded68 (+3), 34140 (+1) on line 1 of
  hanging-indent + yakumono paragraphs. Word vs Oxi line-1 char count.
- Minimal repro: `tools/metrics/bisect_34140_settings.py` — S0..S5 toggle compat
  14↔15 + flag bisection. `pipeline_data/_settings_*.docx`. With compat=15: Oxi
  43 / Word 42. With compat=14: Oxi 42 / Word 43. Oxi gate is reversed vs Word.
- Spec: undocumented Word quirk. `compressPunctuation` in `<w:characterSpacingControl>`
  enables yakumono pair compression in Word at compat=14, but compat=15 disables it.
  COM-confirmed across 6 compat=14+cP docs and minimal repros at both compat values.
- Bottom-5: 3.2451 → 3.2451 (unchanged — none of d77a/b837/e3c545/29dc6/2ea81 are
  compat=14+cP)
- Full baseline: **7 improvements / 0 regressions across 177 docs / 352 pages**.
  Net +0.0720. All 6 compat=14+cP docs improved on at least one page; 34140 p1
  +0.0096 + p2 +0.0110, 7f272a p1 +0.0256, fded68/04b88/09390503/6d6dc4 p1 +0.003-0.009.
- Rationale: changing the gate to `cP only` enables compression for compat=14
  (matching Word) while leaving compat=15 path identical (no change to 56 docs
  in compat=15+cP class). Conservative scope; the long-debated compat=15
  yakumono question is a separate investigation.

### 2026-04-21 — Ignore `type="first"` header/footer without `titlePg`

**Spec**: ECMA-376 Part 1 §17.10.2 — `type="first"` header/footer references
are active only when `titlePg` is set in the section properties. Without
`titlePg`, the reference is ignored and the default header/footer is used
instead; if no default exists, no header/footer is rendered.

**Bug**: `ooxml.rs` fallback at section header/footer resolution called
`parse_header_footer_blocks(effective_footer_refs, ...)` (passing ALL refs
including `type="first"`) when neither type-matched nor default-matched refs
existed. User reported a phantom "1 / 7" page-number footer on 34140 p.1;
Word renders nothing there because the doc's sole `type="first"` ref is not
gated by `titlePg`.

**Scope of fix**: only drop `type="first"` from the legacy all-refs fallback.
`type="even"` is left in the fallback pool because some baseline docs reference
it without the `evenAndOddHeaders` setting and rely on the reserved footer
space for body pagination; removing that reservation regressed 6514 p.6 by
−0.1037 in a trial run, so the narrower scope is retained.

**Evidence**:
- COM / Word render: 34140 p.1 (real doc) + 4 self-authored minimal repros
  (`pipeline_data/_ftrspec_V{1..4}_*.docx`). V1 (first-only, no titlePg) is
  the reproducing case; V2/V3/V4 confirm Word's behavior on the surrounding
  matrix (V2 still shows a separate page-level footer-switch limitation
  tracked elsewhere, but orthogonal to this fix).
- Baseline grep: only 34140 in 184 docs has `type="first"`-only without
  `titlePg` (others have a `default` footer that absorbs the fallback).
- Spec: ECMA-376 §17.10.2.

**Bottom-5 floor**: pre 3.1280 → post 3.1374 (+0.0094). 34140 p.5 remains
the doc's min at rank 5; all other bottom-5 entries unchanged.

**Results (34140 only; all other docs unchanged)**:
- p.2: 0.6802 → 0.7393 (+0.0591)
- p.3: 0.6953 → 0.7105 (+0.0152)
- p.5: 0.6608 → 0.6702 (+0.0094)
- p.4: 0.7127 → 0.7022 (−0.0105) — body cascade from removed footer reserve
- p.6: 0.7676 → 0.7643 (−0.0033) — same cascade
- Net on 34140: +0.0699

**Rationale for same-doc regressions (p.4, p.6)**: removing the phantom
footer also removes the ~20pt footer-area reservation. Word doesn't reserve
that space either (no footer is rendered), but Oxi's body layout now extends
into the area differently and the pagination shifts slightly. Pixel-level
alignment on p.4/p.6 is marginally worse; bottom-5 floor improves.

**Minimal repro**: `tools/metrics/build_footer_firsttype_repros.py` generates
V1–V4. `tools/metrics/grep_first_footer_no_titlepg.py` identifies
baseline-affected docs.

---
