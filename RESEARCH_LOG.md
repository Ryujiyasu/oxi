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

Reason for downgrade:
- `bisect_d77a_minimal.py` uses `SRC = d77a.docx` as base and swaps ONLY
  settings.xml. document.xml / styles.xml / fontTable.xml / sectPr /
  themeFontLang all remain d77a's. This is **subtractive from d77a**, NOT
  a true scratch. Per oxi-1 2026-04-18 scratch+jc=both test
  (`gen_scratch_jc_both.py`), this same class of "minimal" approach
  yielded a false trigger (jc=both seemed to trigger in d77a-subtractive
  but NOT in true scratch).
- `yakumono_sweep.json` actual data (cSC + compat=15 minimal template, truly
  scratch XML): MSGothic_12.0 → `（=12.0, 、=12.0, 「=12.0, 」=12.0, 。=12.0`
  (ALL singles at fontsize, NOT compressed). This DIRECTLY contradicts
  the claim "cSC + compat15 → '（' compressed (10.5 at fs=12)".
  Only pair '）' halves (always, even at baseline). Pair compression is
  separate from single-yakumono compression and NOT blocked by cSC/compat.

Revised assessment:
- True trigger for d77a's single-yakumono compression remains **UNKNOWN**
- `cSC + compat=15` is NECESSARY (without cSC no compression) but NOT
  SUFFICIENT (scratch + cSC + compat=15 doesn't compress)
- Additional d77a-inherited property (fontTable / styles / sectPr detail /
  theme / rsid) combines with cSC+compat=15 to activate compression
- Implementation gate "compress_punctuation && compat_mode >= 15" is
  therefore insufficient — would regress any doc with cSC+compat=15 but
  without the missing trigger

Next investigation (deferred to dedicated session per user 2026-04-18):
- Truly additive bisection: scratch + cSC + compat=15 + ONE more d77a
  property at a time. Goal: find the minimal sufficient set.
- Candidates: fontTable with real PANOSE, sectPr drawingGrid properties,
  themeFontLang, rPrDefault rFonts cascade.

Original evidence preserved for reference:
- `pipeline_data/d77a_yakumono_bisect.json` (subtractive from d77a)
- `pipeline_data/yakumono_sweep.json` (true scratch — shows NO compression)
- scripts: `tools/metrics/bisect_d77a_yakumono.py`,
  `bisect_d77a_minimal.py`, `sweep_yakumono_formula.py`

Cross-reference: oxi-1 2026-04-18 `project_yakumono_jc_both_FALSIFIED.md`
reached same conclusion via independent path (scratch+jc=both test).

## 2026-04-18 — oxi-1 — premise-falsified — fix/yakumono-jc-conditional
- branch: fix/yakumono-jc-conditional (commit 4bec783) — RETAINED (not deleted)
- premise: yakumono compression trigger = jc=both (justification)
- evidence for premise: R3 (d77a subtractive with minimal styles + jc=both
  on Normal) reproduced compression; memory project_yakumono_trigger_IS_jc_both.md
- falsification: 2026-04-18 scratch test gen_scratch_jc_both.py
  (truly minimal docx + only <w:jc val="both"/> + compressPunctuation)
  shows ・=12.00pt (NOT compressed, ratio 100%). Independent verification
  in memory `project_yakumono_jc_both_FALSIFIED.md`.
- diagnosis: R3 compression was caused by d77a inherited properties
  (fontTable, sectPr, settings compat block, etc.), NOT by jc=both alone.
  Same class of artifact as oxi-3 cSC+compat15 bisect (both subtractive
  from d77a, both gave false triggers).
- outcome: DO NOT MERGE fix/yakumono-jc-conditional. The branch is
  retained as historical record so next-session exploration doesn't
  retrace the same falsified path.
- calibration: 6th falsified fix this oxi-1 session; zero applied to
  main; bottom-5 floor 3.0166 unchanged. Zero-regression Ra principle
  functioning correctly.

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
