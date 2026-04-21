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

## Confidence merges (Path B — correct regardless of SSIM)

Merges that landed because the fix is *known correct* via COM + 3 docs + minimal
repro + spec reference, but didn't necessarily improve bottom-5 floor. See
CLAUDE.md §9 Path B for the rules.

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
