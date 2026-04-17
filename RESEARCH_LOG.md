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

---

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

(none yet — first one will land here)

---
