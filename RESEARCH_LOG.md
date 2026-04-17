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

## 🔥 BLOCKER: GDI preset render coverage (Path A fix target)

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

## Confidence merges (Path B — correct regardless of SSIM)

Merges that landed because the fix is *known correct* via COM + 3 docs + minimal
repro + spec reference, but didn't necessarily improve bottom-5 floor. See
CLAUDE.md §9 Path B for the rules.

(none yet — first one will land here)

---
