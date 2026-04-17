# oxi-4 worktree — fix/lm0-cell-formula

## Focus
Derive Word's closed-form **LM0 multi-line table cell height** formula by COM
measurement sweep. This gates 683f p2 (rank-1 bottom-5, 0.4908) and b35 p1
(rank-4, 0.6134) fixes.

## Background
- `pipeline_data/ssim_baseline.json` bottom-5 (2026-04-17 post-d77a-p9 merge):
  1. 683f p2 = 0.4908  ← primary unclaimed
  2. 0e7a p2 = 0.5599  (oxi-1 fix/0e7a-p2)
  3. d77a p9 = 0.6032  (just merged)
  4. b35  p1 = 0.6134
  5. b837 p4 = 0.6321
- Existing data (`oxi-main/tools/metrics/output/lm0_multiline_cell.json`):
  - MS Mincho/Gothic 10.5pt: row_h = 18×n   (line_gap=18, last_alloc=18)
  - MS Mincho/Gothic 12.0pt: row_h = 28+36×(n−1)  (line_gap=36, last_alloc=28)
- Oxi currently allocates `n × line_gap` → overshoots 12pt cells by 8pt each,
  accumulates as bottom-N drift.

## Goal
Produce a closed-form `(line_gap, last_line_alloc)(font, size)` across the
common Word sizes, then implement the fix in `layout/mod.rs`.

## Workflow
1. **Expand sweep** to sizes {9, 10, 10.5, 11, 12, 13, 14, 16, 18} × both fonts
   × n ∈ {1..4} × `<w:adjustLineHeightInTable/>` on/off.
2. Save to `tools/metrics/output/lm0_multiline_cell_v2.json`.
3. Fit `line_gap(size)` and `last_alloc(size)` — should depend only on font
   metrics (ascent/descent) and Word's snap-to-pitch behavior.
4. Cross-check hypothesis against 683f p2 and b35 p1 raw data.
5. Implement in Rust. Rebuild WASM + clear `oxi_png/`.
6. **Quick verify (target docs only)** — render just 683f + b35 with
   `oxi-gdi-renderer`, compute SSIM vs Word for the target pages
   (683f p.2, b35 p.1). Full pipeline.verify takes ~6min on all 177
   docs; quick verify is <30s and gives immediate signal.
   - If either target page did not improve → formula is wrong.
     Revise hypothesis (step 3) and re-implement. Do NOT proceed to
     full verify.
   - If both target pages improved → proceed to step 7.
7. `pipeline.verify` → bottom-5 floor sum gate (phase 2, N=5, pre_sum=2.8994).
   Only run full verify once quick-verify passes, so we don't burn 6min
   per failed formula iteration.
8. Commit only if bottom-5 floor sum strictly increases.

## Merge gate (FINAL 2026-04-16, Phase 2, N=5)
```python
def floor_sum(baseline, n):
    mins = sorted(min(p.values()) for p in baseline.values())
    return sum(mins[:n])

pre  = floor_sum(pre_fix_baseline, 5)   # = 2.8994
post = floor_sum(post_fix_baseline, 5)
merge_ok = post > pre
```

## Worktree coord (2026-04-17)
| worktree | branch | target |
|----------|--------|--------|
| oxi-main | main   | (integration only) |
| oxi-1    | fix/0e7a-p2 | 0e7a p2 (rank 2, LM0 empty-para class) |
| oxi-2    | fix/1ec1-v2 | (memory: obsolete) |
| oxi-3    | fix/2ea81-p2 | 2ea81 p2 (rank 6, LM2 linePitch) |
| **oxi-4**| **fix/lm0-cell-formula** | **LM0 cell formula (gates 683f+b35)** |

## No-deferred rule (2026-04-17)
One loop iteration must complete **measurement → implementation → quick verify**
as a single unit. Do NOT stop after measurement with "next steps deferred"
unless explicitly blocked. If out of time within an iteration, commit
work-in-progress to a WIP branch so the next iteration can resume without
re-measuring.

## Loop prompt
`/loop 進めて` — read this file and CLAUDE.md, continue the Ra loop.
