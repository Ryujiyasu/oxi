#!/usr/bin/env bash
# Tidy the working tree: delete regenerable scratch output, report the big
# local-only directories, and never touch measurement ground truth.
#
#   scripts/tidy-workdir.sh            # dry run: list what would go
#   scripts/tidy-workdir.sh --apply    # delete the root scratch files
#   scripts/tidy-workdir.sh --apply --prune-tmp   # also empty tmp/
#
# What is tracked in git and what is not (the policy this script enforces):
#
#   TRACKED    engine + tool source, tests, docs/ (the public site), the
#              measurement SCRIPTS under tools/metrics/, README/CLAUDE.md.
#   IGNORED    everything a measurement RUN produces. Three tiers:
#     (1) root scratch  - renderer PNG/PDF dropped next to the repo root by an
#                         ad-hoc run. Regenerable in seconds. This script
#                         deletes these.
#     (2) tmp/          - per-report probe artifact sets. Regenerable from the
#                         docx corpus + Word, but that costs Word COM time, so
#                         pruning is opt-in (--prune-tmp).
#     (3) pipeline_data/, scratchpad/
#                       - pipeline_data holds the Word ground truth, the
#                         corpora and the frozen benchmark selections: it is
#                         expensive-to-impossible to rebuild and this script
#                         NEVER touches it. scratchpad holds per-session
#                         working files; prune it by hand when you know a
#                         session is finished.
#
# Local-only notes and handoffs (CLAUDE.local.md, docs/spec/, REPORT_*.md,
# INSTRUCTIONS_FOR_GPT.md) are ignored by design - see .gitignore.
set -uo pipefail

cd "$(dirname "$0")/.."
apply=0
prune_tmp=0
for a in "$@"; do
  case "$a" in
    --apply) apply=1 ;;
    --prune-tmp) prune_tmp=1 ;;
    *) echo "unknown option: $a" >&2; exit 2 ;;
  esac
done

# Root-level scratch only: these globs are anchored to the repository root, so
# nothing inside crates/, tools/ or tests/ is ever matched.
patterns=(
  'compare_*.png'      # 3-panel Word/Oxi/diff comparisons
  'x_p*.png' 'p_p*.png'  # renderer <prefix>_p<N>.png output
  'oxi_*.png' 'oxi_*.pdf'
  'word_*.pdf'
  'test_*.pdf' 'test_*.docx'
  'debug_*.txt'
  '*.stackdump'
  '*.tmp'
  '~$*'                # Word lock files
)

found=()
for p in "${patterns[@]}"; do
  for f in $p; do
    [ -f "$f" ] && found+=("$f")
  done
done

if [ ${#found[@]} -eq 0 ]; then
  echo "root scratch: nothing to delete"
else
  bytes=$(du -cb "${found[@]}" 2>/dev/null | tail -1 | cut -f1)
  echo "root scratch: ${#found[@]} files, $(( bytes / 1024 / 1024 )) MB"
  printf '  %s\n' "${found[@]}" | head -8
  [ ${#found[@]} -gt 8 ] && echo "  ... and $(( ${#found[@]} - 8 )) more"
  if [ "$apply" = 1 ]; then
    rm -f -- "${found[@]}"
    echo "  deleted"
  else
    echo "  (dry run - pass --apply to delete)"
  fi
fi

for d in tmp scratchpad pipeline_data; do
  [ -d "$d" ] || continue
  printf '%-14s %s\n' "$d/" "$(du -sh "$d" 2>/dev/null | cut -f1)"
done

if [ "$prune_tmp" = 1 ] && [ -d tmp ]; then
  if [ "$apply" = 1 ]; then
    rm -rf tmp && mkdir -p tmp && echo "tmp/ emptied"
  else
    echo "tmp/ would be emptied (needs --apply too)"
  fi
fi

# Anything left over that git can still see is a real decision, not scratch.
echo
echo "untracked files git would still report:"
git status --porcelain --untracked-files=normal | sed 's/^/  /' | head -20
