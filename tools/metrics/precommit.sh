#!/usr/bin/env bash
# Pre-commit chain, scheduled so that peak memory stays near 3 GB on this
# 13.9 GB machine (the 2026-10-02 first version ran everything at once and was
# reaped for low memory).
#
#   tools/metrics/precommit.sh <tag> <base-renderer.exe> [OXI_FLAG_DISABLE,...]
#   SKIP_GDI=1  reuse scratch/renderer-release-<tag>.exe instead of building
#   SRC=<dir>   build the renderers and run cargo test from this source tree (a
#               git worktree); the gates and their outputs stay in this repo
#   TGT=<dir>   cargo target root (default: <repo>), one subdirectory per step
#   GOLDEN=1    when flags are given, also run the golden SSIM A/B (ssim_ab.py
#               with the layout-fingerprint prefilter) on the DWrite binary
#
# A. release GDI build (the gates measure this binary)
# B. renderer-only checks side by side: full gate, strict gate, blind-G  (~1 GB)
# C. cargo test, light profile, -j 4                                      (~3 GB)
# D. DWrite release build                                                 (~3 GB)
# E. (GOLDEN=1) golden SSIM A/B with the flags off vs on
# One summary line per step lands in pipeline_data/blindG_20260929/precommit_<tag>.log;
# each step's full output is precommit_<tag>.<step>.log beside it.
set -u
TAG=$1; BASE=$2; FLAGS=${3:-}
REPO=$(cd "$(dirname "$0")/../.." && pwd)
SRC=${SRC:-$REPO}
TGT=${TGT:-$REPO}
cd "$REPO"
OUT=pipeline_data/blindG_20260929
D=$(pwd -W)/$OUT/scratch
SUM=$OUT/precommit_$TAG.log
# strict_gate.py resumes from strict_<tag>.jsonl: a reused tag silently reports
# the OLD binary's results for every case already in that file.
if [ -e "$OUT/strict_${TAG}rel.jsonl" ]; then
  echo "tag $TAG was used before ($OUT/strict_${TAG}rel.jsonl exists); pick a new tag" >&2
  exit 2
fi
: > "$SUM"
t0=$(date +%s)
stamp() { echo "[$(( ($(date +%s) - t0) / 60 ))m] $*" >> "$SUM"; }

if [ -z "${SKIP_GDI:-}" ]; then
  (cd "$SRC/tools/oxi-gdi-renderer" && CARGO_TARGET_DIR="$TGT/target-gdi-pc" cargo build --release -j 8 2>&1 | grep -E "^error|Finished") >> "$SUM"
  cp "$TGT/target-gdi-pc/release/oxi-gdi-renderer.exe" "$D/renderer-release-$TAG.exe" || { stamp "GDI build FAILED"; exit 1; }
fi
stamp "GDI ready"

FLAGARG=""; [ -n "$FLAGS" ] && FLAGARG="--flags $FLAGS"
( python tools/metrics/batch_gate.py True --exe "$D/renderer-release-$TAG.exe" --base "$BASE" $FLAGARG --jobs 3 \
    2>&1 | grep -v "progress\|base=None None  new=None None" > "$OUT/precommit_$TAG.gate.log"; stamp "gate: $(grep 'NEW pass' $OUT/precommit_$TAG.gate.log)" ) &
( python $OUT/strict_gate.py "$D/renderer-release-$TAG.exe" "${TAG}rel" 3 > "$OUT/precommit_$TAG.strict.log" 2>&1; stamp "strict: $(tail -1 $OUT/precommit_$TAG.strict.log)" ) &
( python $OUT/rerun_all.py "$D/renderer-release-$TAG.exe" "${TAG}rel" 2>&1 | grep -E "strict PASS|True->False" > "$OUT/precommit_$TAG.blind.log"; stamp "blind-G: $(tail -1 $OUT/precommit_$TAG.blind.log)" ) &
wait

( cd "$SRC/crates/oxidocs-core" && CARGO_PROFILE_RELEASE_LTO=false CARGO_PROFILE_RELEASE_CODEGEN_UNITS=16 \
    CARGO_TARGET_DIR="$TGT/target-test" cargo test --release -j 4 > "$REPO/$OUT/precommit_$TAG.test.log" 2>&1 )
rc=$?
stamp "cargo test exit=$rc $(grep -E '^test result' $OUT/precommit_$TAG.test.log | awk '{p+=$4; f+=$6; i+=$8} END {print "passed "p" failed "f" ignored "i}')"

(cd "$SRC/tools/oxi-dwrite-renderer" && CARGO_TARGET_DIR="$TGT/target-dwrite-pc" cargo build --release -j 4 2>&1 | grep -E "^error|Finished") > "$OUT/precommit_$TAG.dwrite.log"
cp "$TGT/target-dwrite-pc/release/oxi-dwrite-renderer.exe" "$D/dwrite-release-$TAG.exe" || { stamp "DWrite build FAILED"; exit 1; }
stamp "DWrite: $(tail -1 $OUT/precommit_$TAG.dwrite.log)"

if [ -n "${GOLDEN:-}" ] && [ -n "$FLAGS" ]; then
  OXI_DWRITE_EXE="$D/dwrite-release-$TAG.exe" OXI_GDI_EXE="$D/renderer-release-$TAG.exe" \
    python tools/metrics/ssim_ab.py "$FLAGS" > "$OUT/precommit_$TAG.golden.log" 2>&1
  stamp "golden: $(grep -E '^layout fingerprints|changed bytes|^TOTAL' $OUT/precommit_$TAG.golden.log | tr '\n' ' ')"
fi
stamp "done"
