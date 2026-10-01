#!/usr/bin/env bash
# Pre-commit chain, scheduled so that peak memory stays near 3 GB on this
# 13.9 GB machine (the 2026-10-02 first version ran everything at once and was
# reaped for low memory).
#
#   tools/metrics/precommit.sh <tag> <base-renderer.exe> [OXI_FLAG_DISABLE,...]
#   SKIP_GDI=1  reuse scratch/renderer-release-<tag>.exe instead of building
#
# A. release GDI build (the gates measure this binary)
# B. renderer-only checks side by side: full gate, strict gate, blind-G  (~1 GB)
# C. cargo test, light profile, -j 4                                      (~3 GB)
# D. DWrite release build                                                 (~3 GB)
# One summary line per step lands in pipeline_data/blindG_20260929/precommit_<tag>.log;
# each step's full output is precommit_<tag>.<step>.log beside it.
set -u
TAG=$1; BASE=$2; FLAGS=${3:-}
REPO=$(cd "$(dirname "$0")/../.." && pwd)
cd "$REPO"
OUT=pipeline_data/blindG_20260929
D=$(pwd -W)/$OUT/scratch
SUM=$OUT/precommit_$TAG.log
: > "$SUM"
t0=$(date +%s)
stamp() { echo "[$(( ($(date +%s) - t0) / 60 ))m] $*" >> "$SUM"; }

if [ -z "${SKIP_GDI:-}" ]; then
  (cd tools/oxi-gdi-renderer && cargo build --release -j 8 2>&1 | grep -E "^error|Finished") >> "$SUM" || { stamp "GDI build FAILED"; exit 1; }
  cp tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe "$D/renderer-release-$TAG.exe"
fi
stamp "GDI ready"

FLAGARG=""; [ -n "$FLAGS" ] && FLAGARG="--flags $FLAGS"
( python tools/metrics/batch_gate.py True --exe "$D/renderer-release-$TAG.exe" --base "$BASE" $FLAGARG --jobs 3 \
    2>&1 | grep -v "progress\|base=None None  new=None None" > "$OUT/precommit_$TAG.gate.log"; stamp "gate: $(grep 'NEW pass' $OUT/precommit_$TAG.gate.log)" ) &
( python $OUT/strict_gate.py "$D/renderer-release-$TAG.exe" "${TAG}rel" 3 > "$OUT/precommit_$TAG.strict.log" 2>&1; stamp "strict: $(tail -1 $OUT/precommit_$TAG.strict.log)" ) &
( python $OUT/rerun_all.py "$D/renderer-release-$TAG.exe" "${TAG}rel" 2>&1 | grep -E "strict PASS|True->False" > "$OUT/precommit_$TAG.blind.log"; stamp "blind-G: $(tail -1 $OUT/precommit_$TAG.blind.log)" ) &
wait

( cd crates/oxidocs-core && CARGO_PROFILE_RELEASE_LTO=false CARGO_PROFILE_RELEASE_CODEGEN_UNITS=16 \
    CARGO_TARGET_DIR="$REPO/target-test" cargo test --release -j 4 > "$REPO/$OUT/precommit_$TAG.test.log" 2>&1 )
rc=$?
stamp "cargo test exit=$rc $(grep -E '^test result' $OUT/precommit_$TAG.test.log | awk '{p+=$4; f+=$6; i+=$8} END {print "passed "p" failed "f" ignored "i}')"

(cd tools/oxi-dwrite-renderer && cargo build --release -j 4 2>&1 | grep -E "^error|Finished") > "$OUT/precommit_$TAG.dwrite.log"
stamp "DWrite: $(tail -1 $OUT/precommit_$TAG.dwrite.log)"
stamp "done"
