#!/usr/bin/env bash
# Commit snapshot files on top of origin/main and push, without touching any
# working tree or index (other sessions may have uncommitted work in both).
#
#   tools/metrics/snapshot_commit.sh "<subject>" <repo-path>=<source-file> ...
#
#   BASE=<sha>   the commit the snapshots were made from: abort if any of the
#                paths changed on origin/main since then (someone else's change
#                would be overwritten by the snapshot)
#   NAME / EMAIL commit identity (default: git config user.name / user.email);
#                a Signed-off-by trailer is added for it (DCO)
#   COAUTHOR     optional "Name <email>" for a Co-Authored-By trailer
#   NO_PUSH=1    create the commit and print it, do not push
#
# The source files are hashed with --path=<repo-path>, so the repo's eol/filter
# attributes apply as they would to a normal `git add` (hashing with
# --no-filters stores CRLF and turns every line into a change).
set -eu
SUBJ=$1; shift
REPO=$(cd "$(dirname "$0")/../.." && pwd)
cd "$REPO"
NAME=${NAME:-$(git config user.name)}
EMAIL=${EMAIL:-$(git config user.email)}
git fetch -q origin
P=$(git rev-parse origin/main)
IDX=$(mktemp -u)
export GIT_INDEX_FILE=$IDX
trap 'rm -f "$IDX"' EXIT
git read-tree "$P"
for pair in "$@"; do
  f=${pair%%=*}; src=${pair#*=}
  if [ -n "${BASE:-}" ] && [ "$(git rev-parse --verify -q "$BASE:$f" || true)" != "$(git rev-parse --verify -q "$P:$f" || true)" ]; then
    echo "ABORT: $f changed on origin/main since $BASE" >&2; exit 1
  fi
  h=$(git hash-object -w --path="$f" "$src")
  git update-index --add --cacheinfo 100644,"$h","$f"
done
git diff --cached --stat "$P"
T=$(git write-tree)
MSG=$(printf '%s\n\nSigned-off-by: %s <%s>\n' "$SUBJ" "$NAME" "$EMAIL")
[ -n "${COAUTHOR:-}" ] && MSG=$(printf '%s\nCo-Authored-By: %s\n' "$MSG" "$COAUTHOR")
C=$(echo "$MSG" | git -c user.name="$NAME" -c user.email="$EMAIL" commit-tree "$T" -p "$P")
echo "commit $C"
[ -n "${NO_PUSH:-}" ] && exit 0
for i in 1 2 3; do
  git push -q origin "$C":refs/heads/main && { echo pushed; exit 0; }
  sleep 5; git fetch -q origin
done
echo "push failed (origin/main moved?); rebuild the commit on the new origin/main" >&2
exit 1
