#!/usr/bin/env bash
# SPDX-License-Identifier: MIT OR Apache-2.0
# End-to-end check of the gotenberg-oxi image against a running container.
#
#   docker build -t gotenberg-oxi .
#   docker run -d --rm --name gotenberg-oxi -p 3000:3000 gotenberg-oxi
#   e2e/e2e.sh [http://localhost:3000]
#
# Each case prints the engine Gotenberg reports, the reason when LibreOffice
# was used, and whether the expectation held. Exit status is the number of
# failed cases.
set -u
URL=${1:-http://localhost:3000}
HERE=$(cd "$(dirname "$0")" && pwd)
S=$HERE/samples
OUT=$(mktemp -d)
fails=0

for _ in $(seq 60); do curl -fs "$URL/health" >/dev/null && break; sleep 1; done

# check NAME WANT_ENGINE WANT_REASON_PREFIX curl-args...
check() {
  local name=$1 want=$2 reason=$3; shift 3
  local hdr=$OUT/h body=$OUT/b code
  code=$(curl -s -o "$body" -D "$hdr" -w '%{http_code}' "$URL/forms/libreoffice/convert" "$@")
  local engine got_reason residual
  engine=$(grep -i '^gotenberg-engine:' "$hdr" | cut -d' ' -f2- | tr -d '\r')
  got_reason=$(grep -i '^gotenberg-engine-reason:' "$hdr" | cut -d' ' -f2- | tr -d '\r')
  residual=$(grep -i '^gotenberg-remediation-residual:' "$hdr" | cut -d' ' -f2- | tr -d '\r')
  local magic; magic=$(head -c 2 "$body")
  local ok=PASS
  [ "$code" = 200 ] || ok=FAIL
  [ "$engine" = "$want" ] || ok=FAIL
  [[ "$got_reason" == "$reason"* ]] || ok=FAIL
  [ "$magic" = "%P" ] || [ "$magic" = "PK" ] || ok=FAIL  # a PDF, or a zip of PDFs
  [ $ok = PASS ] || fails=$((fails + 1))
  printf '%-4s %-44s http %s  engine=%-11s reason=%s%s\n' "$ok" "$name" "$code" "${engine:-?}" "${got_reason:--}" \
    "${residual:+  residual=$residual}"
}

check "docx -> oxi"                     oxi ""            -F "files=@$S/simple.docx"
check "two docx -> oxi (zip)"           oxi ""            -F "files=@$S/simple.docx" -F "files=@$S/lists.docx"
check "stages merge -> oxi"             oxi ""            -F "files=@$S/simple.docx" -F "files=@$S/lists.docx" -F 'stages=[{"op":"merge"}]'
check "stages remediate PDF/A-2b -> oxi" oxi ""           -F "files=@$S/simple.docx" -F 'stages=[{"op":"remediate","profile":"PDF/A-2b"}]'
check "landscape=true -> libreoffice"   libreoffice "option: landscape" -F "files=@$S/simple.docx" -F landscape=true
check "nativePdfa -> libreoffice"       libreoffice "option: nativePdfa" -F "files=@$S/simple.docx" -F nativePdfa=PDF/A-2b
check "unsupported content -> libreoffice" libreoffice "unsupported-content" -F "files=@$S/unsupported.docx"
check "xlsx -> libreoffice"             libreoffice "no-docx" -F "files=@$S/sheet.xlsx"
check "opt-out header -> libreoffice"   libreoffice "requested" -H 'Gotenberg-Oxi: off' -F "files=@$S/simple.docx"

# An unknown field must still be rejected by the LibreOffice route.
code=$(curl -s -o /dev/null -w '%{http_code}' "$URL/forms/libreoffice/convert" -F "files=@$S/simple.docx" -F merge=true)
if [ "$code" = 400 ]; then echo "PASS unknown field still answers 400"; else echo "FAIL unknown field answered $code"; fails=$((fails + 1)); fi

# Warm timing: the same document 20 times through each engine.
for mode in on off; do
  start=$(date +%s.%N)
  for _ in $(seq 20); do
    curl -s -o /dev/null "$URL/forms/libreoffice/convert" -H "Gotenberg-Oxi: $mode" -F "files=@$S/lists.docx"
  done
  end=$(date +%s.%N)
  awk -v m="$mode" -v a="$start" -v b="$end" 'BEGIN { printf "timing: 20 x lists.docx with Gotenberg-Oxi: %-3s %.2f s\n", m, b - a }'
done

rm -rf "$OUT"
echo "failed: $fails"
exit "$fails"
