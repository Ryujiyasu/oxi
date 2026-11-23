#!/usr/bin/env bash
# End-to-end check: Cloudreve built from ../Dockerfile makes thumbnails of
# Word, PowerPoint and Excel files with oxidocs, and still makes them for
# other documents with LibreOffice.
#
# Builds the image (Docker), starts it, registers an administrator, uploads
# the fixtures and fetches every thumbnail. LibreOffice is reached through a
# logging wrapper, so the check can tell which files went to it: the three
# Office Open XML files must not, the text file must.
set -euo pipefail
here="$(cd "$(dirname "$0")" && pwd)"
app="$(cd "$here/.." && pwd)"
repo="$(cd "$app/../.." && pwd)"
docker="${DOCKER:-docker}"
name="cloudreve-oxidocs-e2e-$$"
port="${PORT:-15212}"
out="$here/out"
rm -rf "$out" && mkdir -p "$out/fixtures"

cp "$repo/tests/fixtures/comprehensive_test.docx" "$out/fixtures/report.docx"
cp "$repo/tests/fixtures/multi_slide.pptx" "$out/fixtures/slides.pptx"
cp "$repo/tests/fixtures/basic_test.xlsx" "$out/fixtures/sheet.xlsx"
printf 'Plain text goes to LibreOffice.\n' > "$out/fixtures/notes.txt"

cat > "$out/soffice-log" <<'SH'
#!/bin/sh
echo "$*" >> /tmp/libreoffice-calls
exec /usr/bin/soffice "$@"
SH
chmod +x "$out/soffice-log"

cleanup() { [ "${KEEP:-}" = 1 ] && echo "kept container $name" || $docker rm -fv "$name" >/dev/null 2>&1 || true; }
trap cleanup EXIT

$docker build -q -t cloudreve-oxidocs-e2e "$app" >/dev/null
$docker run -d --name "$name" -p "127.0.0.1:$port:5212" \
  -v "$out/soffice-log:/usr/local/lib/soffice-log:ro" \
  -e OXIDOCS_SOFFICE_FALLBACK=/usr/local/lib/soffice-log \
  cloudreve-oxidocs-e2e >/dev/null

for _ in $(seq 60); do
  curl -fsS "http://127.0.0.1:$port/api/v4/site/ping" >/dev/null 2>&1 && break
  sleep 1
done

$docker exec "$name" soffice --version | grep -q "LibreOffice-compatible oxidocs" \
  || { echo "FAIL soffice on PATH is not oxidocs"; exit 1; }

status=0
python3 "$here/check.py" "http://127.0.0.1:$port" "$out/fixtures" "$out" || status=1

calls="$($docker exec "$name" sh -c 'cat /tmp/libreoffice-calls 2>/dev/null' || true)"
if printf '%s\n' "$calls" | grep -Eq '\.(docx|pptx|xlsx)( |$)'; then
  echo "FAIL an Office Open XML file went to LibreOffice:"; printf '%s\n' "$calls"; status=1
fi
printf '%s\n' "$calls" | grep -q 'notes' \
  || { echo "FAIL the text file did not reach LibreOffice"; status=1; }

[ "$status" = 0 ] && echo "PASS thumbnails: docx, pptx, xlsx by oxidocs; txt by LibreOffice"
exit "$status"
