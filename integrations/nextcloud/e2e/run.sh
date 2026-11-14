#!/usr/bin/env bash
# End-to-end check: Nextcloud with this app shows thumbnails of Word
# documents, drawn by the bundled oxidocs, and without it shows none.
#
# Needs bin/oxidocs-linux-x86_64 (a static build; see the workflow) and
# Docker. NEXTCLOUD_VERSION picks the image (default 35).
set -euo pipefail
here="$(cd "$(dirname "$0")" && pwd)"
app="$(cd "$here/.." && pwd)"
repo="$(cd "$app/../.." && pwd)"
docker="${DOCKER:-docker}"
version="${NEXTCLOUD_VERSION:-35}"
name="oxidocs-preview-e2e-$$"
out="$here/out"
rm -rf "$out" && mkdir -p "$out"
pass=0; fail=0
ok() { echo "PASS $*"; pass=$((pass + 1)); }
bad() { echo "FAIL $*"; fail=$((fail + 1)); }

test -x "$app/bin/oxidocs-linux-x86_64" || { echo "missing $app/bin/oxidocs-linux-x86_64"; exit 2; }

cleanup() {
  if [ "$fail" -gt 0 ] || [ "${KEEP:-}" = 1 ]; then
    $docker exec "$name" sh -c 'tail -n 30 /var/www/html/data/nextcloud.log' 2>/dev/null || true
  fi
  [ "${KEEP:-}" = 1 ] && echo "kept container $name" || $docker rm -fv "$name" >/dev/null 2>&1 || true
}
trap cleanup EXIT

# The app is mounted read-only: it must run oxidocs from a copy and never
# write inside its own directory (a signed app would fail its integrity check).
$docker run -d --name "$name" \
  -e SQLITE_DATABASE=nextcloud \
  -e NEXTCLOUD_ADMIN_USER=admin -e NEXTCLOUD_ADMIN_PASSWORD=e2e-admin-password \
  -e NEXTCLOUD_TRUSTED_DOMAINS=localhost \
  -v "$app:/var/www/html/custom_apps/oxidocs_preview:ro" \
  "nextcloud:${version}-apache" >/dev/null

occ() { $docker exec -u www-data "$name" php occ "$@"; }
for _ in $(seq 1 120); do
  if occ status --output=json 2>/dev/null | grep -q '"installed":true'; then break; fi
  sleep 2
done
occ status --output=json | grep -q '"installed":true' || { bad "Nextcloud did not install"; exit 1; }

req() { $docker exec "$name" curl -s -u admin:e2e-admin-password "$@"; }
# Apache starts only after the install finishes.
for _ in $(seq 1 60); do
  if $docker exec "$name" curl -s -o /dev/null http://localhost/status.php; then break; fi
  sleep 2
done
docs=(comprehensive_test.docx with_image.docx)
for doc in "${docs[@]}"; do
  $docker cp "$repo/tests/fixtures/$doc" "$name:/tmp/$doc"
  code=$(req -o /dev/null -w '%{http_code}' -T "/tmp/$doc" "http://localhost/remote.php/dav/files/admin/$doc")
  [ "$code" = 201 ] || [ "$code" = 204 ] || { bad "upload $doc: HTTP $code"; exit 1; }
done

preview() { # doc, output file → HTTP status
  req -o "/tmp/$2" -w '%{http_code}' "http://localhost/index.php/core/preview.png?file=/$1&x=256&y=256&a=1&forceIcon=0"
}

code=$(preview "${docs[0]}" before.png)
if [ "$code" = 404 ]; then ok "without the app: no preview (HTTP 404)"; else bad "without the app: HTTP $code"; fi

occ app:enable oxidocs_preview >/dev/null
occ app:list --output=json | grep -q '"oxidocs_preview":"0' && ok "app enabled" || bad "app not enabled"
# The web server reads the enabled apps through a cache that can lag the
# occ change by a few seconds.
for _ in $(seq 1 30); do
  if $docker exec -u www-data "$name" php -r 'require "/var/www/html/lib/base.php"; \OC_App::loadApps(); exit(\OCP\Server::get(\OCP\IPreview::class)->isMimeSupported("application/vnd.openxmlformats-officedocument.wordprocessingml.document") ? 0 : 1);' 2>/dev/null; then
    break
  fi
  sleep 1
done
sleep 5

for doc in "${docs[@]}"; do
  start=$(date +%s%N)
  code=$(preview "$doc" "$doc.png")
  ms=$(( ($(date +%s%N) - start) / 1000000 ))
  $docker cp "$name:/tmp/$doc.png" "$out/$doc.png" 2>/dev/null || true
  if [ "$code" != 200 ]; then bad "$doc: HTTP $code"; continue; fi
  if python3 - "$out/$doc.png" <<'PY'
import sys
from PIL import Image
im = Image.open(sys.argv[1]).convert("L")
w, h = im.size
ink = sum(1 for p in im.tobytes() if p < 200) / (w * h)
print(f"  {w}x{h} ink={ink:.3f}")
sys.exit(0 if max(w, h) == 256 and min(w, h) < 256 and ink > 0.005 else 1)
PY
  then ok "$doc: thumbnail in ${ms} ms"; else bad "$doc: not a drawn page"; fi
done

if $docker exec "$name" sh -c 'ls -d /tmp/oxidocs_preview-*/oxidocs' >/dev/null 2>&1; then
  ok "ran from the copy in Nextcloud's temporary directory"
else
  bad "no copy of oxidocs in the temporary directory"
fi
if $docker exec "$name" grep -q '"app":"oxidocs_preview"' /var/www/html/data/nextcloud.log 2>/dev/null; then
  bad "the app logged problems:"
  $docker exec "$name" grep '"app":"oxidocs_preview"' /var/www/html/data/nextcloud.log | tail -n 5
else
  ok "nothing logged by the app"
fi

echo "passed: $pass failed: $fail"
[ "$fail" -eq 0 ]
