#!/usr/bin/env bash
# SPDX-License-Identifier: MIT OR Apache-2.0
# End-to-end check: Paperless-ngx 3.4.2 + this plugin consume three .docx files
# (no Tika, no Gotenberg). Prints what Paperless stored for each and exits
# non-zero if a check fails.
#
#   e2e/run.sh            (needs docker, python3 with pip, curl)
set -euo pipefail
HERE=$(cd "$(dirname "$0")" && pwd); ROOT=$(cd "$HERE/.." && pwd)
SAMPLES=$(cd "$ROOT/../gotenberg/e2e/samples" && pwd)
DOCKER=${DOCKER:-docker}
NET=oxidocs-paperless-e2e; PORT=${PORT:-8010}; API="http://localhost:$PORT/api"; AUTH=admin:admin
WORK=$(mktemp -d); mkdir -p "$WORK/consume"; chmod 777 "$WORK/consume"

cleanup() { $DOCKER rm -f oxidocs-paperless oxidocs-broker >/dev/null 2>&1 || true; $DOCKER network rm $NET >/dev/null 2>&1 || true; rm -rf "$WORK"; }
trap cleanup EXIT
cleanup; WORK=$(mktemp -d); mkdir -p "$WORK/consume"; chmod 777 "$WORK/consume"

rm -rf "$HERE/dist"; python3 -m pip wheel --quiet --no-deps -w "$HERE/dist" "$ROOT"
$DOCKER build -q -t paperless-ngx-oxidocs:e2e "$HERE" >/dev/null
$DOCKER network create $NET >/dev/null
$DOCKER run -d --name oxidocs-broker --network $NET redis:8 >/dev/null
$DOCKER run -d --name oxidocs-paperless --network $NET -p "$PORT:8000" \
  -e PAPERLESS_REDIS=redis://oxidocs-broker:6379 \
  -e PAPERLESS_SECRET_KEY="$(python3 -c 'import secrets; print(secrets.token_urlsafe(48))')" \
  -e PAPERLESS_ADMIN_USER=admin -e PAPERLESS_ADMIN_PASSWORD=admin \
  -e PAPERLESS_CONSUMER_POLLING=2 \
  -v "$WORK/consume:/usr/src/paperless/consume" \
  paperless-ngx-oxidocs:e2e >/dev/null

up=0
for _ in $(seq 150); do
  curl -fs -u $AUTH "$API/documents/" >/dev/null 2>&1 && { up=1; break; }
  [ "$($DOCKER inspect -f '{{.State.Running}}' oxidocs-paperless)" = true ] || break
  sleep 2
done
[ $up = 1 ] || { echo "FAIL Paperless did not come up:"; $DOCKER logs oxidocs-paperless 2>&1 | tail -20; exit 1; }
cp "$SAMPLES/simple.docx" "$SAMPLES/lists.docx" "$WORK/consume/"
# unsupported.docx (SmartArt Oxi cannot draw) is a LibreOffice test file whose
# zip order makes libmagic call it application/octet-stream, which Paperless
# refuses before any parser sees it. Re-pack it [Content_Types].xml first, as
# Word writes packages, so it reaches the parser.
python3 - "$SAMPLES/unsupported.docx" "$WORK/consume/unsupported.docx" <<'PY'
import sys, zipfile
src, dst = sys.argv[1:3]
with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst + ".tmp", "w", zipfile.ZIP_DEFLATED) as zout:
    names = sorted(zin.namelist(), key=lambda n: (n != "[Content_Types].xml", n != "_rels/.rels"))
    for n in names:
        zout.writestr(n, zin.read(n))
import os; os.replace(dst + ".tmp", dst)
PY

for _ in $(seq 90); do
  n=$(curl -fs -u $AUTH "$API/documents/" | python3 -c 'import sys,json; print(json.load(sys.stdin)["count"])' 2>/dev/null || echo 0)
  [ "$n" -ge 3 ] && break; sleep 2
done

fails=0
python3 - "$API" "$AUTH" <<'PY' || fails=$((fails+1))
import base64, json, sys, urllib.request
api, auth = sys.argv[1], sys.argv[2]
def get(path, raw=False):
    req = urllib.request.Request(api + path, headers={"Authorization": "Basic " + base64.b64encode(auth.encode()).decode()})
    with urllib.request.urlopen(req) as r:
        return r.read() if raw else json.load(r)
docs = get("/documents/?ordering=original_file_name")["results"]
failed = 0
for d in docs:
    pdf = get(f"/documents/{d['id']}/download/", raw=True)  # the archive (PDF) version
    # Paperless's default output type is PDF/A-2b: ocrmypdf marks it in the XMP.
    pdfa = b"pdfaid:part" in pdf or b"pdfaid:part=" in pdf
    ok = d["page_count"] and d["page_count"] >= 1 and d["content"].strip() and pdf[:5] == b"%PDF-" and pdfa
    failed += not ok
    print(f"{'PASS' if ok else 'FAIL'} {d['original_file_name']:18} pages={d['page_count']} chars={len(d['content'])} archive={len(pdf)} bytes pdfa={pdfa}  {d['content'][:40]!r}")
if len(docs) != 3:
    print(f"FAIL expected 3 documents, got {len(docs)}"); failed += 1
sys.exit(failed)
PY
logs=$($DOCKER logs oxidocs-paperless 2>&1)
grep -q "Loaded third-party parser 'Oxi DOCX Parser'" <<<"$logs" && echo "PASS plugin loaded" || { echo "FAIL plugin not loaded"; fails=$((fails+1)); }
conv=$(grep -c "Oxi converted" <<<"$logs" || true)
[ "$conv" -ge 2 ] && echo "PASS Oxi converted $conv documents" || { echo "FAIL Oxi converted $conv documents"; fails=$((fails+1)); }
grep -q "Tika is off: converting it as well as Oxi can" <<<"$logs" && echo "PASS unsupported.docx: strict refused, converted without strict" || { echo "FAIL no strict fallback seen"; fails=$((fails+1)); }
if [ $fails -gt 0 ]; then echo "--- related log lines"; grep -i -E 'oxi|unsupported\.docx|ParseError|error' <<<"$logs" | grep -v -i 'redis' | tail -15; fi
echo "failed: $fails"; exit $fails
