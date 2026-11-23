"""Upload files to a running Cloudreve and fetch each one's thumbnail.

    python3 check.py BASE_URL FIXTURE_DIR OUT_DIR

Registers the first user (who becomes the administrator), uploads every file
in FIXTURE_DIR, asks for its thumbnail and writes the image to OUT_DIR.
Prints one line per file, `ok NAME BYTES SECONDS` or `fail NAME WHY`, and
exits 1 if any failed. Standard library only.
"""

import json
import os
import sys
import time
import urllib.error
import urllib.parse
import urllib.request

base, fixtures, out = sys.argv[1].rstrip("/"), sys.argv[2], sys.argv[3]
api = base + "/api/v4"
email, password = "admin@example.com", "oxi-e2e-password"


def call(method, path, body=None, token=None, raw=False, ctype="application/json"):
    data = None
    if body is not None:
        data = body if isinstance(body, bytes) else json.dumps(body).encode()
    req = urllib.request.Request(path if path.startswith("http") else api + path, data=data, method=method)
    if data is not None:
        req.add_header("Content-Type", ctype)
    if token:
        req.add_header("Authorization", "Bearer " + token)
    with urllib.request.urlopen(req, timeout=180) as resp:
        payload = resp.read()
        return (resp.headers, payload) if raw else json.loads(payload)


call("POST", "/user", {"email": email, "password": password})
token = call("POST", "/session/token", {"email": email, "password": password})["data"]["token"]["access_token"]

failed = 0
for name in sorted(os.listdir(fixtures)):
    blob = open(os.path.join(fixtures, name), "rb").read()
    uri = "cloudreve://my/" + urllib.parse.quote(name)
    session = call("PUT", "/file/upload", {"uri": uri, "size": len(blob), "last_modified": int(time.time() * 1000)}, token)
    if session["code"] != 0:
        print(f"fail {name} upload session: {session}")
        failed += 1
        continue
    call("POST", f"/file/upload/{session['data']['session_id']}/0", blob, token, ctype="application/octet-stream")

for name in sorted(os.listdir(fixtures)):
    uri = "cloudreve://my/" + urllib.parse.quote(name)
    start = time.time()
    try:
        thumb = call("GET", "/file/thumb?uri=" + urllib.parse.quote(uri), token=token)
    except urllib.error.HTTPError as e:
        print(f"fail {name} HTTP {e.code}")
        failed += 1
        continue
    if thumb["code"] != 0:
        print(f"fail {name} {thumb.get('msg')}")
        failed += 1
        continue
    url = urllib.parse.urlsplit(thumb["data"]["url"])
    headers, image = call("GET", base + url.path + ("?" + url.query if url.query else ""), raw=True)
    # Served as application/octet-stream; the bytes say what it is.
    kind = next((k for k, magic in (("png", b"\x89PNG"), ("jpeg", b"\xff\xd8\xff"), ("webp", b"RIFF"))
                 if image.startswith(magic)), None)
    if kind is None or len(image) < 1000:
        print(f"fail {name} not an image: {image[:8]!r} {len(image)} bytes")
        failed += 1
        continue
    with open(os.path.join(out, name + ".thumb"), "wb") as f:
        f.write(image)
    print(f"ok {name} {kind} {len(image)} bytes {time.time() - start:.2f}s")

sys.exit(1 if failed else 0)
