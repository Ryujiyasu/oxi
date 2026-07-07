#!/usr/bin/env bash
# Build the publishable npm package (docs-only API surface) into pkg-npm/.
#
# The npm package name is `oxidocs`; the DEFAULT cargo features build only
# the docx bindings (parse/edit/layout/render/docx_to_pdf). The `suite`
# feature (xlsx/pptx/generic-PDF/hanko, used by the in-repo web editor) is
# deliberately NOT part of the published package.
#
# wasm-pack regenerates package.json on every build, so this script re-applies
# the npm metadata (name, keywords, README) afterwards.
set -euo pipefail
cd "$(dirname "$0")"

wasm-pack build --target web --release --out-dir pkg-npm

python - <<'PY'
import json
p = 'pkg-npm/package.json'
d = json.load(open(p, encoding='utf-8'))
d['name'] = 'oxidocs'
d['keywords'] = ['docx', 'word', 'ooxml', 'layout', 'wasm', 'document', 'pdf']
if 'README.md' not in d['files']:
    d['files'].append('README.md')
d['homepage'] = 'https://gitlab.com/Ryujiyasu/oxi'
d['repository'] = {'type': 'git', 'url': 'git+https://gitlab.com/Ryujiyasu/oxi.git'}
json.dump(d, open(p, 'w', encoding='utf-8', newline=''), indent=2, ensure_ascii=False)
print('package.json patched (name=oxidocs)')
PY
cp README-npm.md pkg-npm/README.md
echo "pkg-npm ready: (cd pkg-npm && npm publish)"
