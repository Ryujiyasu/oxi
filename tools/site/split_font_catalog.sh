#!/bin/sh
# Split the per-face font metric catalog into one file per face for the site.
#
# The site's engine is built without the catalog (`external-font-catalog`) and
# docs/font-catalog.js fetches the faces a document needs as
# font_catalog/<offset>.bin, <offset> being the member's byte offset in
# font_catalog_metrics.gz as the index records it. GitLab Pages does not honour
# Range requests, so the members are published as files of their own.
#
#   tools/site/split_font_catalog.sh <out_dir>
set -eu
out=$1
data=crates/oxidocs-core/src/font/data
mkdir -p "$out"
grep -o '"length":[0-9]*,"offset":[0-9]*' "$data/font_catalog_index.json" |
    sed 's/"length":\([0-9]*\),"offset":\([0-9]*\)/\2 \1/' |
    while read -r offset length; do
        dd if="$data/font_catalog_metrics.gz" of="$out/$offset.bin" bs=1M \
            iflag=skip_bytes,count_bytes skip="$offset" count="$length" status=none
    done
echo "$(ls "$out" | wc -l) catalog members in $out"
