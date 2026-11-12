#!/usr/bin/env bash
# Build the App Store tarball: build/oxidocs_preview-<version>.tar.gz, one
# top-level oxidocs_preview/ directory as Nextcloud expects. Needs both
# static builds in bin/ (oxidocs-linux-x86_64 and oxidocs-linux-aarch64).
set -euo pipefail
cd "$(dirname "$0")"
version="$(sed -n 's:.*<version>\(.*\)</version>.*:\1:p' appinfo/info.xml)"
stage=build/oxidocs_preview
rm -rf build && mkdir -p "$stage/bin"
cp -r appinfo lib README.md LICENSE-MIT LICENSE-APACHE "$stage/"
for arch in x86_64 aarch64; do
  install -m 0755 "bin/oxidocs-linux-$arch" "$stage/bin/"
done
# The bundled engine is MPL-2.0; its source is the Oxi repository.
cp ../../LICENSE "$stage/bin/LICENSE-MPL-2.0"
tar -C build --owner=0 --group=0 -czf "build/oxidocs_preview-$version.tar.gz" oxidocs_preview
echo "build/oxidocs_preview-$version.tar.gz"
