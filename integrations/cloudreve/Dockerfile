# Cloudreve with Oxi drawing the thumbnails of Word, PowerPoint and Excel
# files.
#
# Cloudreve makes document thumbnails by running `soffice --headless
# --convert-to png --outdir DIR FILE` (its "LibreOffice" generator, path
# `soffice` by default). This image puts oxidocs first on PATH under that
# name. oxidocs answers that command line for .docx, .pptx and .xlsx, and
# hands every other file (.doc, .odt, .pdf, .txt, ...) to the LibreOffice
# the official image already has, through OXIDOCS_SOFFICE_FALLBACK. Nothing
# in Cloudreve's settings changes.
#
#   docker build -t cloudreve-oxidocs .
#
# The oxidocs release is pinned by version and checksum, as the official
# image pins nothing else from outside Alpine.
ARG CLOUDREVE_VERSION=4.19.1
FROM cloudreve/cloudreve:${CLOUDREVE_VERSION}

ARG OXIDOCS_VERSION=oxidocs-v0.9.0-static.7
ARG OXIDOCS_SHA_AMD64=59f537d81108336dad089eeabc5d15214ba76d0eaa888fdf004764e98640b042
ARG OXIDOCS_SHA_ARM64=fbd266d9624b34c516b7ef70cd7b59570eb089925ed039fb5103887775016c15
ARG TARGETARCH

# oxidocs writes the fonts it carries next to itself on first use, so that is
# done here, once, where /opt/oxidocs is writable.
RUN set -eu; \
    case "${TARGETARCH:-$(uname -m)}" in \
      amd64|x86_64) arch=x86_64; sha="$OXIDOCS_SHA_AMD64" ;; \
      arm64|aarch64) arch=aarch64; sha="$OXIDOCS_SHA_ARM64" ;; \
      *) echo "no oxidocs build for $(uname -m)" >&2; exit 1 ;; \
    esac; \
    asset="oxidocs-${arch}-unknown-linux-musl"; \
    wget -q -O /tmp/oxidocs.tar.gz \
      "https://github.com/Ryujiyasu/oxi/releases/download/${OXIDOCS_VERSION}/${asset}.tar.gz"; \
    echo "${sha}  /tmp/oxidocs.tar.gz" | sha256sum -c -; \
    mkdir -p /opt/oxidocs; \
    tar -xzf /tmp/oxidocs.tar.gz -C /tmp; \
    cp "/tmp/${asset}/"* /opt/oxidocs/; \
    ln -sf /opt/oxidocs/oxidocs /usr/local/bin/soffice; \
    /opt/oxidocs/oxidocs fonts-index > /dev/null; \
    chmod -R a+rX /opt/oxidocs; \
    rm -rf /tmp/oxidocs.tar.gz "/tmp/${asset}"

# What oxidocs does not read goes on to LibreOffice.
ENV OXIDOCS_SOFFICE_FALLBACK=/usr/bin/soffice
