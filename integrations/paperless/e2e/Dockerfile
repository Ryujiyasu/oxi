# SPDX-License-Identifier: MIT OR Apache-2.0
# Paperless-ngx with the Oxi .docx parser. For the published plugin, replace
# the COPY/RUN pair with: RUN pip install --no-cache-dir paperless-ngx-oxidocs
ARG PAPERLESS_VERSION=3.4.2
FROM ghcr.io/paperless-ngx/paperless-ngx:${PAPERLESS_VERSION}
COPY dist/ /tmp/plugin/
RUN pip install --no-cache-dir /tmp/plugin/*.whl && rm -rf /tmp/plugin
