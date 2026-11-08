# SPDX-License-Identifier: MIT OR Apache-2.0
#
# Gotenberg 9 with .docx converted by Oxi.
#
#   docker build -t gotenberg-oxi .
#   docker run --rm -p 3000:3000 gotenberg-oxi
#
# OXI_COMMIT is the Oxi commit the oxidocs binary is built from (fetched from
# GitHub, so the build context stays this directory). The default is the one
# the 4,988-document run in README.md measured.
#
# The official image, with its gotenberg binary replaced by a build that adds
# the oxi module, and the oxidocs binary next to it. OXI_ENABLE=true is set
# here; everything else is the official image's behaviour.

ARG GOTENBERG_VERSION=9.1.0
ARG OXI_COMMIT=7774ac9cd25f992abf94b6ec9a631d32c99703c2

FROM rust:1.93-trixie AS oxi-stage
ARG OXI_COMMIT
WORKDIR /src
RUN git init -q oxi && cd oxi \
    && git remote add origin https://github.com/Ryujiyasu/oxi.git \
    && git fetch -q --depth 1 origin "$OXI_COMMIT" \
    && git checkout -q FETCH_HEAD
RUN cd oxi && cargo build --release -p oxidocs-cli \
    && cp target/release/oxidocs /oxidocs

FROM golang:1.27.3-trixie AS gotenberg-stage
ARG GOTENBERG_VERSION
WORKDIR /src
COPY go.mod go.sum *.go ./
COPY cmd ./cmd
# Build against the Gotenberg release the final image is based on.
RUN go mod edit -require "github.com/gotenberg/gotenberg/v9@v$GOTENBERG_VERSION" \
    && go mod tidy \
    && CGO_ENABLED=0 go build -o /gotenberg \
       -ldflags "-s -w -X 'github.com/gotenberg/gotenberg/v9/cmd.Version=$GOTENBERG_VERSION+oxi'" \
       ./cmd/gotenberg

FROM gotenberg/gotenberg:${GOTENBERG_VERSION}
COPY --from=gotenberg-stage /gotenberg /usr/bin/gotenberg
COPY --from=oxi-stage /oxidocs /usr/bin/oxidocs
ENV OXI_ENABLE=true
ENV OXI_BIN_PATH=/usr/bin/oxidocs
