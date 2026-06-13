# Contributing to Oxi

Thank you for your interest in contributing. Oxi has a simple acceptance criterion:

**Every merged PR must improve the pixel accuracy of at least one document.**

See the [Contributing section of the README](README.md#contributing) for what belongs in core vs. an Extension or Fork.

## Developer Certificate of Origin (DCO)

Oxi uses the [Developer Certificate of Origin v1.1](https://developercertificate.org/) instead of a CLA. This means:

- **Your code stays yours.** You keep your copyright. There is no copyright assignment and no license grant beyond the project licenses themselves. The project cannot relicense your contribution out from under you.
- **You certify the origin of your work.** By signing off, you state that you wrote the code (or have the right to submit it) under the license of the files you are touching.

Sign off every commit:

```bash
git commit -s
```

This appends a line to the commit message:

```
Signed-off-by: Your Name <your.email@example.com>
```

The name and email must be real enough to identify you (GitHub-noreply emails are fine). PRs with unsigned commits cannot be merged; fix them with `git rebase --signoff`.

## Licensing of contributions

Oxi is licensed in layers — your contribution is accepted under the license of the files you touch:

| Where | License |
|-------|---------|
| Core engine crates (`oxi-common`, `oxidocs-core`, `oxicells-core`, `oxislides-core`, `oxipdf-core`, `oxihanko`, `oxi-cli`, `oxi-desktop`) | [MPL-2.0](LICENSE) |
| Bindings (`oxi-wasm`, `oxidocs-python`) | MIT OR Apache-2.0 |
| Conformance corpus (self-authored repro documents in `tools/golden-test/repros/`) | CC BY-SA 4.0 |

When adding a **new source file** to a core engine crate, include the MPL-2.0 header at the top:

```rust
// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.
```

MPL-2.0 is file-level copyleft, so the per-file header is what carries the license — please don't omit it.

## Workflow

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Run tests and lint (`cargo test && cargo clippy`)
4. Commit with sign-off (`git commit -s`)
5. Submit a pull request with pixel accuracy results

## Test documents

- Test fixtures go in `tests/fixtures/`
- New low-accuracy test documents are welcome, but they must use openly licensed fonts and you must have the right to publish them under CC BY-SA 4.0 (i.e., author minimal repros yourself; never submit third-party documents you don't own)
