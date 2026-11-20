# Third-party notices for the bundled oxidocs binary

The `oxidocs` binary in this package includes oxidocs-docling, whose equation
converter (`crates/oxidocs-docling/src/omml.rs` in the Oxi repository) is a
Rust port of docling's OMML-to-LaTeX converter, itself adapted from dwml. The
port is distributed under MPL-2.0 with the rest of the engine; the terms below
apply to the portions derived from those projects. This file is a copy of
`crates/oxidocs-docling/THIRD_PARTY_NOTICES.md`.

## docling — `docling/backend/docx/latex/omml.py`

https://github.com/docling-project/docling

```
MIT License

Copyright The Docling Contributors

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

## dwml — `dwml/omml.py`

https://github.com/xiilei/dwml — Licensed under the Apache License, Version 2.0
(full text in this package's `LICENSE-APACHE`, and in
`crates/oxidocs-docling/licenses/Apache-2.0.txt` in the Oxi repository). Modified: docling adapted it in Python;
this crate rewrites that adaptation in Rust.
