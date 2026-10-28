# Third-party notices

`src/omml.rs` is a Rust port of docling's OMML-to-LaTeX converter, which was
itself adapted from dwml. The port is distributed under this crate's license
(MPL-2.0); the terms below apply to the portions derived from those projects.

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
(full text in `licenses/Apache-2.0.txt`). Modified: docling adapted it in Python;
this crate rewrites that adaptation in Rust.
