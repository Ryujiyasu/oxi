# -*- coding: utf-8 -*-
"""Move the giant `impl LayoutEngine` methods of layout/mod.rs into one file each.

Why: rustc puts every item of a module into the same codegen unit, and the
LLVM time-trace of oxidocs-core (2026-10-02) showed cgu.00 (the layout module)
optimising for 70 s and running 310 s on its thread while the other 15 CGUs
took 1-12 s -- the whole crate build waits on that one unit. A method moved to
a child module lands in its own CGU, so the four giants (break_into_lines_with_grid,
layout_table_with_fit_pass, layout_page, layout_paragraph) can be optimised in
parallel. Behaviour must not change: `dump_corpus.py compare` is byte-identical.

    python tools/metrics/split_layout_mod.py          # apply
    python tools/metrics/split_layout_mod.py --dry    # report ranges only
"""
import re, sys
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
MOD = REPO / "crates/oxidocs-core/src/layout/mod.rs"
HEADER = """// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! `LayoutEngine::{name}` -- moved out of `layout/mod.rs` so that it is its own
//! codegen unit (see tools/metrics/split_layout_mod.py). Behaviour-preserving.

use super::*;

impl LayoutEngine {{
"""
MOVES = [  # (method name, new file stem)
    ("break_into_lines_with_grid", "line_break"),
    ("layout_table_with_fit_pass", "table_fit"),
    ("layout_page", "page"),
    ("layout_paragraph", "paragraph"),
]
ITEM = re.compile(r"^    (?:pub(?:\((?:crate|super)\))? )?(?:fn|const|static|type|struct|enum|macro_rules!)\b|^}")


def find_range(lines, name):
    start = next(i for i, l in enumerate(lines) if re.match(rf"^    (?:pub(?:\(\w+\))? )?fn {name}\(", l))
    # pull in the contiguous doc comments / attributes above the signature
    top = start
    while top > 0 and re.match(r"^    (///|#\[)", lines[top - 1]):
        top -= 1
    # the method ends at the last `    }` before the next item at this indent
    nxt = next(i for i in range(start + 1, len(lines)) if ITEM.match(lines[i]))
    end = max(i for i in range(start, nxt) if lines[i] == "    }")
    return top, start, end


def main():
    dry = "--dry" in sys.argv
    src = MOD.read_text(encoding="utf-8")
    lines = src.split("\n")
    ranges = {name: find_range(lines, name) for name, _ in MOVES}
    for name, (top, start, end) in ranges.items():
        print(f"{name}: doc {top + 1}, fn {start + 1}..{end + 1} ({end - start + 1} lines)")
    if dry:
        return
    # apply bottom-up so earlier ranges stay valid
    for name, stem in sorted(MOVES, key=lambda m: -ranges[m[0]][1]):
        top, start, end = ranges[name]
        body = lines[top : end + 1]
        sig = body[start - top]
        body[start - top] = re.sub(r"^    (?:pub(?:\(\w+\))? )?fn ", "    pub(super) fn ", sig)
        out = HEADER.format(name=name) + "\n".join(body) + "\n}\n"
        (MOD.parent / f"{stem}.rs").write_text(out, encoding="utf-8", newline="\n")
        marker = f"    // `{name}` lives in layout/{stem}.rs (its own codegen unit)."
        lines[top : end + 1] = [marker]
        print(f"  -> layout/{stem}.rs ({len(body)} lines)")
    # declare the new modules next to the existing ones
    k = next(i for i, l in enumerate(lines) if l.startswith("mod ruby;"))
    lines[k + 1 : k + 1] = [f"mod {stem};" for _, stem in MOVES]
    MOD.write_text("\n".join(lines), encoding="utf-8", newline="\n")
    print(f"mod.rs now {len(lines)} lines")


if __name__ == "__main__":
    main()
