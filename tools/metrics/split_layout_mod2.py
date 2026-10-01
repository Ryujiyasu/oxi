# -*- coding: utf-8 -*-
"""Second step of the layout split: make each moved giant a method of its OWN
type so rustc gives it its own codegen unit.

Why: rustc places a method's machine code in the codegen unit of its SELF
TYPE's module, not of the file the `impl` block sits in
(`characteristic_def_id_of_mono_item`). After split_layout_mod.py the four
giants were still all in cgu.00 (LLVM time-trace: 68 s, unchanged) and the
build time did not move (8m49 vs 8m38/9m03). A newtype `XLayouter<'a>(&'a
LayoutEngine)` defined in the child module, with `Deref<Target = LayoutEngine>`
so the body's `self.field` / `self.method()` still resolve, puts the method in
the child module's unit. `Self::assoc` inside the body becomes
`LayoutEngine::assoc`; `layout/mod.rs` keeps a one-line delegating method.

    python tools/metrics/split_layout_mod2.py
"""
import re
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
L = REPO / "crates/oxidocs-core/src/layout"
MOVES = [  # (method, file stem, wrapper type)
    ("break_into_lines_with_grid", "line_break", "LineBreaker"),
    ("layout_table_with_fit_pass", "table_fit", "TableFitLayouter"),
    ("layout_page", "page", "PageLayouter"),
    ("layout_paragraph", "paragraph", "ParagraphLayouter"),
]


def split_params(sig, name):
    """'fn name(&self, a: T, mut b: U<X, Y>) -> R' -> ['a', 'b'] (top-level commas only)."""
    inner = sig[sig.index(name + "(") + len(name) + 1 :]
    inner = re.sub(r"//[^\n]*", "", inner)  # parameter lists carry `// S###` comments
    depth, buf, parts = 0, "", []
    for ch in inner:
        if ch in "([<{":
            depth += 1
        elif ch in ")]>}":
            if depth == 0:
                break
            depth -= 1
        if ch == "," and depth == 0:
            parts.append(buf); buf = ""
        else:
            buf += ch
    if buf.strip():
        parts.append(buf)
    names = []
    for p in parts:
        p = re.sub(r"#\[[^\]]*\]\s*", "", p).strip()  # `#[allow(unused)] x: T`
        if p in ("&self", "&mut self", "self"):
            continue
        m = re.fullmatch(r"(?:mut\s+)?([A-Za-z_]\w*)\s*:.*", p, re.S)
        assert m, f"unsupported parameter pattern: {p[:60]}"
        names.append(m.group(1))
    return names


def main():
    mod = (L / "mod.rs").read_text(encoding="utf-8")
    for name, stem, wrapper in MOVES:
        f = L / f"{stem}.rs"
        s = f.read_text(encoding="utf-8")
        assert "impl LayoutEngine {" in s and "Self::" or True
        # the signature: from `pub(super) fn name(` to the first ` {` that ends it
        i = s.index(f"    pub(super) fn {name}(")
        j = s.index("{\n", i)
        sig = s[i:j].rstrip()
        params = split_params(sig, name)
        ret = re.search(r"\)\s*->\s*(.+)$", sig, re.S)
        body = s[i:]
        body = body.replace("Self::", "LayoutEngine::")
        head = s[:i].replace(
            "impl LayoutEngine {",
            f"/// `{name}` as a method of its own type: rustc puts a method's code in the\n"
            f"/// codegen unit of its self type's module, so this (not the file move alone)\n"
            f"/// is what gives the giant its own unit. Deref keeps `self.x` meaning the engine.\n"
            f"pub(super) struct {wrapper}<'a>(pub(super) &'a LayoutEngine);\n\n"
            f"impl<'a> std::ops::Deref for {wrapper}<'a> {{\n"
            f"    type Target = LayoutEngine;\n"
            f"    fn deref(&self) -> &LayoutEngine {{\n        self.0\n    }}\n}}\n\n"
            f"impl<'a> {wrapper}<'a> {{")
        f.write_text(head + body, encoding="utf-8", newline="\n")
        # delegating method in mod.rs, same signature as the original (private `fn`)
        marker = f"    // `{name}` lives in layout/{stem}.rs (its own codegen unit)."
        assert marker in mod, marker
        orig_sig = sig.replace("    pub(super) fn ", "    fn ", 1)
        stub = (f"{orig_sig} {{\n"
                f"        // body in layout/{stem}.rs (its own codegen unit, see split_layout_mod2.py)\n"
                f"        {stem}::{wrapper}(self).{name}({', '.join(params)})\n"
                f"    }}")
        mod = mod.replace(marker, stub)
        print(f"{name}: {len(params)} params forwarded -> {stem}::{wrapper}")
    (L / "mod.rs").write_text(mod, encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
