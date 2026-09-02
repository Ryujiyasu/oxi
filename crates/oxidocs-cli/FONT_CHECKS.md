# Checking font resolution on a platform

`docx-to-pdf` has to find a font, embed it, and draw with it. Each of those can
fail on its own, and the ways they fail differ by platform — Windows keeps user
and cloud fonts outside the system directory, macOS ships most families as
collections and some as on-demand assets, Linux usually has none of the
Microsoft faces at all. This is how to check a platform we have not checked, or
re-check one after a change.

## Build

```sh
cargo build --release -p oxidocs-cli
cargo test  -p oxipdf-core
cargo test  -p oxidocs-cli
```

Do not pipe `cargo build` into `grep`: the pipeline reports the *filter's* exit
code, so a failed build looks like a success and the next step measures a stale
binary.

## What the machine can see

```sh
./target/release/oxidocs fonts-index                 # everything
./target/release/oxidocs fonts-index "Times New Roman"
```

Each line ends in `#N`, the member index inside the file. `#0` is a plain font
or the first member of a collection; `#1` and beyond only appear where families
ship as collections, which on macOS is most of them.

A family the machine does not have prints `ABSENT`, and that is the correct
answer — it is the cue to substitute or to warn, not a failure.

## The three checks a rendered PDF has to pass

Structure alone is not enough. A PDF can name every font correctly, carry every
font program, and still draw nothing: a subset that kept none of the glyphs it
was asked for embeds a font with no outlines in it. So look at the ink too.

```python
import re, fitz, numpy as np

def check(path):
    data = open(path, 'rb').read()
    type1     = len(re.findall(rb'/Subtype /Type1', data))
    composite = len(re.findall(rb'/Subtype /CIDFontType2', data))
    programs  = len(re.findall(rb'/FontFile2', data))

    ink_rows = 0
    for page in fitz.open(path):
        pix = page.get_pixmap(matrix=fitz.Matrix(150 / 72, 150 / 72), alpha=False)
        rgb = (np.frombuffer(pix.samples, dtype=np.uint8)
                 .reshape(pix.height, pix.width, pix.n)[:, :, :3])
        ink_rows += int((rgb.sum(axis=2) < 600).any(axis=1).sum())

    return {
        'type1': type1,                      # families this machine does not have
        'composite_ok': composite == programs,
        'ink_rows': ink_rows,                # 0 means a blank page
    }
```

**`composite_ok` must hold for every document.** A composite font without its
program writes glyph numbers with nothing in the file to resolve them: not a
different rendering, an undefined one. This is an invariant, and
`every_composite_font_carries_its_font_file` asserts it in the writer's tests.

**`ink_rows` must not be zero** for a page with text on it, and should be in
proportion to the lines that page has. This is the check that catches an
embedded font with no glyphs in it.

**`type1` counts the families the machine does not have**, and each one is named
on stderr as `Warning: no font file for X`. Zero is the goal, but the way to
reach it is to install the family or to add a substitute whose metrics have been
measured to match — not to change the code. A document naming a font nobody has
cannot reach `type1 == 0` without drawing it in something that does not fit,
which is worse than saying so.

## Substitutes

`SUBSTITUTES` in `src/main.rs` maps a family to a free face **only where the
advance widths have been measured to be identical**, so substituting moves no
glyph and the line breaks stay where Word put them. Looking similar is not a
reason: a substitute that merely resembles the family scores worse than letting
the viewer choose, which is measurable and has been measured.

The faces in `fonts/` are the ones we may redistribute. A face that matches but
carries an incompatible licence — Liberation Sans Narrow is GPLv2 with a font
exception — is mapped but not shipped: it is used when the machine already has
it, which costs nothing and redistributes nothing.

## When something is wrong

The stderr lines say which stage failed:

| line | meaning |
|---|---|
| `Substituting X for Y (not installed; advance widths match on ASCII)` | normal — the bundled stand-in was used |
| `Python not available for … subsetting` + `Embedding … whole` | normal — no fontTools, so the whole file was embedded. Larger PDF, same page |
| `Font subset (X) mapped no glyph; embedding the whole font instead` | normal — a symbol font, whose cmap the subsetter cannot read |
| `Warning: no font file for X` | X is not on this machine and has no substitute. Report the family |

An `ink_rows` of zero with none of those lines is the interesting case, and
worth reporting with the document.
