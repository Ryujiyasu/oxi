# -*- coding: utf-8 -*-
"""Find the construct that makes this engine render a document completely blank.

Two documents in the JA sets lay out with ZERO elements — no text, no borders,
no images — while Word reads 337 and 86 paragraphs out of them. That is not a
pagination error, it is the whole body going missing, and it drags two
documents to score 0.0000 in the Phase-1 gate.

The body is cut down to its first k top-level children and re-rendered until
the output stops being blank. The first k that renders nothing names the child
that kills it, without anyone having to read the document.

    python tools/metrics/_blank_page_bisect.py <docx path>
"""
from __future__ import annotations

import json
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
GDI = REPO / "tools" / "oxi-gdi-renderer" / "target" / "release" / "oxi-gdi-renderer.exe"

# Top-level children of <w:body>. The tag list is deliberately short: anything
# else is left attached to whatever precedes it.
CHILD = re.compile(r"<w:(p|tbl|sectPr|sdt|bookmarkStart|bookmarkEnd|altChunk)[ >]")


def split_body(xml: str) -> tuple[str, list[str], str]:
    start = xml.index("<w:body>") + len("<w:body>")
    end = xml.rindex("</w:body>")
    head, body, tail = xml[:start], xml[start:end], xml[end:]
    # Walk the body counting depth so a nested <w:p> inside a textbox or a cell
    # is not mistaken for a top-level child.
    parts, depth, last = [], 0, 0
    for m in re.finditer(r"<(/?)w:(\w+)([^>]*?)(/?)>", body):
        closing, name, _attrs, selfclose = m.groups()
        if name not in ("p", "tbl", "sectPr", "sdt", "altChunk"):
            continue
        if selfclose:
            continue
        if closing:
            depth -= 1
            if depth == 0:
                parts.append(body[last:m.end()])
                last = m.end()
        else:
            depth += 1
    if last < len(body):
        parts.append(body[last:])
    return head, parts, tail


def render_count(path: Path) -> int:
    with tempfile.TemporaryDirectory() as tmp:
        dump = Path(tmp) / "l.json"
        subprocess.run([str(GDI), str(path), str(Path(tmp) / "p"), "150",
                        f"--dump-layout={dump}"], capture_output=True)
        if not dump.is_file():
            return -1
        data = json.loads(dump.read_text(encoding="utf-8"))
    return sum(len(pg.get("elements", [])) for pg in data.get("pages", []))


def build(src: Path, head: str, parts: list[str], tail: str, k: int, out: Path) -> Path:
    body = "".join(parts[:k])
    # Keep the final sectPr so the page geometry does not change under us.
    if "<w:sectPr" not in body:
        last = next((p for p in reversed(parts) if "<w:sectPr" in p), "")
        body += last
    xml = head + body + tail
    at = out / f"cut{k:04d}.docx"
    shutil.copy(src, at)
    # Rewriting one entry means rebuilding the zip.
    with zipfile.ZipFile(src) as z:
        with zipfile.ZipFile(at, "w", zipfile.ZIP_DEFLATED) as o:
            for item in z.infolist():
                if item.filename == "word/document.xml":
                    o.writestr(item, xml.encode("utf-8"))
                else:
                    o.writestr(item, z.read(item.filename))
    return at


def main() -> int:
    if len(sys.argv) < 2:
        print(f"usage: {Path(sys.argv[0]).name} <docx path>")
        return 2
    src = Path(sys.argv[1])
    xml = zipfile.ZipFile(src).read("word/document.xml").decode("utf-8")
    head, parts, tail = split_body(xml)
    print(f"{src.name}: {len(parts)} top-level body children")
    whole = render_count(src)
    print(f"  whole document renders {whole} elements")

    with tempfile.TemporaryDirectory() as tmp:
        out = Path(tmp)
        # Grow the prefix until it goes blank: the first blank k names the child.
        lo, hi = 1, len(parts)
        first_blank = None
        while lo <= hi:
            mid = (lo + hi) // 2
            n = render_count(build(src, head, parts, tail, mid, out))
            print(f"  first {mid:4} children -> {n:6} elements", flush=True)
            if n == 0:
                first_blank = mid
                hi = mid - 1
            else:
                lo = mid + 1
        if first_blank is None:
            print("  no prefix renders blank — the killer needs the whole body")
            return 0
        print(f"\n  blank from child {first_blank} onward")
        killer = parts[first_blank - 1]
        tags = re.findall(r"<([\w:]+)", killer)
        seen = []
        for t in tags:
            if t not in seen:
                seen.append(t)
        print(f"  that child is {len(killer)} bytes, tags: {' '.join(seen[:24])}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    raise SystemExit(main())
