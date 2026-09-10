# -*- coding: utf-8 -*-
"""What version of each engine is on this machine, right now.

A comparison table is only honest if every engine in it was current when it was
measured. That is easy to get wrong quietly: the Rust side is rebuilt every
session and the others are whatever was installed months ago, so the newest Oxi
ends up measured against a competitor's old build and the table flatters us
without anybody deciding to cheat.

So the versions are read off the machine at measurement time and written into
the result, next to the numbers they produced.

    python tools/metrics/engine_versions.py           # print them
    python tools/metrics/engine_versions.py --json    # for a result file
"""
from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
ORACLE = REPO / "tools" / "browser-oracle" / "node_modules"

WINDOWS_APPS = {
    "word": r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",
    "libreoffice": r"C:\Program Files\LibreOffice\program\soffice.exe",
    "onlyoffice": r"C:\Program Files\ONLYOFFICE\DesktopEditors\DesktopEditors.exe",
    "polaris": r"C:\Program Files (x86)\Polaris Office\Office11\Binary\PWord_PC11.exe",
}
NODE_PACKAGES = {
    "silurus": "@silurus/ooxml",
    "eigenpal": "@eigenpal/docx-editor-react",
    "betteroffice": "@betteroffice/docx",
}


def file_version(path: str) -> str | None:
    """The version Windows records inside an executable."""
    if not Path(path).is_file():
        return None
    run = subprocess.run(
        ["powershell", "-NoProfile", "-Command",
         f"(Get-Item '{path}').VersionInfo.ProductVersion"],
        capture_output=True, text=True)
    got = (run.stdout or "").strip()
    return got or None


def node_version(package: str) -> str | None:
    at = ORACLE / package / "package.json"
    if not at.is_file():
        return None
    return json.loads(at.read_text(encoding="utf-8")).get("version")


def collect() -> dict:
    found: dict[str, str | None] = {}
    for name, path in WINDOWS_APPS.items():
        found[name] = file_version(path)
    for name, package in NODE_PACKAGES.items():
        found[name] = node_version(package)

    run = subprocess.run(["npm", "ls", "-g", "@officecli/officecli"],
                         capture_output=True, text=True, shell=True)
    for line in (run.stdout or "").splitlines():
        if "@officecli/officecli@" in line:
            found["officecli"] = line.rsplit("@", 1)[-1].strip()
            break
    found.setdefault("officecli", None)

    gen = REPO / "scratchpad" / "genoffice"
    if gen.is_dir():
        run = subprocess.run(["git", "log", "-1", "--format=%h %ad", "--date=short"],
                             cwd=gen, capture_output=True, text=True)
        found["genoffice"] = (run.stdout or "").strip() or None
    else:
        found["genoffice"] = None

    run = subprocess.run(["git", "rev-parse", "--short", "HEAD"],
                         cwd=REPO, capture_output=True, text=True)
    found["oxi"] = (run.stdout or "").strip() or None
    renderer = REPO / "tools/oxi-dwrite-renderer/target/release/oxi-dwrite-renderer.exe"
    if renderer.is_file():
        import time
        found["oxi_renderer_built"] = time.strftime(
            "%Y-%m-%d %H:%M", time.localtime(renderer.stat().st_mtime))
    return found


def main() -> int:
    found = collect()
    if "--json" in sys.argv:
        print(json.dumps(found, indent=1, ensure_ascii=False))
        return 0
    for name, version in found.items():
        print("%-22s %s" % (name, version if version else "not installed"))
    missing = [n for n, v in found.items() if not v]
    if missing:
        print("\nnot found: " + ", ".join(missing))
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    raise SystemExit(main())
